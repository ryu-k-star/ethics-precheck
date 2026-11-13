"""Compact case workflow: one internal state, Word, and human questions if needed.

Existing deterministic validators receive temporary derived views. These are never
model-written reports or persistent case deliverables. Old v3 outputs remain readable.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
import json
import os
from pathlib import Path
import subprocess
import shutil
import sys
import tempfile

import precheck_packets as p
from precheck_render import render as render_views
from build_return_docx_from_template import build_docx, sanitize_filename_component
from precheck_verify import verify as verify_views
from docx import Document
import precheck_history as history

BASE_FILES = ("00_source_manifest.json", "00_extracted_pages.json", "00_timeline_facts.json",
              "00_rule_ledger.json", "00_fact_pack.json", "00_r3_hits.json")
ROLES = ("research_ethics", "privacy_consent")
STATE_NAME = "_internal/review_state.json"


def state_path(out):
    return Path(out) / STATE_NAME


def read(out):
    return p.read_json(state_path(out))


@contextmanager
def update_lock(out, expected=None):
    path = state_path(out)
    path.parent.mkdir(parents=True, exist_ok=True)
    lock = path.with_suffix(".lock")
    # Do not silently lose another reviewer's update.
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise ValueError("Another update is in progress; retry recording")
    os.close(fd)
    try:
        if expected == "absent" and path.exists():
            raise ValueError("Output was created concurrently; use another folder")
        if expected not in {None, "absent"} and (not path.exists() or p.object_hash(read(out)) != expected):
            raise ValueError("State changed; reread before updating")
        yield path
    finally:
        lock.unlink(missing_ok=True)


def write_state(path, state):
    history.protect_baseline(p.read_json(path) if path.exists() else {}, state)
    temporary = path.with_suffix(".tmp")
    try:
        p.dump(temporary, state)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def save(out, state, expected=None):
    with update_lock(out, expected) as path:
        write_state(path, state)


def default_template():
    root = Path(__file__).resolve().parents[1]
    packaged = root / "assets/return-template.docx"
    return packaged if packaged.exists() else root / "★倫理申請_修正対応依頼.docx"


def prepare(study, out, cache, rules, force=False):
    study, out = Path(study).resolve(), Path(out).resolve()
    from precheck_scan import is_excluded
    if out.is_relative_to(study) and not is_excluded(out.relative_to(study) / "state.json"):
        raise ValueError("Inside a study, use a 99_プレチェック出力* output folder")
    if state_path(out).exists():
        state = read(out)
        if Path(state["data"]["00_source_manifest.json"]["study"]) != study or Path(state["rules"]) != Path(rules).resolve():
            raise ValueError("Study/rules differ; use a new output folder")
        with materialize(state):
            pass
        return {"reused_existing_output": True, "internal_file": str(state_path(out))}
    if out.exists() and any(out.iterdir()):
        raise ValueError("Existing outputs preserved; choose a new empty output folder")
    with tempfile.TemporaryDirectory(prefix="precheck-prepare-") as tmp:
        command = [sys.executable, str(Path(__file__).with_name("precheck_scan.py")), str(study),
                   "--out", tmp, "--rules", str(Path(rules).resolve()), "--cache", str(Path(cache).resolve()), "--legacy-layout"]
        if force:
            command.append("--force")
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"})
        if result.returncode:
            raise ValueError(result.stderr)
        stage = Path(tmp)
        bundle = p.read_json(stage / "00_packet_manifest.json")
        state = {"schema_version": 1, "layout": "compact", "rules": str(Path(rules).resolve()),
                 "rules_key": bundle["rules_key"], "cache": str(Path(cache).resolve()),
                 "data": {name: p.read_json(stage / name) for name in BASE_FILES},
                 "receipts": {}, "pending": {}, "checks": {}, "title": "",
                 "template": str(default_template().resolve()), "statistics": bundle["statistics"]}
        state["source_data_hash"] = p.object_hash({n: state["data"][n] for n in BASE_FILES if n not in {"00_rule_ledger.json", "00_timeline_facts.json"}})
        state["timeline_evidence_hash"] = p.object_hash(state["data"]["00_timeline_facts.json"]["evidence"])
        save(out, state, "absent")
    return {"layout": "compact", "internal_file": str(state_path(out)), "rules": len(bundle["rule_ids"]), "persistent_case_files": 1}


@contextmanager
def materialize(state, documents=False):
    data = state["data"]
    if state["source_data_hash"] != p.object_hash({n: data[n] for n in BASE_FILES if n not in {"00_rule_ledger.json", "00_timeline_facts.json"}}):
        raise ValueError("Stored source data changed")
    if state["timeline_evidence_hash"] != p.object_hash(data["00_timeline_facts.json"]["evidence"]):
        raise ValueError("Original timeline evidence changed")
    compiled, _ = p.compile_rules(Path(state["rules"]), Path(state["cache"]))
    if compiled["key"] != state["rules_key"]:
        raise ValueError("Rules changed; prepare a new case output")
    with tempfile.TemporaryDirectory(prefix="precheck-view-") as tmp:
        stage = Path(tmp)
        for name, value in data.items():
            p.dump(stage / name, value)
        p.build(stage, Path(state["rules"]), Path(state["cache"]))
        if documents:
            if not state["title"]:
                raise ValueError("Record the study title before drafting")
            render_views(stage, stage, Path(state["template"]), state["title"])
            check = state.get("checks", {}).get("final", {}).get("receipt", {})
            # This view quotes the real factcheck record; it does not assert success.
            notes = [f"{f['finding_id']}: {f.get('status')} / {f.get('note', '')}" for f in check.get("findings", [])]
            (stage / "06_ファクトチェック結果.md").write_text("\n".join(notes), encoding="utf-8")
        yield stage


def snapshot(state, stage, packet_name, word=None):
    names = ["00_rule_ledger.json", "00_timeline_facts.json", "03_サブAI別レビュー結果.md", "04_指摘事項書ドラフト.md", "05_人間確認事項.md"]
    return {"packet_sha256": p.digest(stage / packet_name),
            "inputs": {name: p.digest(stage / name) for name in names},
            "template_sha256": p.digest(Path(state["template"])),
            "word": str(word.resolve()) if word else None, "word_sha256": p.digest(word) if word else None}


def factcheck_packet(state, stage, word=None):
    result = p.later_packet(stage, Path(state["rules"]), Path(state["cache"]), "factcheck", word)
    path = stage / result["packet"]
    # Findings, original evidence and timeline remain. Do not send machine-derived
    # 108-row and 52-row duplicate reports to the model a second time.
    body = path.read_text(encoding="utf-8").split("\n# 03_サブAI別レビュー結果.md", 1)[0]
    body += "\n# 人間確認事項\n" + (stage / "05_人間確認事項.md").read_text(encoding="utf-8")
    if word:
        doc = Document(word)
        body += "\n# 実際のWord（記入欄と下部枠。空欄・全行転記は機械検証）\n"
        body += "\n".join(para.text for para in doc.paragraphs if para.text.strip())
        body += "\n" + "\n".join(" | ".join(c.text for c in row.cells) for row in doc.tables[0].rows if row.cells[1].text.strip())
        body += "\n" + doc.tables[1].cell(0, 0).text
    path.write_text(body, encoding="utf-8")
    meta_path = stage / (result["packet"] + ".manifest.json")
    meta = p.read_json(meta_path)
    meta["packet_sha256"] = p.digest(path)
    p.dump(meta_path, meta)
    receipt_path = stage / (result["packet"] + ".receipt-template.json")
    receipt = p.read_json(receipt_path)
    receipt["packet_sha256"] = meta["packet_sha256"]
    p.dump(receipt_path, receipt)
    return result


def preflight_stage(state, stage):
    """Run the same pre-Word checks used by final verification."""
    return verify_views(argparse.Namespace(
        rules=Path(state["rules"]), timeline=stage / "00_timeline_facts.json",
        ledger=stage / "00_rule_ledger.json", review=stage / "03_サブAI別レビュー結果.md",
        draft=stage / "04_指摘事項書ドラフト.md", factcheck=stage / "06_ファクトチェック結果.md",
        preflight=True))


def require_preflight(state, stage):
    result = preflight_stage(state, stage)
    failures = [f"{c['id']}: {c['detail']}" for c in result["checks"] if not c["passed"]]
    if failures:
        raise ValueError("Preflight failed: " + "; ".join(failures))


def preflight(out):
    state = read(out)
    with materialize(state, documents=True) as stage:
        return preflight_stage(state, stage)


def view(out, role, template=False, ids=None, query=None, word=None):
    state = read(out)
    before = p.object_hash(state)
    if role == "history":
        initial = state.get("ai_initial")
        if initial:
            history.validate_baseline(initial)
        return json.dumps({"baseline": {k: initial[k] for k in ("captured_at", "word_name", "integrity_sha256")} if initial else None,
                           "human_revision_history": state.get("human_revision_history", [])}, ensure_ascii=False, indent=2)
    with materialize(state, documents=role == "factcheck") as stage:
        if role in ROLES:
            name = f"packets/{role}.receipt-template.json" if template else f"packets/{role}.md"
            return (stage / name).read_text(encoding="utf-8")
        if role == "sources":
            value = p.select_sources(stage, ids or [], query) if ids or query else p.read_json(stage / "packets/source_catalog.json")
            return json.dumps(value, ensure_ascii=False, indent=2)
        if role == "ledger":
            ledger = state["data"]["00_rule_ledger.json"]
            return json.dumps({"completed_roles": list(state["receipts"]), "rules": len(ledger["rules"]),
                               "findings": ledger.get("review_findings", []), "dispositions": ledger.get("dispositions", []),
                               "human_notes": ledger.get("human_notes", [])}, ensure_ascii=False, indent=2)
        if role == "timeline":
            return json.dumps(state["data"]["00_timeline_facts.json"], ensure_ascii=False, indent=2)
        if role == "editor":
            result = p.later_packet(stage, Path(state["rules"]), Path(state["cache"]), "editor")
            return (stage / result["packet"]).read_text(encoding="utf-8")
        if role != "factcheck":
            raise ValueError("Unknown view")
        if not word:
            require_preflight(state, stage)
        if word and state.get("word") != str(word.resolve()):
            raise ValueError("Use the Word generated for this output")
        result = factcheck_packet(state, stage, word)
        bound = snapshot(state, stage, result["packet"], word)
        phase = "final" if word else "preliminary"
        state["pending"][phase] = bound
        if state["checks"].get(phase, {}).get("snapshot") != bound:
            state["checks"].pop(phase, None)
        save(out, state, before)
        if template:
            receipt = p.read_json(stage / (result["packet"] + ".receipt-template.json"))
            receipt.update(role="factcheck", phase=phase, snapshot_hash=p.object_hash(bound))
            return json.dumps(receipt, ensure_ascii=False, indent=2)
        return f"phase={phase}\nsnapshot_hash={p.object_hash(bound)}\n" + (stage / result["packet"]).read_text(encoding="utf-8")


def check_receipt(state, stage, receipt, phase):
    word = Path(state["word"]) if phase == "final" else None
    packet = factcheck_packet(state, stage, word)["packet"]
    expected = state["pending"].get(phase)
    if (not expected or expected != snapshot(state, stage, packet, word)
            or receipt.get("packet_sha256") != expected["packet_sha256"]
            or receipt.get("snapshot_hash") != p.object_hash(expected)):
        raise ValueError("Factcheck inputs changed or no bound input was issued")
    primary = {r.get("reviewer") for r in state["receipts"].values()}
    if not receipt.get("reviewer") or receipt["reviewer"] in primary:
        raise ValueError("Factchecker must be independent")
    dispositions = state["data"]["00_rule_ledger.json"].get("dispositions", [])
    findings = receipt.get("findings", [])
    if sorted(f.get("finding_id", "") for f in findings) != sorted(f["finding_id"] for f in dispositions):
        raise ValueError("Factcheck finding coverage differs")
    if any(f.get("status") not in {"confirmed", "inferred", "human"} or not f.get("note", "").strip() for f in findings):
        raise ValueError("Factcheck pending/unexplained")
    fields = ["human_items_checked", "transfer_checked"] + (["visual_word_checked"] if phase == "final" else [])
    if not all(receipt.get(k) is True for k in fields):
        raise ValueError("Factcheck incomplete")
    return expected


def record(out, payload):
    state = read(out)
    before = p.object_hash(state)
    role = payload.get("role")
    with materialize(state, documents=role == "factcheck") as stage:
        if role in ROLES:
            combined = dict(state["receipts"])
            combined[role] = payload
            errors, _ = p.review_errors(stage, list(combined.values()))
            errors = [e for e in errors if e != "missing reviewer"]
            if errors:
                raise ValueError("; ".join(errors))
            state["receipts"][role] = payload
            if set(state["receipts"]) == set(ROLES):
                paths = []
                for name, receipt in state["receipts"].items():
                    path = stage / f"{name}.json"
                    p.dump(path, receipt)
                    paths.append(path)
                p.merge_receipts(stage, paths)
                state["data"]["00_rule_ledger.json"] = p.read_json(stage / "00_rule_ledger.json")
        elif role == "director":
            allowed = {"role", "title", "dispositions", "human_notes", "timeline_review_fields", "template"}
            if set(payload) - allowed:
                raise ValueError("Unknown director fields")
            ledger = state["data"]["00_rule_ledger.json"]
            for key in ("dispositions", "human_notes"):
                if key in payload:
                    ledger[key] = payload[key]
            if "title" in payload:
                state["title"] = payload["title"]
            if "template" in payload:
                state["template"] = str(Path(payload["template"]).resolve())
            if "timeline_review_fields" in payload:
                state["data"]["00_timeline_facts.json"].update(status="reviewed", review_fields=payload["timeline_review_fields"])
        elif role == "factcheck":
            phase = payload.get("phase")
            if phase not in {"preliminary", "final"}:
                raise ValueError("Factcheck phase must be preliminary/final")
            bound = check_receipt(state, stage, payload, phase)
            state["checks"][phase] = {"snapshot": bound, "receipt": payload}
        else:
            raise ValueError("Unknown role")
    state.pop("verification", None)
    save(out, state, before)
    return {"recorded": role, "internal_file": str(state_path(out))}


def render(out, date):
    state = read(out)
    before = p.object_hash(state)
    if len(date) != 4 or not date.isdigit():
        raise ValueError("Provide a four-digit MMDD date")
    with materialize(state, documents=True) as stage:
        require_preflight(state, stage)
        preliminary = state["checks"].get("preliminary")
        if not preliminary:
            raise ValueError("Independent preliminary factcheck required")
        check_receipt(state, stage, preliminary["receipt"], "preliminary")
        word = Path(out).resolve() / f"★倫理申請_修正対応依頼_{sanitize_filename_component(state['title'])}_{date}.docx"
        human = Path(out) / "確認事項.md"
        text = (stage / "05_人間確認事項.md").read_text(encoding="utf-8")
        human_needed = any(line.startswith("- ") for line in text.splitlines())
        with update_lock(out, before) as state_file:
            if word.exists() or human.exists():
                raise ValueError("Existing human documents preserved; reconcile edits before regeneration")
            created = []
            try:
                with tempfile.TemporaryDirectory(prefix="99_プレチェック出力_") as candidate_dir:
                    candidate = Path(candidate_dir) / word.name
                    build_docx(Path(state["template"]), stage / "04_指摘事項書ドラフト.md", candidate)
                    with word.open("xb") as destination:
                        created.append(word)
                        with candidate.open("rb") as source:
                            shutil.copyfileobj(source, destination)
                if human_needed:
                    with human.open("x", encoding="utf-8") as stream:
                        created.append(human)
                        stream.write(text)
                state.update(word=str(word), human_sha256=p.digest(human) if human_needed else None)
                write_state(state_file, state)
            except Exception:
                for path in created:
                    path.unlink(missing_ok=True)
                raise
        return {"word": str(word), "human_questions": str(human) if human_needed else None, "status": "final_factcheck_required"}


def verify(out):
    state = read(out)
    before = p.object_hash(state)
    try:
        word = Path(state["word"])
        final = state["checks"]["final"]
        human = Path(out) / "確認事項.md"
        if (p.digest(human) if human.exists() else None) != state.get("human_sha256"):
            raise ValueError("Human questions changed; reconcile them before verification")
        with materialize(state, documents=True) as stage:
            check_receipt(state, stage, final["receipt"], "final")
            p.dump(stage / "packets/factcheck-with-word.receipt.json", final["receipt"])
            result = verify_views(argparse.Namespace(rules=Path(state["rules"]), timeline=stage / "00_timeline_facts.json",
                ledger=stage / "00_rule_ledger.json", review=stage / "03_サブAI別レビュー結果.md", draft=stage / "04_指摘事項書ドラフト.md",
                factcheck=stage / "06_ファクトチェック結果.md", word=word))
        # Detailed checks stay in the one internal state; no additional report file.
    except (OSError, ValueError, KeyError, TypeError) as exc:
        result = {"passed": False, "errors": [str(exc)]}
    state["verification"] = result
    if result.get("passed"):
        history.seal_baseline(state, word)
    save(out, state, before)
    return result


def compare_human(out, word):
    state = read(out)
    before = p.object_hash(state)
    result = history.compare(state, word)
    current = Path(state["word"])
    if not current.exists() or p.digest(current) != state["ai_initial"]["snapshot"]["word_sha256"]:
        state["verification"] = {"passed": False, "errors": ["Return Word changed after the initial AI baseline; revised Word requires separate verification"]}
    save(out, state, before)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=["view", "record", "preflight", "render", "verify", "compare-human"])
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--role", default="ledger")
    parser.add_argument("--template", action="store_true")
    parser.add_argument("--ids", nargs="*")
    parser.add_argument("--query")
    parser.add_argument("--word", type=Path)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--consume", action="store_true")
    parser.add_argument("--date")
    args = parser.parse_args()
    try:
        if args.command == "view":
            print(view(args.out, args.role, args.template, args.ids, args.query, args.word))
            return 0
        if args.command == "record":
            if not args.input:
                parser.error("--input required")
            if args.consume:
                source = args.input.resolve()
                internal = (args.out / "_internal").resolve()
                if source.parent != internal or source == state_path(args.out).resolve() or source.suffix != ".json":
                    raise ValueError("Only a direct temporary JSON inside _internal may be consumed")
            result = record(args.out, p.read_json(args.input))
            if args.consume:
                args.input.unlink()
        elif args.command == "preflight":
            result = preflight(args.out)
        elif args.command == "render":
            result = render(args.out, args.date or "")
        elif args.command == "compare-human":
            if not args.word:
                parser.error("--word required")
            result = compare_human(args.out, args.word)
        else:
            result = verify(args.out)
        print(json.dumps(result, ensure_ascii=False))
        return 1 if result.get("passed") is False else 0
    except (OSError, ValueError, KeyError, TypeError) as exc:
        parser.exit(2, str(exc) + "\n")


if __name__ == "__main__":
    raise SystemExit(main())
