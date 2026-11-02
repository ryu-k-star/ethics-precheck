"""Preserve the first verified AI text and compare later Word files locally."""
from __future__ import annotations

from datetime import datetime, timezone
from hashlib import sha256
import json
import re
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def digest_object(value):
    return sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def text(node):
    # Retain tracked additions/deletions visibly instead of silently losing them.
    if node.tag in {W + "ins", W + "del"}:
        kind = "追加" if node.tag == W + "ins" else "削除"
        return "【変更履歴:" + kind + ":" + "".join(text(c) for c in node) + "】"
    if node.tag in {W + "t", W + "delText"}:
        return node.text or ""
    if node.tag == W + "tab":
        return "\t"
    if node.tag in {W + "br", W + "cr"}:
        return "\n"
    return "".join(text(c) for c in node)


def snapshot(path):
    path = Path(path)
    data = path.read_bytes()
    cells = {}
    shapes = {}
    tracked = 0
    with ZipFile(path) as archive:
        names = sorted(n for n in archive.namelist() if not n.endswith("/"))
        hashes = {n: sha256(archive.read(n)).hexdigest() for n in names}
        parts = [n for n in names if n.endswith(".xml") and (n == "word/document.xml" or n.startswith(("word/header", "word/footer", "word/footnotes", "word/endnotes")))]
        def walk(node, location):
            nonlocal tracked
            if node.tag == W + "p":
                cells[location] = text(node)
                tracked += sum(x.tag in {W + "ins", W + "del"} for x in node.iter())
                return
            if node.tag == W + "tbl":
                rows = node.findall(W + "tr")
                shapes[location] = [len(r.findall(W + "tc")) for r in rows]
            counts = {}
            for child in node:
                name = child.tag.rsplit("}", 1)[-1]
                number = counts.get(name, 0)
                counts[name] = number + 1
                walk(child, f"{location}/{name}[{number}]")
        for part in parts:
            walk(ET.fromstring(archive.read(part)), part)
    return {"word_sha256": sha256(data).hexdigest(), "text": cells,
            "table_shapes": shapes, "tracked_changes": tracked, "package_hashes": hashes, "package_hash_scope": "all_parts"}


def seal_baseline(state, word):
    if state.get("ai_initial"):
        validate_baseline(state["ai_initial"])
        return
    baseline = {"schema_version": 1, "captured_at": datetime.now(timezone.utc).isoformat(),
                "word_name": Path(word).name, "capture_stage": "first_verified_return",
                "snapshot": snapshot(word),
                "verification": state.get("verification", {}),
                "finding_links": [{k: d.get(k) for k in ("finding_id", "finding_type", "item_numbers", "document", "source_rule_ids", "source_ids", "included_in_return", "disposition_reason")}
                                  for d in state["data"]["00_rule_ledger.json"].get("dispositions", [])]}
    baseline["integrity_sha256"] = digest_object(baseline)
    state["ai_initial"] = baseline


def validate_baseline(baseline):
    payload = {k: v for k, v in baseline.items() if k != "integrity_sha256"}
    if baseline.get("integrity_sha256") != digest_object(payload):
        raise ValueError("AI baseline integrity check failed; do not replace it with a human-edited Word")


def protect_baseline(previous, updated):
    old = previous.get("ai_initial")
    new = updated.get("ai_initial")
    if old:
        validate_baseline(old)
        if new != old:
            raise ValueError("The initial AI baseline cannot be replaced or removed")
    if new:
        validate_baseline(new)


def compare(state, word):
    baseline = state.get("ai_initial")
    if not baseline:
        raise ValueError("No verified AI baseline exists. Do not infer it from an edited Word; recover the verified initial version first")
    validate_baseline(baseline)
    before = baseline["snapshot"]
    after = snapshot(word)
    history = state.setdefault("human_revision_history", [])
    for entry in history:
        if entry["word_sha256"] == after["word_sha256"]:
            return {"revision_id": entry["revision_id"], "reused": True, "text_changes": len(entry["changes"])}
    if before["word_sha256"] == after["word_sha256"]:
        return {"unchanged": True, "text_changes": 0}
    changes = []
    for key in sorted(set(before["text"]) | set(after["text"])):
        left, right = before["text"].get(key), after["text"].get(key)
        if left == right:
            continue
        kind = "added" if not left else "deleted" if not right else "modified"
        row = re.match(r"^(word/document.xml/body\[0\]/tbl\[0\]/tr\[\d+\])/", key)
        label = "本文・見出し等"
        if row:
            prefix = row[1] + "/tc[0]/"
            label = " ".join(v for k, v in before["text"].items() if k.startswith(prefix)).strip()
            if not label:
                label = " ".join(v for k, v in after["text"].items() if k.startswith(prefix)).strip()
        elif "word/document.xml/body[0]/tbl[1]/" in key:
            label = "B・C欄等の下部枠"
        changes.append({"location": key, "item_label": label, "kind": kind, "before": left, "after": right,
                        "reason_category": None, "reason": None, "reviewed_for_learning": False})
    # Do not retrofit or replace the immutable baseline captured by older versions.
    legacy_scope = before.get("package_hash_scope") != "all_parts"
    after_hashes = after["package_hashes"]
    if legacy_scope:
        after_hashes = {k: v for k, v in after_hashes.items() if k.startswith("word/") and k.endswith(".xml")}
    package_changes = sorted(k for k in set(before["package_hashes"]) | set(after_hashes)
                             if before["package_hashes"].get(k) != after_hashes.get(k))
    unclassified_package_change = legacy_scope and not changes and not package_changes
    if not changes and not package_changes and not unclassified_package_change:
        return {"unchanged": True, "text_changes": 0}
    entry = {"revision_id": f"human-{len(history) + 1:04d}", "recorded_at": datetime.now(timezone.utc).isoformat(),
             "word_name": Path(word).name, "word_sha256": after["word_sha256"],
             "baseline_sha256": baseline["integrity_sha256"], "changes": changes,
             "changed_package_parts": package_changes, "table_shapes": after["table_shapes"],
             "unclassified_package_change": unclassified_package_change,
             "manual_mapping_required": (after["table_shapes"] != before["table_shapes"]
                                         or set(after["text"]) != set(before["text"])),
             "tracked_changes": after["tracked_changes"],
             "authorship": "user_supplied_revision_not_independently_verified",
             "revised_word_verified": False,
             "reviewed_for_learning": False}
    history.append(entry)
    return {"revision_id": entry["revision_id"], "text_changes": len(changes),
            "manual_mapping_required": entry["manual_mapping_required"],
            "tracked_changes": entry["tracked_changes"], "reviewed_for_learning": False}
