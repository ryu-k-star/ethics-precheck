"""Synchronize maintained local sources to package/installed mirrors, with backups."""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAPPING = {
    "rules/README.md": "references/rules/rules-index.md",
    "rules/R1_項目間整合ルール.md": "references/rules/r1-cross-document-consistency.md",
    "rules/R2_必須記載チェックリスト.md": "references/rules/r2-required-content.md",
    "rules/R3_書式残骸・表記ルール.md": "references/rules/r3-format-and-word-qa.md",
    "rules/R4_文体規範と文例集.md": "references/rules/r4-writing-style.md",
    "rules/R5_事務局定型文.md": "references/rules/r5-office-standard-text.md",
    "rules/抽出根拠.md": "references/rules/rule-provenance.md",
    "rules/逆適用検証_20260625-290.md": "references/rules/reverse-validation.md",
    "rules/precheck_routes.json": "references/rules/precheck_routes.json",
    "docs/precheck_v3_workflow.md": "references/precheck_v3_workflow.md",
    "docs/precheck_v2_workflow.md": "references/precheck_v2_workflow.md",
    "docs/review_quality_playbook.md": "references/review_quality_playbook.md",
    "docs/review_operating_contract.md": "references/workflow.md",
    "examples/gold_reviews/README.md": "references/examples/review_examples_policy.md",
    "★倫理申請_修正対応依頼.docx": "assets/return-template.docx",
}
for name in ("precheck_scan.py", "precheck_verify.py", "precheck_packets.py", "precheck_render.py", "precheck_compact.py", "precheck_history.py", "precheck_usage.py", "build_return_docx_from_template.py", "extract_research_docs.py"):
    MAPPING["tools/" + name] = "scripts/" + name


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else None


def synchronize(install=None, export=None, check=False):
    package = ROOT / "codex-skill"
    pairs = [(ROOT / src, package / dst) for src, dst in MAPPING.items()]
    # Resolve package-generated mappings from their original source even in check mode.
    package_sources = {dst: ROOT / src for src, dst in MAPPING.items()}
    for path in package.rglob("*"):
        if path.is_file() and "__pycache__" not in path.parts:
            package_sources.setdefault(path.relative_to(package).as_posix(), path)
    if install:
        pairs += [(source, install / dst) for dst, source in package_sources.items()]
    if export:
        pairs += [(ROOT / src, export / src) for src in MAPPING]
        pairs += [(source, export / "codex-skill" / dst) for dst, source in package_sources.items()]
        pairs += [(ROOT / n, export / n) for n in ("AGENTS.md", "tools/sync_precheck_skill.py")]
    changes = []
    backup = ROOT / "outputs" / ("skill-sync-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))
    for src, dst in pairs:
        if not src.is_file():
            raise FileNotFoundError(src)
        before, after = sha(dst), sha(src)
        if before == after:
            continue
        entry = {"source": str(src), "destination": str(dst), "before_sha256": before, "after_sha256": after}
        changes.append(entry)
        if not check:
            if dst.exists():
                saved = backup / f"{len(changes):03d}" / dst.name
                saved.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(dst, saved)
                entry["backup"] = str(saved)
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
    report = {"check_only": check, "differences": len(changes), "changes": changes}
    if not check:
        backup.mkdir(parents=True, exist_ok=True)
        (backup / "sync-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--install", type=Path)
    p.add_argument("--export", type=Path)
    p.add_argument("--check", action="store_true")
    a = p.parse_args()
    r = synchronize(a.install, a.export, a.check)
    print(json.dumps(r, ensure_ascii=False))
    raise SystemExit(1 if a.check and r["differences"] else 0)
