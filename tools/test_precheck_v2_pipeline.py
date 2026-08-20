from __future__ import annotations

import argparse
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "codex-skill"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class PrecheckV2PipelineTests(unittest.TestCase):
    def test_scan_reuses_hash_cache_and_excludes_prior_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            study = base / "study-001"
            study.mkdir()
            (study / "plan.txt").write_text(
                "研究期間 2026年8月1日から2027年3月31日まで\n"
                "予定対象者数 38名\n統計解析期間は未記載\n",
                encoding="utf-8",
            )
            excluded = study / "99_プレチェック出力_old"
            excluded.mkdir()
            Document().save(excluded / "old.docx")
            output = study / "99_プレチェック出力_v2"
            cache = base / "cache"
            command = [
                sys.executable,
                str(SKILL / "scripts" / "precheck_scan.py"),
                str(study),
                "--out",
                str(output),
                "--cache",
                str(cache),
                "--rules",
                str(SKILL / "references" / "rules"),
            ]

            first = subprocess.run(command, check=True, capture_output=True, text=True)
            second = subprocess.run(command, check=True, capture_output=True, text=True)
            first_summary = json.loads(first.stdout)
            second_summary = json.loads(second.stdout)
            ledger = json.loads((output / "00_rule_ledger.json").read_text(encoding="utf-8"))
            manifest = json.loads((output / "00_source_manifest.json").read_text(encoding="utf-8"))

            self.assertEqual(first_summary["source_count"], 1)
            self.assertEqual(first_summary["cache_hits"], 0)
            self.assertEqual(second_summary["cache_hits"], 1)
            self.assertEqual(len(ledger["rules"]), 108)
            self.assertTrue(any("old.docx" in warning for warning in manifest["warnings"]))

    def test_verifier_reads_composite_rule_rows_and_skill_rule_names(self) -> None:
        verifier = load_module("precheck_verify", SKILL / "scripts" / "precheck_verify.py")
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            ledger = {
                "allowed_evidence_statuses": ["confirmed", "inferred", "human"],
                "allowed_return_routes": [
                    "applicant_return",
                    "attachment_return",
                    "researcher_profile",
                    "secretariat_internal",
                    "human_decision",
                    "resolved_elsewhere",
                ],
                "rules": [
                    {"rule_id": rule_id, "status": "satisfied"}
                    for rule_id in verifier.expected_rule_ids(SKILL / "references" / "rules")
                ],
                "dispositions": [],
            }
            ledger_path = base / "ledger.json"
            ledger_path.write_text(json.dumps(ledger), encoding="utf-8")
            review = base / "03.md"
            review.write_text(
                "## A欄候補\n\n"
                "| ルールID | 指摘先項目 | 根拠頁 | 内容 |\n"
                "|---|---|---|---|\n"
                "| G6-1,G6-2,G6-7 | 17 | 4,9 | 期間を確認 |\n\n"
                "## B欄候補\n",
                encoding="utf-8",
            )
            draft = base / "04.md"
            rows = ["| 項目 | 対応要否 | 指摘内容 |", "|---|---|---|"]
            for item in range(1, 52):
                status = "要対応" if item == 17 else "対応不要"
                comment = "登録・追跡・解析を含む期間を整理してください。" if item == 17 else ""
                rows.append(f"| {item}. item | {status} | {comment} |")
            rows.append("| 連絡担当者 | 対応不要 |  |")
            draft.write_text("\n".join(rows), encoding="utf-8")
            factcheck = base / "06.md"
            factcheck.write_text("確認済み", encoding="utf-8")

            document = Document()
            main = document.add_table(rows=52, cols=2)
            for index, row in enumerate(main.rows):
                label = f"{index + 1}. item" if index < 51 else "連絡担当者"
                row.cells[0].text = label
                if index == 16:
                    row.cells[1].text = "登録・追跡・解析を含む期間を整理してください。"
            document.add_table(rows=1, cols=1)
            word = base / "★倫理申請_修正対応依頼_テスト研究_0820.docx"
            document.save(word)

            result = verifier.verify(
                argparse.Namespace(
                    rules=SKILL / "references" / "rules",
                    ledger=ledger_path,
                    review=review,
                    draft=draft,
                    factcheck=factcheck,
                    word=word,
                )
            )
            transfer = next(check for check in result["checks"] if check["id"] == "transfer.03_to_04")
            self.assertTrue(result["passed"])
            self.assertIn("X=[17]", transfer["detail"])


if __name__ == "__main__":
    unittest.main()
