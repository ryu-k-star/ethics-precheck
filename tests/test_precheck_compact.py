import copy
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import precheck_compact as c
import precheck_packets as p


class CompactTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.study = self.base / "study"
        self.study.mkdir()
        (self.study / "plan.md").write_text("研究課題名：検証\n研究期間 2026年1月1日から2027年3月31日\n同意撤回・情報保管\n", encoding="utf-8")
        self.out = self.study / "99_プレチェック出力_compact"
        self.cache = self.base / "cache"
        c.prepare(self.study, self.out, self.cache, ROOT / "rules")

    def complete_primary(self, human=False, finding=False):
        for role in c.ROLES:
            receipt = json.loads(c.view(self.out, role, template=True))
            receipt["reviewer"] = "synthetic-" + role
            for row in receipt["sources"]:
                row["status"] = "read"
            for row in receipt["rules"]:
                row.update(status="satisfied", note="検証用記録", source_ids=[receipt["sources"][0]["source_id"]])
            if human and role == c.ROLES[0]:
                receipt["rules"][0]["status"] = "human"
            if finding and role == c.ROLES[0]:
                target = next(r for r in receipt["rules"] if r["rule_id"] == "G1-4")
                target["status"] = "candidate"
                receipt["findings"] = [{"finding_id": "A-001", "finding_type": "A", "source_rule_ids": ["G1-4"],
                    "item_numbers": [9, 43, 44], "candidate_text": "試験用候補", "source_ids": target["source_ids"]}]
            c.record(self.out, receipt)
        fields = copy.deepcopy(c.read(self.out)["data"]["00_timeline_facts.json"]["review_fields"])
        fields["finding"] = "検証用の期間評価"
        c.record(self.out, {"role": "director", "title": "合成検証", "dispositions": [], "timeline_review_fields": fields})
        if finding:
            original = c.read(self.out)["receipts"][c.ROLES[0]]["findings"][0]
            c.record(self.out, {"role": "director", "dispositions": [{**original, "document": None,
                "evidence_status": "confirmed", "return_route": "applicant_return", "action_owner": "申請者",
                "included_in_return": True, "disposition_reason": "検証用の採用記録", "root_issue_id": "test",
                "final_text": "区分と対応を整合させてください。"}]})

    def factcheck(self, word=None):
        receipt = json.loads(c.view(self.out, "factcheck", template=True, word=word))
        receipt.update(reviewer="synthetic-factchecker", human_items_checked=True, transfer_checked=True, visual_word_checked=bool(word))
        for f in receipt["findings"]:
            f.update(status="confirmed", note="合成データによる試験記録")
        c.record(self.out, receipt)
        return receipt

    def complete(self, human=False):
        self.complete_primary(human)
        self.factcheck()
        result = c.render(self.out, "1027")
        word = Path(result["word"])
        self.factcheck(word)
        return word

    def test_default_scan_creates_one_internal_file_and_views_create_no_md(self):
        self.assertEqual([x.relative_to(self.out).as_posix() for x in self.out.rglob("*") if x.is_file()], [c.STATE_NAME])
        c.view(self.out, "research_ethics")
        c.view(self.out, "sources")
        c.view(self.out, "ledger")
        self.assertEqual(len([x for x in self.out.rglob("*") if x.is_file()]), 1)
        out2 = self.study / "99_プレチェック出力_cli"
        run = subprocess.run([sys.executable, str(ROOT / "tools/precheck_scan.py"), str(self.study), "--out", str(out2), "--cache", str(self.cache)],
                             capture_output=True, text=True, encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"})
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(json.loads(run.stdout)["layout"], "compact")
        self.assertEqual(len([x for x in out2.rglob("*") if x.is_file()]), 1)

    def test_r3_finding_ids_are_validated_against_actual_rules(self):
        self.complete_primary()
        state = c.read(self.out)
        source_id = state["receipts"][c.ROLES[0]]["sources"][0]["source_id"]
        finding = dict(finding_id="B-001", finding_type="B", document="説明文書",
                       source_rule_ids=["T2-1"], source_ids=[source_id],
                       evidence_status="confirmed", return_route="attachment_return",
                       action_owner="申請者", included_in_return=True,
                       disposition_reason="合成資料の旧名称", root_issue_id="name",
                       final_text="大学名を修正してください。")
        c.record(self.out, {"role": "director", "dispositions": [finding]})
        self.factcheck()
        word = Path(c.render(self.out, "1027")["word"])
        self.factcheck(word)
        self.assertTrue(c.verify(self.out)["passed"])
        finding["source_rule_ids"] = ["T2-999"]
        c.record(self.out, {"role": "director", "dispositions": [finding]})
        self.factcheck(word)
        result = c.verify(self.out)
        self.assertFalse(result["passed"])
        self.assertFalse(next(x for x in result["checks"] if x["id"] == "disposition.valid")["passed"])

    def test_human_revision_preserves_baseline_and_invalidates_current_word_verdict(self):
        word = self.complete()
        self.assertTrue(c.verify(self.out)["passed"])
        initial = copy.deepcopy(c.read(self.out)["ai_initial"])
        doc = Document(word)
        doc.tables[0].cell(0, 1).text = "人間による修正"
        doc.save(word)
        result = c.compare_human(self.out, word)
        self.assertEqual(result["text_changes"], 1)
        saved = c.read(self.out)
        self.assertEqual(initial, saved["ai_initial"])
        self.assertFalse(saved["verification"]["passed"])
        self.assertEqual(len(json.loads(c.view(self.out, "history"))["human_revision_history"]), 1)
        self.assertEqual(len([p for p in self.out.rglob("*") if p.is_file()]), 2)

    def test_full_pipeline_only_word_and_one_internal_file(self):
        word = self.complete()
        result = c.verify(self.out)
        self.assertIn("ai_initial", c.read(self.out))
        self.assertTrue(result["passed"], result)
        self.assertEqual({p.name for p in self.out.iterdir()}, {"_internal", word.name})
        self.assertEqual(len([p for p in self.out.rglob("*") if p.is_file()]), 2)
        self.assertFalse(list(self.out.rglob("*.md")))
        view = c.view(self.out, "factcheck", word=word)
        self.assertNotIn("# 03_サブAI別レビュー結果.md", view)
        self.assertNotIn("# 04_指摘事項書ドラフト.md", view)

    def test_human_questions_only_when_needed_and_edits_are_preserved(self):
        word = self.complete(human=True)
        self.assertTrue(c.verify(self.out)["passed"])
        human = self.out / "確認事項.md"
        self.assertTrue(human.exists())
        self.assertEqual(len(list(self.out.rglob("*.md"))), 1)
        human.write_text(human.read_text(encoding="utf-8") + "人手追記", encoding="utf-8")
        self.assertFalse(c.verify(self.out)["passed"])
        with self.assertRaises(ValueError):
            c.render(self.out, "1027")
        self.assertIn("人手追記", human.read_text(encoding="utf-8"))

    def test_reissued_view_cannot_reuse_old_visual_approval(self):
        word = self.complete()
        old = c.read(self.out)["checks"]["final"]["receipt"]
        doc = Document(word)
        doc.sections[0].top_margin += 10000
        doc.save(word)
        self.assertFalse(c.verify(self.out)["passed"])
        c.view(self.out, "factcheck", word=word)
        with self.assertRaises(ValueError):
            c.record(self.out, old)
        self.assertFalse(c.verify(self.out)["passed"])

    def test_missing_primary_factcheck_or_source_change_stops(self):
        with self.assertRaises(ValueError):
            c.render(self.out, "1027")
        self.complete_primary()
        with self.assertRaises(ValueError):
            c.render(self.out, "1027")
        (self.study / "plan.md").write_text("変更", encoding="utf-8")
        with self.assertRaises(ValueError):
            c.view(self.out, "research_ethics")

    def test_preflight_catches_core_duration_before_factcheck_or_word(self):
        self.complete_primary()
        receipt = copy.deepcopy(c.read(self.out)["receipts"]["research_ethics"])
        row = next(r for r in receipt["rules"] if r["rule_id"] == "G6-2")
        row.update(status="na", note="前向き追跡なし")
        c.record(self.out, receipt)
        result = c.preflight(self.out)
        self.assertFalse(result["passed"])
        self.assertFalse(next(x for x in result["checks"] if x["id"] == "duration.core_applicable")["passed"])
        with self.assertRaisesRegex(ValueError, "duration.core_applicable"):
            c.view(self.out, "factcheck")
        with self.assertRaisesRegex(ValueError, "duration.core_applicable"):
            c.render(self.out, "1027")
        self.assertFalse(list(self.out.glob("*.docx")))
        self.assertFalse((self.out / "確認事項.md").exists())
        row.update(status="human", note="解析完了見込みは未確認")
        c.record(self.out, receipt)
        self.assertTrue(c.preflight(self.out)["passed"])
        self.factcheck()
        word = Path(c.render(self.out, "1027")["word"])
        self.factcheck(word)
        self.assertTrue(c.verify(self.out)["passed"])

    def test_preflight_rejects_return_route_for_excluded_finding(self):
        self.complete_primary(finding=True)
        disposition = copy.deepcopy(c.read(self.out)["data"]["00_rule_ledger.json"]["dispositions"][0])
        disposition.update(included_in_return=False, return_route="applicant_return")
        c.record(self.out, {"role": "director", "dispositions": [disposition]})
        result = c.preflight(self.out)
        self.assertFalse(result["passed"])
        self.assertFalse(next(x for x in result["checks"] if x["id"] == "transfer.non_return_route")["passed"])
        with self.assertRaisesRegex(ValueError, "transfer.non_return_route"):
            c.view(self.out, "factcheck")
        disposition["return_route"] = "human_decision"
        c.record(self.out, {"role": "director", "dispositions": [disposition]})
        self.assertTrue(c.preflight(self.out)["passed"])
        self.factcheck()
        word = Path(c.render(self.out, "1027")["word"])
        self.factcheck(word)
        self.assertTrue(c.verify(self.out)["passed"])

    def test_lock_and_failed_build_leave_no_orphan_word(self):
        self.complete_primary()
        self.factcheck()
        with c.update_lock(self.out):
            with self.assertRaises(ValueError):
                c.render(self.out, "1027")
        self.assertFalse(list(self.out.glob("*.docx")))
        with patch.object(c, "write_state", side_effect=OSError("simulated disk error")):
            with self.assertRaises(OSError):
                c.render(self.out, "1027")
        self.assertFalse(list(self.out.glob("*.docx")))
        self.assertNotIn("word", c.read(self.out))

    def test_findings_transfer_and_omission_gate_remain_active(self):
        self.complete_primary(finding=True)
        self.factcheck()
        word = Path(c.render(self.out, "1027")["word"])
        self.factcheck(word)
        self.assertTrue(c.verify(self.out)["passed"])
        changed = copy.deepcopy(c.read(self.out)["data"]["00_rule_ledger.json"]["dispositions"])
        changed[0]["item_numbers"] = [9]
        c.record(self.out, {"role": "director", "dispositions": changed})
        self.assertFalse(c.verify(self.out)["passed"])

    def test_shared_visual_check_can_be_reused_by_second_reviewer(self):
        doc = Document()
        doc.add_paragraph("研究目的と同意説明、情報の保管")
        doc.save(self.study / "source.docx")
        self.out = self.study / "99_プレチェック出力_visual"
        c.prepare(self.study, self.out, self.cache, ROOT / "rules")
        for i, role in enumerate(c.ROLES):
            receipt = json.loads(c.view(self.out, role, template=True))
            receipt["reviewer"] = "synthetic-" + role
            for row in receipt["sources"]:
                row["status"] = "visual_checked" if i == 0 else "read"
            for row in receipt["rules"]:
                row.update(status="satisfied", source_ids=[receipt["sources"][0]["source_id"]], note="試験")
            c.record(self.out, receipt)
        self.assertEqual(set(c.read(self.out)["receipts"]), set(c.ROLES))


if __name__ == "__main__":
    unittest.main()
