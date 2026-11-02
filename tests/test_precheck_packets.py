from __future__ import annotations

import argparse
import copy
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import precheck_packets as packets
from precheck_render import render
from precheck_verify import verify
from build_return_docx_from_template import build_docx


class PacketTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.study = self.base / "study"
        self.study.mkdir()
        self.rules = self.base / "rules"
        shutil.copytree(ROOT / "rules", self.rules)
        self.out = self.study / "99_プレチェック出力_test"
        self.cache = self.base / "cache"
        self.raw = ("1. 研究課題名 検証用研究\n" + "研究期間 2026年1月1日から2027年3月31日\n" * 90
                    + "同意撤回と情報の保管\n" * 100 + "独自の記載\n" * 100)
        (self.study / "source.md").write_text(self.raw, encoding="utf-8")
        self.scan()

    def scan(self, out=None, rules=None, success=True):
        cmd = [sys.executable, str(ROOT / "tools/precheck_scan.py"), str(self.study), "--legacy-layout", "--out", str(out or self.out),
               "--cache", str(self.cache), "--rules", str(rules or self.rules)]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"})
        if success:
            self.assertEqual(result.returncode, 0, result.stderr)
        else:
            self.assertNotEqual(result.returncode, 0)
        return result

    def receipts(self):
        bundle = packets.load_bundle(self.out)
        result = []
        for role in bundle["packets"]:
            receipt = packets.read_json(self.out / f"packets/{role}.receipt-template.json")
            receipt["reviewer"] = "synthetic-test-" + role
            for source in receipt["sources"]:
                source["status"] = "read"
            for row in receipt["rules"]:
                row.update(status="satisfied", source_ids=[receipt["sources"][0]["source_id"]], note="試験用の確認記録")
            result.append(receipt)
        return result

    def merge(self, receipts=None):
        receipts = receipts or self.receipts()
        paths = []
        for r in receipts:
            path = self.out / f"packets/{r['role']}.receipt.json"
            packets.dump(path, r)
            paths.append(path)
        packets.merge_receipts(self.out, paths)
        return packets.read_json(self.out / "00_rule_ledger.json")

    def full_fixture(self):
        receipts = self.receipts()
        receipts[0]["findings"] = [{"finding_id": "A-001", "finding_type": "A", "source_rule_ids": ["G1-4"],
                                    "item_numbers": [9, 43, 44], "source_ids": [receipts[0]["sources"][0]["source_id"]],
                                    "candidate_text": "区分と対応の整合を確認"}]
        ledger = self.merge(receipts)
        ledger["dispositions"] = [{**receipts[0]["findings"][0], "document": None, "evidence_status": "confirmed",
                                   "return_route": "applicant_return", "action_owner": "申請者", "included_in_return": True,
                                   "disposition_reason": "試験用指摘", "root_issue_id": "test1", "final_text": "区分と対応を整合させてください。"}]
        packets.dump(self.out / "00_rule_ledger.json", ledger)
        timeline = packets.read_json(self.out / "00_timeline_facts.json")
        timeline["status"] = "reviewed"
        timeline["review_fields"]["finding"] = "試験用の期間確認"
        packets.dump(self.out / "00_timeline_facts.json", timeline)
        render(self.out, self.out, ROOT / "★倫理申請_修正対応依頼.docx", "検証用研究")
        (self.out / "06_ファクトチェック結果.md").write_text("A-001を試験fixtureで照合。実研究の評価ではない。", encoding="utf-8")
        word = self.out / "★倫理申請_修正対応依頼_検証用研究_1027.docx"
        build_docx(ROOT / "★倫理申請_修正対応依頼.docx", self.out / "04_指摘事項書ドラフト.md", word)
        packets.later_packet(self.out, self.rules, self.cache, "factcheck", word)
        receipt = packets.read_json(self.out / "packets/factcheck-with-word.md.receipt-template.json")
        receipt.update(reviewer="synthetic-test-factchecker", human_items_checked=True, transfer_checked=True, visual_word_checked=True)
        for f in receipt["findings"]:
            f.update(status="confirmed", note="試験用記録。実際の目視実施を表すものではない。")
        packets.dump(self.out / "packets/factcheck-with-word.receipt.json", receipt)
        return word

    def verify_fixture(self, word):
        return verify(argparse.Namespace(rules=self.rules, timeline=self.out / "00_timeline_facts.json",
             ledger=self.out / "00_rule_ledger.json", review=self.out / "03_サブAI別レビュー結果.md",
             draft=self.out / "04_指摘事項書ドラフト.md", factcheck=self.out / "06_ファクトチェック結果.md", word=word))

    def test_all_raw_text_and_rules_are_assigned_with_exceptions(self):
        bundle = packets.load_bundle(self.out)
        index = packets.read_json(self.out / "00_source_index.json")
        self.assertEqual("".join(row["text"] for row in index.values()), self.raw)
        self.assertEqual(set(index), {sid for p in bundle["packets"].values() for sid in p["source_ids"]})
        rules = [rid for p in bundle["packets"].values() for rid in p["rule_ids"]]
        self.assertEqual(len(rules), 108)
        self.assertEqual(len(rules), len(set(rules)))
        private = (self.out / "packets/privacy_consent.md").read_text(encoding="utf-8")
        self.assertIn("G3-8の施設入力仕様と判定手順", private)
        self.assertNotIn("### 43. 重篤", private)  # R5 drafting examples are not primary-review input.

    def test_cache_reuse_preserves_review_and_invalidates_changed_rule(self):
        self.merge()
        before = (self.out / "00_rule_ledger.json").read_bytes()
        self.assertTrue(json.loads(self.scan().stdout)["reused_existing_output"])
        self.assertEqual(before, (self.out / "00_rule_ledger.json").read_bytes())
        other = self.study / "99_プレチェック出力_second"
        result = json.loads(self.scan(other).stdout)
        self.assertEqual(result["cache_hits"], 1)
        self.assertTrue(result["packets"]["rule_cache_hit"])
        with (self.rules / "R1_項目間整合ルール.md").open("a", encoding="utf-8") as f:
            f.write("\n新しい補足\n")
        self.assertFalse(packets.gate(self.out)["passed"])
        self.scan(success=False)
        compiled, hit = packets.compile_rules(self.rules, self.cache)
        self.assertFalse(hit)
        self.assertNotEqual(compiled["key"], packets.read_json(self.out / "00_packet_manifest.json")["rules_key"])

    def test_new_source_change_and_rule_directory_change_are_rejected(self):
        rules2 = self.base / "rules2"
        shutil.copytree(self.rules, rules2)
        self.scan(rules=rules2, success=False)
        (self.study / "new.txt").write_text("追加資料", encoding="utf-8")
        self.assertFalse(packets.gate(self.out)["passed"])
        self.scan(success=False)

    def test_missing_or_duplicate_rule_assignment_fails(self):
        path = self.rules / "precheck_routes.json"
        routes = packets.read_json(path)
        routes["roles"]["privacy_consent"]["groups"].append("G1")
        packets.dump(path, routes)
        with self.assertRaises(ValueError):
            packets.compile_rules(self.rules, self.cache)

    def test_unread_and_unreferenced_rules_cannot_merge(self):
        receipts = self.receipts()
        receipts[0]["sources"][0]["status"] = "pending"
        errors, _ = packets.review_errors(self.out, receipts)
        self.assertTrue(any("unread" in e for e in errors))
        receipts = self.receipts()
        receipts[0]["rules"][0]["source_ids"] = []
        self.assertTrue(packets.review_errors(self.out, receipts)[0])
        receipts = self.receipts()
        receipts[0]["rules"].pop()
        self.assertTrue(packets.review_errors(self.out, receipts)[0])

    def test_unknown_binary_and_layout_cannot_pass_as_read(self):
        doc = Document()
        doc.add_paragraph("説明文書。図表と選択欄を含む想定。")
        doc.save(self.study / "sheet.docx")
        (self.study / "old.xlsm").write_bytes(b"unsupported")
        self.out = self.study / "99_プレチェック出力_binary"
        self.scan()
        index = packets.read_json(self.out / "00_source_index.json")
        self.assertIn("old.xlsm", {r["source"] for r in index.values()})
        errors, _ = packets.review_errors(self.out, self.receipts())
        self.assertTrue(any("visual check missing" in e for e in errors))

    def test_human_rule_and_missing_finding_scope_fail(self):
        receipts = self.receipts()
        receipts[0]["rules"][0]["status"] = "human"
        self.merge(receipts)
        self.assertTrue(any("human rule" in e for e in packets.gate(self.out)["errors"]))
        receipts = self.receipts()
        receipts[0]["findings"] = [{"finding_id": "A-001", "finding_type": "A", "source_rule_ids": ["G1-4"], "item_numbers": [9, 43, 44]}]
        ledger = self.merge(receipts)
        ledger["dispositions"] = [{**receipts[0]["findings"][0], "item_numbers": [9]}]
        packets.dump(self.out / "00_rule_ledger.json", ledger)
        self.assertTrue(any("scope" in e for e in packets.gate(self.out)["errors"]))

    def test_full_pipeline_and_stale_final_evidence(self):
        word = self.full_fixture()
        result = self.verify_fixture(word)
        self.assertTrue(result["passed"], result)
        with self.assertRaises(ValueError):
            render(self.out, self.out, ROOT / "★倫理申請_修正対応依頼.docx", "検証用研究")
        timeline = self.out / "00_timeline_facts.json"
        data = packets.read_json(timeline)
        data["review_fields"]["finding"] = "変更後"
        packets.dump(timeline, data)
        self.assertFalse(packets.gate(self.out, word)["passed"])

    def test_target_drop_from_03_and_04_is_detected(self):
        word = self.full_fixture()
        review = self.out / "03_サブAI別レビュー結果.md"
        review.write_text(review.read_text(encoding="utf-8").replace("9,43,44", "9"), encoding="utf-8")
        draft = self.out / "04_指摘事項書ドラフト.md"
        text = draft.read_text(encoding="utf-8")
        text = "\n".join(line.replace("要対応", "対応不要").replace("区分と対応を整合させてください。", "") if line.startswith(("| 43.", "| 44.")) else line for line in text.splitlines())
        draft.write_text(text, encoding="utf-8")
        result = self.verify_fixture(word)
        self.assertFalse(next(c for c in result["checks"] if c["id"] == "transfer.findings_to_04")["passed"])

    def test_final_visual_and_independent_confirmation_required(self):
        word = self.full_fixture()
        path = self.out / "packets/factcheck-with-word.receipt.json"
        receipt = packets.read_json(path)
        receipt["visual_word_checked"] = False
        packets.dump(path, receipt)
        self.assertFalse(packets.gate(self.out, word)["passed"])
        receipt["visual_word_checked"] = True
        receipt["reviewer"] = self.receipts()[0]["reviewer"]
        packets.dump(path, receipt)
        self.assertFalse(packets.gate(self.out, word)["passed"])

    def test_get_does_not_include_other_source_and_detects_tampering(self):
        index = packets.read_json(self.out / "00_source_index.json")
        sid = next(iter(index))
        result = packets.select_sources(self.out, [sid])
        self.assertEqual([r["source_id"] for r in result], [sid])
        index[sid]["text"] = "tampered"
        packets.dump(self.out / "00_source_index.json", index)
        self.assertFalse(packets.gate(self.out)["passed"])


if __name__ == "__main__":
    unittest.main()
