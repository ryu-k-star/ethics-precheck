import copy
from zipfile import ZipFile
from pathlib import Path
import sys
import tempfile
import unittest
from docx import Document
from docx.oxml import OxmlElement

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import precheck_history as h


class HistoryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.original = Path(self.tmp.name) / "original.docx"
        self.edited = Path(self.tmp.name) / "edited.docx"
        doc = Document()
        doc.add_paragraph("研究名")
        table = doc.add_table(rows=52, cols=2)
        for i, row in enumerate(table.rows):
            row.cells[0].text = str(i + 1)
        table.cell(0, 1).text = "修正してください。"
        table.cell(2, 1).text = "削除対象の依頼。"
        doc.add_table(rows=1, cols=1).cell(0, 0).text = "添付資料を再提出してください。"
        doc.save(self.original)
        self.state = {"data": {"00_rule_ledger.json": {"dispositions": []}}}
        h.seal_baseline(self.state, self.original)

    def test_changes_preserve_original_and_deduplicate(self):
        baseline = copy.deepcopy(self.state["ai_initial"])
        doc = Document(self.original)
        doc.tables[0].cell(0, 1).text = "確認してください。"
        doc.tables[0].cell(1, 1).text = "新しい指摘。"
        doc.tables[0].cell(2, 1).text = ""
        doc.tables[1].cell(0, 0).text = "説明文書を再提出してください。"
        doc.save(self.edited)
        result = h.compare(self.state, self.edited)
        self.assertEqual(result["text_changes"], 4)
        self.assertEqual({x["item_label"] for x in self.state["human_revision_history"][0]["changes"]}, {"1", "2", "3", "B・C欄等の下部枠"})
        self.assertEqual({x["kind"] for x in self.state["human_revision_history"][0]["changes"]}, {"added", "deleted", "modified"})
        self.assertEqual(self.state["ai_initial"], baseline)
        self.assertTrue(h.compare(self.state, self.edited)["reused"])
        self.assertEqual(len(self.state["human_revision_history"]), 1)
        self.assertFalse(self.state["human_revision_history"][0]["reviewed_for_learning"])

    def test_non_xml_package_changes_are_recorded_without_text_edits(self):
        # Images and relationship targets can change the rendered document without
        # changing its text or Word XML. Preserve both as human revisions.
        for part in ("word/media/image1.png", "word/_rels/document.xml.rels"):
            with self.subTest(part=part):
                state = copy.deepcopy(self.state)
                with ZipFile(self.original) as source, ZipFile(self.edited, "w") as target:
                    for entry in source.infolist():
                        data = source.read(entry.filename)
                        if entry.filename == part:
                            data += b"\n"
                        target.writestr(entry, data)
                    if part not in source.namelist():
                        target.writestr(part, b"synthetic image content")
                result = h.compare(state, self.edited)
                self.assertFalse(result.get("unchanged", False))
                self.assertEqual(result["text_changes"], 0)
                revision = state["human_revision_history"][0]
                self.assertIn(part, revision["changed_package_parts"])
                self.assertFalse(revision["revised_word_verified"])
                self.assertEqual(state["ai_initial"], self.state["ai_initial"])
                self.assertTrue(h.compare(state, self.edited)["reused"])

    def test_legacy_baseline_does_not_hide_unparsed_package_changes(self):
        state = copy.deepcopy(self.state)
        baseline = state["ai_initial"]
        snap = baseline["snapshot"]
        snap.pop("package_hash_scope")
        snap["package_hashes"] = {k: v for k, v in snap["package_hashes"].items()
                                  if k.startswith("word/") and k.endswith(".xml")}
        baseline.pop("integrity_sha256")
        baseline["integrity_sha256"] = h.digest_object(baseline)
        original_baseline = copy.deepcopy(baseline)
        self.assertTrue(h.compare(state, self.original)["unchanged"])
        with ZipFile(self.original) as source, ZipFile(self.edited, "w") as target:
            for entry in source.infolist():
                target.writestr(entry, source.read(entry.filename))
            target.writestr("word/media/image1.png", b"synthetic image content")
        self.assertFalse(h.compare(state, self.edited).get("unchanged", False))
        self.assertTrue(state["human_revision_history"][0]["unclassified_package_change"])
        self.assertEqual(state["ai_initial"], original_baseline)

    def test_baseline_cannot_be_removed_replaced_or_corrupted(self):
        for edit in (lambda s: s.pop("ai_initial"), lambda s: s["ai_initial"]["snapshot"].update(word_sha256="bad")):
            changed = copy.deepcopy(self.state)
            edit(changed)
            with self.assertRaises(ValueError):
                h.protect_baseline(self.state, changed)
        with self.assertRaises(ValueError):
            h.compare({}, self.original)

    def test_structure_and_tracked_edits_require_attention(self):
        doc = Document(self.original)
        doc.tables[0].add_row()
        ins = OxmlElement("w:ins")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "履歴付き追加"
        run.append(text)
        ins.append(run)
        doc.tables[0].cell(0, 1).paragraphs[0]._p.append(ins)
        doc.save(self.edited)
        result = h.compare(self.state, self.edited)
        self.assertTrue(result["manual_mapping_required"])
        self.assertEqual(result["tracked_changes"], 1)
        self.assertTrue(any("履歴付き追加" in (x["after"] or "") for x in self.state["human_revision_history"][0]["changes"]))

    def test_unchanged_creates_no_revision_and_format_only_is_not_text_change(self):
        self.assertTrue(h.compare(self.state, self.original)["unchanged"])
        self.assertEqual(self.state["human_revision_history"], [])
        doc = Document(self.original)
        doc.tables[0].cell(0, 1).paragraphs[0].runs[0].bold = True
        doc.save(self.edited)
        result = h.compare(self.state, self.edited)
        self.assertEqual(result["text_changes"], 0)
        self.assertTrue(self.state["human_revision_history"][0]["changed_package_parts"])


if __name__ == "__main__":
    unittest.main()
