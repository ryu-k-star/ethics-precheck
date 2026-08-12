from __future__ import annotations

import importlib.util
import re
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).with_name("build_return_docx_from_template.py")
SPEC = importlib.util.spec_from_file_location("return_builder", MODULE_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"Cannot load builder module: {MODULE_PATH}")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)


class ReturnDocxFilenameGateTests(unittest.TestCase):
    @staticmethod
    def template_path() -> Path:
        return Path(__file__).parents[1] / "★倫理申請_修正対応依頼.docx"

    @classmethod
    def template_labels(cls) -> list[str]:
        document = BUILDER.Document(cls.template_path())
        return [
            re.sub(r"\s+", " ", row.cells[0].text).strip()
            for row in document.tables[0].rows
        ]

    @classmethod
    def make_draft_text(
        cls,
        *,
        labels: list[str] | None = None,
        heading_before_table: str = "",
        footer: str = "A．倫理審査申請システム質疑事項\nB．修正が必要な添付書類の再アップロード",
        issue_index: int | None = None,
    ) -> str:
        labels = labels or cls.template_labels()
        table_rows = ["| 項目 | 対応要否 | 内容 |", "|---|---|---|"]
        for index, label in enumerate(labels):
            status = "要対応" if index == issue_index else "対応不要"
            comment = "施設区分を確認してください。" if index == issue_index else ""
            table_rows.append(f"| {label} | {status} | {comment} |")
        return "\n".join(
            [
                "研究課題名【睡眠に関する研究】",
                heading_before_table,
                *table_rows,
                "",
                footer,
            ]
        )

    @staticmethod
    def make_test_paths(label: str) -> tuple[Path, Path]:
        temp_root = Path(__file__).parents[1] / "tmp" / "tests" / label
        temp_root.mkdir(parents=True, exist_ok=True)
        draft = temp_root / "04_指摘事項書ドラフト.md"
        output = (
            temp_root
            / f"99_プレチェック出力_{label}"
            / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx"
        )
        output.unlink(missing_ok=True)
        return draft, output

    def test_builds_return_docx_with_required_japanese_filename(self) -> None:
        template = self.template_path()
        draft_text = self.make_draft_text(
            heading_before_table="## A．倫理審査申請システム質疑事項",
            issue_index=5,
        )
        draft, output = self.make_test_paths("filename_gate")
        draft.write_text(draft_text, encoding="utf-8")

        BUILDER.build_docx(template, draft, output)

        self.assertTrue(output.is_file())
        document = BUILDER.Document(output)
        self.assertEqual([len(table.rows) for table in document.tables], [52, 1])
        self.assertEqual(
            document.tables[0].cell(5, 1).text,
            "施設区分を確認してください。",
        )
        footer_text = document.tables[1].cell(0, 0).text
        self.assertTrue(footer_text.startswith("A．倫理審査申請システム質疑事項"))
        self.assertNotIn("| 項目 | 対応要否 |", footer_text)
        expected_footer = draft_text.split(
            "A．倫理審査申請システム質疑事項", 2
        )[-1]
        expected_footer = "A．倫理審査申請システム質疑事項" + expected_footer
        self.assertEqual(footer_text, expected_footer.strip())
        self.assertGreater(len(document.tables[1].cell(0, 0).paragraphs), 1)

    def test_rejects_missing_a_section_row_before_copy(self) -> None:
        labels = self.template_labels()[:-1]
        draft_text = self.make_draft_text(labels=labels)
        draft, output = self.make_test_paths("missing_row")
        draft.write_text(draft_text, encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "exactly 52 rows"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertFalse(output.exists())

    def test_rejects_reordered_a_section_rows_before_copy(self) -> None:
        labels = self.template_labels()
        labels[4], labels[5] = labels[5], labels[4]
        draft_text = self.make_draft_text(labels=labels)
        draft, output = self.make_test_paths("reordered_rows")
        draft.write_text(draft_text, encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "row labels or order"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertFalse(output.exists())

    def test_rejects_footer_marker_missing_after_a_table(self) -> None:
        draft_text = self.make_draft_text(
            heading_before_table="## A．倫理審査申請システム質疑事項",
            footer="B．修正が必要な添付書類の再アップロード",
        )
        draft, output = self.make_test_paths("missing_footer_marker")
        draft.write_text(draft_text, encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "exactly one footer marker"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertFalse(output.exists())

    def test_rejects_duplicate_footer_markers_after_a_table(self) -> None:
        marker = "A．倫理審査申請システム質疑事項"
        draft_text = self.make_draft_text(footer=f"{marker}\n{marker}")
        draft, output = self.make_test_paths("duplicate_footer_marker")
        draft.write_text(draft_text, encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "found 2"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertFalse(output.exists())

    def test_rejects_unsupported_status_before_copy(self) -> None:
        draft_text = self.make_draft_text().replace(
            "| 1. 研究課題名 | 対応不要 |  |",
            "| 1. 研究課題名 | 要確認 |  |",
        )
        draft, output = self.make_test_paths("unsupported_status")
        draft.write_text(draft_text, encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "unsupported status"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertFalse(output.exists())

    def test_invalid_draft_does_not_replace_existing_output(self) -> None:
        draft_text = self.make_draft_text(labels=self.template_labels()[:-1])
        draft, output = self.make_test_paths("preserve_existing")
        draft.write_text(draft_text, encoding="utf-8")
        output.parent.mkdir(parents=True, exist_ok=True)
        original = b"existing-output-must-remain"
        output.write_bytes(original)

        with self.assertRaisesRegex(ValueError, "exactly 52 rows"):
            BUILDER.build_docx(self.template_path(), draft, output)
        self.assertEqual(output.read_bytes(), original)

    def test_accepts_required_japanese_filename(self) -> None:
        output = Path(
            "99_プレチェック出力"
        ) / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx"
        BUILDER.validate_output_path(output, "睡眠に関する研究")

    def test_accepts_model_specific_output_folder(self) -> None:
        output = Path(
            "99_プレチェック出力_Codex"
        ) / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx"
        BUILDER.validate_output_path(output, "睡眠に関する研究")

    def test_replaces_windows_forbidden_characters_with_fullwidth(self) -> None:
        output = Path(
            "99_プレチェック出力"
        ) / "★倫理申請_修正対応依頼_睡眠／覚醒：比較？研究_0803.docx"
        BUILDER.validate_output_path(output, "睡眠/覚醒:比較?研究")

    def test_rejects_placeholder_english_filename(self) -> None:
        with self.assertRaisesRegex(ValueError, "Japanese naming rule"):
            BUILDER.validate_output_path(
                Path("99_プレチェック出力") / "return.docx",
                "睡眠に関する研究",
            )

    def test_rejects_wrong_study_title(self) -> None:
        with self.assertRaisesRegex(ValueError, "Japanese naming rule"):
            BUILDER.validate_output_path(
                Path("99_プレチェック出力")
                / "★倫理申請_修正対応依頼_別の研究_0803.docx",
                "睡眠に関する研究",
            )

    def test_rejects_non_four_digit_date(self) -> None:
        with self.assertRaisesRegex(ValueError, "Japanese naming rule"):
            BUILDER.validate_output_path(
                Path("99_プレチェック出力")
                / "★倫理申請_修正対応依頼_睡眠に関する研究_20260803.docx",
                "睡眠に関する研究",
            )

    def test_rejects_wrong_output_folder(self) -> None:
        with self.assertRaisesRegex(ValueError, "must be saved"):
            BUILDER.validate_output_path(
                Path("output") / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx",
                "睡眠に関する研究",
            )

    def test_rejects_output_folder_without_required_separator(self) -> None:
        with self.assertRaisesRegex(ValueError, "must be saved"):
            BUILDER.validate_output_path(
                Path("99_プレチェック出力仮")
                / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx",
                "睡眠に関する研究",
            )


if __name__ == "__main__":
    unittest.main()
