from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).with_name("build_return_docx_from_template.py")
SPEC = importlib.util.spec_from_file_location("return_builder", MODULE_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"Cannot load builder module: {MODULE_PATH}")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)


class ReturnDocxFilenameGateTests(unittest.TestCase):
    def test_builds_return_docx_with_required_japanese_filename(self) -> None:
        template = Path(__file__).parents[1] / "★倫理申請_修正対応依頼.docx"
        table_rows = ["| 項目 | 対応要否 | 内容 |", "|---|---|---|"]
        table_rows.extend(
            f"| {item} | 対応不要 | |" for item in range(1, 53)
        )
        draft_text = "\n".join(
            [
                "研究課題名【睡眠に関する研究】",
                *table_rows,
                "",
                "A．倫理審査申請システム質疑事項",
                "B．修正が必要な添付書類の再アップロード",
            ]
        )

        temp_root = Path(__file__).parents[1] / "tmp" / "tests" / "filename_gate"
        temp_root.mkdir(parents=True, exist_ok=True)
        draft = temp_root / "04_指摘事項書ドラフト.md"
        draft.write_text(draft_text, encoding="utf-8")
        output = (
            temp_root
            / "99_プレチェック出力"
            / "★倫理申請_修正対応依頼_睡眠に関する研究_0803.docx"
        )

        BUILDER.build_docx(template, draft, output)

        self.assertTrue(output.is_file())
        document = BUILDER.Document(output)
        self.assertEqual([len(table.rows) for table in document.tables], [52, 1])

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
