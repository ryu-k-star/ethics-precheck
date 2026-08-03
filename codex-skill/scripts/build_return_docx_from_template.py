from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

from docx import Document
from docx.shared import Pt


RETURN_PREFIX = "★倫理申請_修正対応依頼_"
OUTPUT_FOLDER_PREFIX = "99_プレチェック出力"
WINDOWS_FILENAME_TRANSLATION = str.maketrans(
    {
        '"': "”",
        "*": "＊",
        "/": "／",
        ":": "：",
        "<": "＜",
        ">": "＞",
        "?": "？",
        "\\": "￥",
        "|": "｜",
    }
)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8").strip()


def split_bottom_sections(text: str) -> tuple[str, str]:
    marker = "A．倫理審査申請システム質疑事項"
    idx = text.find(marker)
    if idx == -1:
        return text, ""
    return text[:idx].strip(), text[idx:].strip()


def extract_title(text: str) -> str:
    for line in text.splitlines():
        if line.startswith("研究課題名"):
            return line.strip()
    return "研究課題名【】"


def extract_study_title(text: str) -> str:
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("研究課題名"):
            continue
        value = stripped[len("研究課題名") :].strip()
        if value.startswith("【") and value.endswith("】"):
            value = value[1:-1].strip()
        else:
            value = value.lstrip("：:").strip()
        if value:
            return value
    raise ValueError("Draft must contain a non-empty Japanese study title.")


def sanitize_filename_component(value: str) -> str:
    sanitized = value.translate(WINDOWS_FILENAME_TRANSLATION)
    sanitized = "".join(" " if ord(char) < 32 else char for char in sanitized)
    sanitized = re.sub(r"\s+", " ", sanitized).strip(" .")
    if not sanitized:
        raise ValueError("Study title becomes empty after filename sanitization.")
    return sanitized


def validate_output_path(output: Path, study_title: str) -> None:
    output_folder = output.parent.name
    if output_folder != OUTPUT_FOLDER_PREFIX and not output_folder.startswith(
        f"{OUTPUT_FOLDER_PREFIX}_"
    ):
        raise ValueError(
            f"Return DOCX must be saved in {OUTPUT_FOLDER_PREFIX}/ or "
            f"{OUTPUT_FOLDER_PREFIX}_<label>/; got: {output_folder}"
        )

    safe_title = sanitize_filename_component(study_title)
    expected_pattern = re.compile(
        rf"^{re.escape(RETURN_PREFIX + safe_title)}_(\d{{4}})\.docx$",
        re.IGNORECASE,
    )
    if not expected_pattern.fullmatch(output.name):
        expected = f"{RETURN_PREFIX}{safe_title}_作成日4桁.docx"
        raise ValueError(
            "Return DOCX filename does not follow the required Japanese naming rule. "
            f"Expected: {expected}; got: {output.name}"
        )


def parse_markdown_table(text: str) -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    in_table = False
    for line in text.splitlines():
        if line.startswith("| 項目 | 対応要否 |"):
            in_table = True
            continue
        if not in_table:
            continue
        if line.startswith("|---"):
            continue
        if not line.startswith("|"):
            if rows:
                break
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            continue
        rows.append((cells[0], cells[1], cells[2].replace("<br>", "\n")))
    return rows


def clear_cell(cell) -> None:
    cell.text = ""


def build_docx(template: Path, draft: Path, output: Path) -> None:
    draft_text = read_text(draft)
    study_title = extract_study_title(draft_text)
    validate_output_path(output, study_title)

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    doc = Document(output)
    doc.styles["Normal"].font.name = "Yu Gothic"
    doc.styles["Normal"].font.size = Pt(10.5)

    top_text, bottom_text = split_bottom_sections(draft_text)
    title = extract_title(top_text)

    if len(doc.paragraphs) > 1:
        doc.paragraphs[1].text = title

    rows = parse_markdown_table(draft_text)
    if not rows:
        raise ValueError("No A-section markdown table found in draft.")
    if len(doc.tables) < 2:
        raise ValueError("Template must contain at least two tables.")

    table = doc.tables[0]
    for row_idx in range(len(table.rows)):
        clear_cell(table.cell(row_idx, 1))

    max_rows = min(len(rows), len(table.rows))
    for idx in range(max_rows):
        _item, status, comment = rows[idx]
        if status == "要対応" and comment:
            table.cell(idx, 1).text = comment

    doc.tables[1].cell(0, 0).text = bottom_text
    doc.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a return DOCX while preserving the official template tables."
    )
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--draft", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    build_docx(args.template, args.draft, args.output)


if __name__ == "__main__":
    main()
