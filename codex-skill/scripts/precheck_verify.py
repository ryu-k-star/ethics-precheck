from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

from docx import Document


RULE_ID = re.compile(r"^\|\s*((?:G|C)\d+-[0-9A-Za-z]+)\s*\|")
A_ROW = re.compile(r"^\|\s*([^|]+?)\s*\|\s*(要対応|対応不要|人間確認|要検討)\s*\|\s*(.*?)\s*\|$")
CANDIDATE_ROW = re.compile(r"^\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|")
RULE_FILE_GROUPS = (
    ("R1_項目間整合ルール.md", "r1-cross-document-consistency.md"),
    ("R2_必須記載チェックリスト.md", "r2-required-content.md"),
)


def expected_rule_ids(rules_dir: Path) -> list[str]:
    result: list[str] = []
    for alternatives in RULE_FILE_GROUPS:
        path = next((rules_dir / name for name in alternatives if (rules_dir / name).is_file()), None)
        if path is None:
            raise FileNotFoundError(f"Missing rule file under {rules_dir}: one of {alternatives}")
        for line in path.read_text(encoding="utf-8").splitlines():
            match = RULE_ID.match(line)
            if match and match.group(1) not in result:
                result.append(match.group(1))
    return result


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def item_numbers(value: str) -> set[int]:
    return {int(x) for x in re.findall(r"(?<!\d)([1-9]|[1-4]\d|5[01])(?!\d)", value)}


def docx_xml_flags(path: Path) -> dict[str, int]:
    comments = tracked = 0
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        comments = sum(1 for name in names if name.startswith("word/comments"))
        for name in names:
            if name.startswith("word/") and name.endswith(".xml"):
                data = package.read(name)
                tracked += len(re.findall(br"<w:(?:ins|del)(?:\s|>)", data))
    return {"comment_parts": comments, "tracked_change_tags": tracked}


def add(checks: list[dict[str, Any]], check_id: str, passed: bool, detail: str) -> None:
    checks.append({"id": check_id, "passed": bool(passed), "detail": detail})


def row_item_number(label: str) -> int | None:
    match = re.match(r"\s*(\d{1,2})\.", label or "")
    return int(match.group(1)) if match else None


def verify(args: argparse.Namespace) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    expected = expected_rule_ids(args.rules)
    ledger = json.loads(args.ledger.read_text(encoding="utf-8"))
    by_id = {row.get("rule_id"): row for row in ledger.get("rules", [])}
    missing = [rid for rid in expected if rid not in by_id]
    invalid = [rid for rid, row in by_id.items() if row.get("status") not in {"na", "satisfied", "candidate", "human"}]
    add(checks, "rules.all_present", not missing, f"missing={missing}")
    add(checks, "rules.all_classified", not invalid, f"unclassified_or_invalid={invalid}")
    duration_ids = [rid for rid in expected if rid.startswith("G6-")]
    duration_bad = [rid for rid in duration_ids if rid not in by_id or by_id[rid].get("status") not in {"na", "satisfied", "candidate", "human"}]
    add(checks, "duration.g6_complete", not duration_bad, f"required={duration_ids}; bad={duration_bad}")

    review = args.review.read_text(encoding="utf-8")
    draft = args.draft.read_text(encoding="utf-8")
    factcheck = args.factcheck.read_text(encoding="utf-8")
    x: set[int] = set()
    in_a_candidates = False
    for line in review.splitlines():
        if line.strip() == "## A欄候補":
            in_a_candidates = True
            continue
        if in_a_candidates and line.startswith("## "):
            in_a_candidates = False
        if not in_a_candidates:
            continue
        match = CANDIDATE_ROW.match(line)
        if match and match.group(1).strip() not in {"ルールID", "---"}:
            x |= item_numbers(match.group(2))
    y: set[int] = set()
    draft_comments: list[str] = []
    for line in draft.splitlines():
        match = A_ROW.match(line)
        if match and match.group(2) == "要対応":
            y |= item_numbers(match.group(1))
            if match.group(3).strip():
                draft_comments.append(match.group(3).replace("<br>", "\n"))
    omitted = sorted(x - y)
    allowed_evidence = set(ledger.get("allowed_evidence_statuses", ["confirmed", "inferred", "human"]))
    non_return_routes = {
        "researcher_profile", "secretariat_internal", "human_decision", "resolved_elsewhere",
    }
    dispositions = ledger.get("dispositions", [])
    allowed_routes = set(ledger.get("allowed_return_routes", {
        "applicant_return", "attachment_return", "researcher_profile",
        "secretariat_internal", "human_decision", "resolved_elsewhere",
    }))
    invalid_dispositions: list[str] = []
    routed_items: set[int] = set()
    for row in dispositions:
        raw_finding_id = row.get("finding_id")
        finding_id = str(raw_finding_id or "<missing>")
        reason = str(row.get("disposition_reason") or "").strip()
        evidence_status = row.get("evidence_status")
        route = row.get("return_route")
        included = row.get("included_in_return")
        numbers = {int(n) for n in row.get("item_numbers", []) if isinstance(n, int) or str(n).isdigit()}
        action_owner = str(row.get("action_owner") or "").strip()
        if not raw_finding_id or not reason or evidence_status not in allowed_evidence or route not in allowed_routes or not action_owner or not numbers or not isinstance(included, bool):
            invalid_dispositions.append(finding_id)
            continue
        if not included and route in non_return_routes:
            routed_items |= numbers
    add(checks, "disposition.valid", not invalid_dispositions, f"invalid={invalid_dispositions}")
    unresolved = [number for number in omitted if number not in routed_items]
    add(checks, "transfer.03_to_04", not unresolved, f"X={sorted(x)}; Y={sorted(y)}; unresolved={unresolved}")

    document = Document(args.word)
    add(checks, "word.table_count", len(document.tables) == 2, f"actual={len(document.tables)}")
    if len(document.tables) >= 2:
        dims = [(len(t.rows), len(t.columns)) for t in document.tables[:2]]
        add(checks, "word.table_shape", dims == [(52, 2), (1, 1)], f"actual={dims}")
        word_text = normalize("\n".join(cell.text for table in document.tables for row in table.rows for cell in row.cells))
        missing_text = [text[:80] for text in draft_comments if normalize(text) not in word_text]
        add(checks, "transfer.04_to_word", not missing_text, f"missing_count={len(missing_text)}; samples={missing_text[:3]}")
        non_action_filled: list[int] = []
        for row in document.tables[0].rows:
            item_no = row_item_number(row.cells[0].text)
            if item_no not in y and len(row.cells) > 1 and normalize(row.cells[1].text):
                non_action_filled.append(item_no or -1)
        add(checks, "word.non_action_blank", not non_action_filled, f"filled={non_action_filled}")
        forbidden = ["読んだら消してください", "前回の申請では"]
        residue = [value for value in forbidden if value in word_text]
        add(checks, "word.no_standard_text_residue", not residue, f"residue={residue}")
    flags = docx_xml_flags(args.word)
    add(checks, "word.no_comments", flags["comment_parts"] == 0, str(flags))
    add(checks, "word.no_tracked_changes", flags["tracked_change_tags"] == 0, str(flags))
    add(checks, "word.filename", args.word.name.startswith("★倫理申請_修正対応依頼_") and re.search(r"_\d{4}\.docx$", args.word.name) is not None, args.word.name)
    return {"passed": all(c["passed"] for c in checks), "checks": checks}


def write_report(result: dict[str, Any], json_path: Path, md_path: Path) -> None:
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# プレチェック機械検証結果", "", f"総合結果：{'PASS' if result['passed'] else 'FAIL'}", "", "| 検証 | 結果 | 詳細 |", "|---|---|---|"]
    for check in result["checks"]:
        detail = str(check["detail"]).replace("|", "\\|").replace("\n", " ")
        lines.append(f"| {check['id']} | {'PASS' if check['passed'] else 'FAIL'} | {detail} |")
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Deterministically verify precheck transfers and return DOCX structure.")
    parser.add_argument("--rules", type=Path, default=Path("rules"))
    parser.add_argument("--ledger", type=Path, required=True)
    parser.add_argument("--review", type=Path, required=True, help="03 review markdown")
    parser.add_argument("--draft", type=Path, required=True, help="04 draft markdown")
    parser.add_argument("--factcheck", type=Path, required=True, help="06 fact-check markdown")
    parser.add_argument("--word", type=Path, required=True)
    parser.add_argument("--json-out", type=Path, required=True)
    parser.add_argument("--md-out", type=Path, required=True)
    args = parser.parse_args()
    result = verify(args)
    write_report(result, args.json_out, args.md_out)
    print(json.dumps({"passed": result["passed"], "failed": [c["id"] for c in result["checks"] if not c["passed"]]}, ensure_ascii=False))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
