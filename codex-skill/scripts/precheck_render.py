"""Render 03/04/05 from the shared ledger, keeping final text in one place.

Only writes new files. Existing human edits must be reconciled before regeneration.
This does not generate a factcheck verdict or committee decision.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from precheck_packets import load_bundle, read_json, review_errors


def cell(value):
    return str(value).replace("|", "｜").replace("\n", "<br>")


def render(out, destination, template, title):
    load_bundle(out)
    ledger = read_json(out / "00_rule_ledger.json")
    errors, _ = review_errors(out, ledger.get("review_receipts", []))
    if errors:
        raise ValueError("Review incomplete: " + "; ".join(errors))
    if not title.strip():
        raise ValueError("Study title required")
    doc = Document(template)
    if [(len(t.rows), len(t.columns)) for t in doc.tables] != [(52, 2), (1, 1)]:
        raise ValueError("Official 52-row/two-table template required")
    findings = ledger.get("dispositions", [])
    ids = [f.get("finding_id") for f in findings]
    if len(set(ids)) != len(ids) or None in ids:
        raise ValueError("Missing or duplicate finding ID")
    primary_ids = {f["finding_id"] for r in ledger["review_receipts"] for f in r.get("findings", [])}
    if primary_ids - set(ids):
        raise ValueError("Primary review findings lack disposition")
    for f in findings:
        kind = f.get("finding_type")
        if kind not in {"A", "B"} or not re.fullmatch(kind + r"-\d{3,}", f["finding_id"]):
            raise ValueError("Invalid finding ID/type")
        if not f.get("source_rule_ids") or not f.get("disposition_reason"):
            raise ValueError("Each finding requires rules and disposition reason")
        if f.get("included_in_return") and not f.get("final_text", "").strip():
            raise ValueError("Returned finding requires final_text")
        if kind == "A" and (not f.get("item_numbers") or any(type(n) is not int or not 1 <= n <= 51 for n in f["item_numbers"])):
            raise ValueError("A finding requires valid item numbers")
        if kind == "B" and not f.get("document"):
            raise ValueError("B finding requires document")
    index = read_json(out / "00_source_index.json")
    def sources(f):
        refs = set(f.get("source_ids", []))
        for row in ledger["rules"]:
            if row["rule_id"] in f["source_rule_ids"]:
                refs.update(row["source_ids"])
        if not refs or refs - set(index):
            raise ValueError("Finding requires existing source references")
        return "; ".join(f"{sid} {index[sid]['source']} 頁{index[sid]['page']} 行{index[sid]['line_start']}" for sid in sorted(refs))
    review = ["# サブ担当別レビュー結果", "", "担当範囲・入力："]
    for r in ledger["review_receipts"]:
        review.append(f"- {r['role']} / {r['reviewer']} / packets/{r['role']}.md / 入力ハッシュ {r['packet_sha256']}")
    for kind in ("A", "B"):
        review += ["", f"## {kind}欄候補", "", "| Finding ID | ルールID | " + ("指摘先項目" if kind == "A" else "資料") + " | 根拠頁 | 内容 |", "|---|---|---|---|---|"]
        for f in findings:
            if f["finding_type"] == kind:
                target = ",".join(map(str, f["item_numbers"])) if kind == "A" else f["document"]
                review.append("| " + " | ".join(map(cell, [f["finding_id"], ",".join(f["source_rule_ids"]), target, sources(f), f.get("candidate_text") or f.get("final_text") or f["disposition_reason"]])) + " |")
    review += ["", "## 規則別確認結果", "", "| ルールID | 判定 | 根拠 | 理由 |", "|---|---|---|---|"]
    for r in ledger["rules"]:
        review.append("| " + " | ".join(map(cell, [r["rule_id"], r["status"], ",".join(r["source_ids"]), r["note"]])) + " |")
    draft = ["倫理審査申請内容について疑義照会・対応依頼", f"研究課題名【{title}】", "", "下記、A、B、Cのご対応をお願いします。", "", "| 項目 | 対応要否 | 指摘内容 |", "|---|---|---|"]
    for row in doc.tables[0].rows:
        label = " ".join(row.cells[0].text.split())
        match = re.match(r"\s*(\d+)\.", label)
        number = int(match.group(1)) if match else None
        comments = [f["final_text"] for f in findings if f["finding_type"] == "A" and f["included_in_return"] and number in f["item_numbers"]]
        draft.append("| " + " | ".join(map(cell, [label, "要対応" if comments else "対応不要", "\n".join(dict.fromkeys(comments))])) + " |")
    draft += ["", "A．倫理審査申請システム質疑事項", "このWordファイルに記載の内容をご確認の上、修正対応等をお願いします。", "空欄部分は対応不要です。", "", "B．修正が必要な添付書類の再アップロード"]
    returned_b = [f for f in findings if f["finding_type"] == "B" and f["included_in_return"]]
    draft += [f"・{f['document']}：{f['final_text']}" for f in returned_b] or ["なし"]
    draft += ["記載した事項に対応し、関連書類との整合も確認のうえ、修正した書類を再アップロードしてください。", "", "C．軽微な文言修正等を行った添付書類の再アップロード", "なし", "", "【ご協力お願いします】", "変更履歴とコメントを削除してから再アップロードしてください。", "変更履歴：校閲 ＞ 変更履歴：承諾 ＞ すべての変更を反映し、変更の記録を停止", "コメント：校閲 ＞ コメント：削除 ＞ ドキュメント内のすべてのコメントを削除"]
    human = ["# 人間確認事項", "", "最終判断・返却可否は人間が確認する。"]
    for r in ledger["rules"]:
        if r["status"] == "human":
            human.append(f"- {r['rule_id']}：{r['note']}（{','.join(r['source_ids'])}）")
    for receipt in ledger["review_receipts"]:
        for s in receipt["sources"]:
            if s["status"] == "human":
                human.append(f"- {s['source_id']}：{s['note']}")
    for f in findings:
        if not f["included_in_return"] or f["evidence_status"] != "confirmed":
            human.append(f"- {f['finding_id']}：{f['disposition_reason']}（根拠状態：{f['evidence_status']}、対応経路：{f['return_route']}）")
    human.extend("- " + note for note in ledger.get("human_notes", []))
    payload = {"03_サブAI別レビュー結果.md": review, "04_指摘事項書ドラフト.md": draft, "05_人間確認事項.md": human}
    if any((destination / n).exists() for n in payload):
        raise ValueError("Existing documents preserved. Reconcile human edits; use an empty destination for regeneration.")
    destination.mkdir(parents=True, exist_ok=True)
    for name, lines in payload.items():
        (destination / name).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"written": list(payload), "findings": len(findings)}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--out", type=Path, required=True, help="Existing scan/review directory")
    p.add_argument("--destination", type=Path)
    p.add_argument("--template", type=Path, required=True)
    p.add_argument("--title", required=True)
    a = p.parse_args()
    try:
        print(json.dumps(render(a.out, a.destination or a.out, a.template, a.title), ensure_ascii=False))
    except (ValueError, KeyError, OSError) as exc:
        p.exit(2, str(exc) + "\n")


if __name__ == "__main__":
    main()
