---
name: ethics-precheck
description: Precheck Japanese clinical-research ethics application packages by auditing evidence across PDF/DOCX files, recording item-by-item findings, and producing a template-preserving return DOCX. Use for initial or resubmitted IRB/ethics applications; do not use for generic ethical advice, committee approval decisions, or legal/compliance certification.
---

# Ethics Precheck

Perform an evidence-first precheck without replacing the ethics committee's judgment. Produce a traceable review package and a template-preserving return Word file only after all quality gates pass.

## Confirm the assignment

- Identify the user-designated study folder, whether this is an initial or resubmitted application, and the requested output location before processing files.
- Locate the current application package, facility-specific rules and template, and any prior correction request supplied by the user. Treat missing material as a review limitation, not as evidence that the application is deficient.
- Confirm that Python 3.9 or newer plus the packages in `requirements.txt` are available before running the bundled scripts. If a required tool is unavailable, continue only where reliable and state the resulting limitation.
- Handle all study data under the user's authorization and institutional data rules. Do not upload source documents, extracted text, or outputs to an external service unless the user has explicitly authorized that destination.

## Start safely

1. Work inside the user-designated study folder. Never commit or publish study files, extracted text, personal data, or generated review outputs.
2. Preserve every original. Write extraction files to a separate temporary folder and results to a new `99_プレチェック出力/` folder; never overwrite a prior output folder.
3. Read these references in order before reviewing source materials:
   - `references/rules/rules-index.md`
   - `references/rules/r1-cross-document-consistency.md`
   - `references/rules/r2-required-content.md`
   - `references/rules/r3-format-and-word-qa.md`
   - `references/review_quality_playbook.md`
4. Read `references/precheck_v2_workflow.md` for the standard low-token execution path. Read `references/workflow.md` only when the complete operating contract, detailed schemas, resubmission rules, or facility-specific quality gates are needed.
5. After the evidence findings are fixed, read only the relevant sections of `references/rules/r4-writing-style.md` and `references/rules/r5-office-standard-text.md` needed to draft those findings.
6. Do not open historical review outputs before independently applying R1-R3. After drafting, read `references/examples/review_examples_policy.md` only to calibrate tone and length. Treat no historical case as Gold unless independently audited and explicitly human-approved.

Keep the split rule files as the source of truth. Do not concatenate the workflow and all rules into one prompt. Load R1-R3 for evidence review, then only the relevant R4/R5 sections for drafting.

## Use the v2 precheck pipeline

For a standard local study folder, start with one deterministic extraction:

```text
python scripts/precheck_scan.py <study-root> --out <new-output-folder> --cache <cache-folder> --rules references/rules
```

- The scanner hashes each source and reuses extraction only when the extractor version and source SHA-256 match.
- It excludes prior `99_プレチェック出力*`, `tmp`, `output(s)`, cache, and `.git` directories from source discovery. It still records a warning when an excluded directory contains PDF/DOCX files so misplaced source documents are visible. Review `00_r3_hits.json`, which applies machine-readable R3 T1/T2 patterns to whitespace-normalized source text.
- Reuse extracted text and the compact fact pack across reviewers. The fact pack includes R3 hits, cross-document fact candidates, and application-item gap candidates; reconcile them against cited source pages rather than asking each reviewer to rediscover the same inventory or extract the same pages.
- Treat cached extraction as a speed optimization, not as an exemption from visual inspection of image-only or layout-dependent pages.

## Run the review

### 1. Inventory and extract

- List every file, type, version, date, study title, and principal investigator found in the package.
- Use `scripts/extract_research_docs.py <study-root> --out <temporary-output>` for batch PDF/DOCX extraction when appropriate. Inspect extraction errors and visually inspect image-only or layout-dependent pages; extracted text is not proof of visual completeness.
- Create a page-addressable fact table covering people and roles, sites, study design, intervention/invasiveness, participant criteria and counts, periods, consent route, samples versus information, processing category, transfers, storage, disposal, future use, costs, compensation, conflicts, and contact details.
- Complete the G6 timeline gate using `00_timeline_facts.json`: study start/end, target and maximum enrollment, expected enrollment completion, longest follow-up, data lock, analysis duration, and remaining analysis margin. If enrollment speed, follow-up, or analysis duration is absent, do not invent it; identify the missing input and return the feasibility question for human confirmation.
- If legal or regulatory currency affects a finding, verify the current official Japanese source before relying on it. Record the official source and access date; do not rely on memory or a secondary summary alone.

### 2. Apply every rule

- Apply every applicable ID in R1 and R2 to the raw source package, plus the mechanical checks in R3.
- For each rule, record one status: not applicable, satisfied, finding candidate, or human confirmation required.
- Record the source document, page/item/table, and the compared wording for every candidate.
- Keep unsupported inferences out of findings. Put facility-policy questions, unreadable content, missing evidence, and committee judgments into `05_人間確認事項.md`.
- On resubmissions, keep two independent passes. A blind reviewer should complete the zero-based 51-item review without prior findings. A separate pass should disposition every prior finding. If only one reviewer is available, finish and save the zero-based pass before opening the prior review.

### 3. Create the audit trail

Create these files in the study output folder:

1. `00_source_manifest.json`
2. `00_extracted_pages.json`
3. `00_timeline_facts.json`
4. `00_r3_hits.json`
5. `00_rule_ledger.json`
6. `00_fact_pack.json`
7. `01_資料一覧.md`
8. `02_研究種別_仮分類.md`
9. `03_サブAI別レビュー結果.md`
10. `04_指摘事項書ドラフト.md`
11. `05_人間確認事項.md`
12. `06_ファクトチェック結果.md`
13. `07_検証結果.json` and `07_検証結果.md`
14. `08_token_usage.json` when usage totals are available, or an unavailable entry with the reason

Use the schemas in `references/workflow.md`. Cover all seven review domains even when one reviewer handles several domains. If subagents are available, independently assign at least research-plan/ethics review and fact-check/evidence reconciliation. Give each subagent only its role, the compact fact/timeline pack, assigned rules, and necessary source pages; do not pass the full conversation or every reference file. If subagents are unavailable, state that fact and the reason in 03 or 06.

For v2 candidates, use the five-column Finding ID schemas in `references/precheck_v2_workflow.md`; they supersede the legacy four-column candidate examples in the complete workflow. Give every A/B candidate a stable ID and a matching ledger disposition so multiple findings under one application item cannot collapse into one.

### 4. Draft for action

- Use R4's current-state, problem, request structure and R5's institution wording where applicable.
- Separate A (application-system items), B (documents requiring corrected re-upload), C (minor wording/format corrections), and human-only decisions.
- Cite the source and page/item for each adopted finding. Prefer a confirmation request when facts or facility policy remain uncertain.
- Do not optimize for issue count. Remove only unsupported, speculative, merely stylistic, or reference-only candidates, and log the reason.
- Apply the facility overreach guards in R2 and R4. Do not require a separate fixed end date for future-use retention, media-specific disposal steps when an inclusive disposal policy is present, an application-system email address to be copied into the participant information sheet, or withdrawal-form submission details when withdrawal rights, non-disadvantage, and the required form are already present.

## Enforce three distinct audits

Do not substitute one audit for another:

1. **Omission and overreach audit:** Independently search raw sources for missing issues and test each proposed finding against its evidence.
2. **03 to 04 transfer audit:** Confirm every A/B candidate in 03 appears in 04 or has an explicit, evidence-based disposition in 06.
3. **04 to Word transfer audit:** Confirm every final A/B text appears in the correct Word location and every non-applicable A cell is blank.

Do not declare completion when any candidate disappears silently.

Run `scripts/precheck_verify.py` after Word generation. It must pass all rule classification, G6, disposition, 03-to-04, 04-to-Word, table-shape, blank-cell, comment, tracked-change, template-residue, and filename checks before the Word file is a return candidate. A PASS does not replace evidence review or all-page visual inspection.

## Build and verify the Word return

1. Copy `assets/return-template.docx`; never edit the bundled template itself. Its English asset name is internal only and must not become the researcher-facing filename.
2. Before using the builder, require the A-section header `| 項目 | 対応要否 |` and all 52 application rows in exact template order. After the table, require exactly one footer marker `A．倫理審査申請システム質疑事項` that begins the A/B/C instruction block. The same words may appear in a heading before the table; the builder must ignore pre-table occurrences. The builder maps rows positionally rather than by item number and must stop before copying when rows are missing, duplicated, out of order, use unsupported statuses, or the post-table footer marker is missing or duplicated.
3. Save the copied return Word inside `99_プレチェック出力/` or `99_プレチェック出力_<label>/` with the exact Japanese pattern `★倫理申請_修正対応依頼_<研究課題名>_<作成日4桁>.docx`. Never accept `return.docx`, another placeholder, or an English filename as final. Replace Windows-forbidden title characters (`\ / : * ? " < > |`) with their full-width equivalents so the study title remains recognizable. Then use `scripts/build_return_docx_from_template.py --template <template> --draft <04.md> --output "<output-folder>/★倫理申請_修正対応依頼_<研究課題名>_<作成日4桁>.docx"`. The builder validates the folder, filename, template shape, row labels and order, statuses, and footer marker before creating or replacing the output.
4. Verify the requirements in R3 and `references/workflow.md`, including two tables, Table 1 at 52 rows by 2 columns, Table 2 at 1 row by 1 column, preserved page setup and margins, no comments, no tracked changes, correct study title, correct A-item placement, correct B/C filenames, the required Japanese filename, and exact 04-to-Word transfer.
5. Render the DOCX and inspect every page when rendering is available. If only structural checks were possible, state that visual QA remains incomplete.
6. Treat the result as a draft for human approval, never as committee approval or legal advice.

## Reconcile human edits

When a human corrects the returned Word, compare all 52 rows with 04. Update 04 and preserve removed candidates with a reason. Separate whether the finding is valid from how it should be resolved: applicant return, attachment correction, researcher-profile change, secretariat handling, human decision, or another route. A finding removed from the return Word is not a negative example unless the human confirmed that it was wrong or overreaching. Update the appropriate rule, wording example, or regression test only when the correction is generalizable. Never mechanically restore a candidate the human intentionally deleted.

## Resource map

- `references/precheck_v2_workflow.md`: standard cached, role-scoped, verified execution path
- `references/workflow.md`: complete workflow and output contract
- `references/review_quality_playbook.md`: omission, overreach, resubmission, and transfer QA
- `references/rules/`: rule IDs, wording patterns, template checks, provenance, and reverse-application evidence
- `references/examples/review_examples_policy.md`: permitted use of historical examples
- `scripts/extract_research_docs.py`: batch extraction and manifest creation
- `scripts/precheck_scan.py`: hash-based one-time extraction, source exclusion, R3 T1/T2 scanning, timeline evidence, and rule ledger
- `scripts/precheck_verify.py`: deterministic rule, transfer, DOCX structure, and residue verification
- `scripts/precheck_usage.py`: deduplicated per-study final token totals or an explicit unavailable record
- `scripts/build_return_docx_from_template.py`: template-preserving DOCX builder
- `assets/return-template.docx`: blank institutional return template
