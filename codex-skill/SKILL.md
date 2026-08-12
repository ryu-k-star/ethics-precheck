---
name: ethics-precheck
description: Review Japanese clinical-research and ethics-committee application packages, audit cross-document consistency and required disclosures, prepare evidence-linked Markdown findings, and build the institution-style return DOCX. Use for initial or resubmitted IRB/ethics applications containing PDF or DOCX plans, application-system exports, participant information and consent forms, opt-out notices, questionnaires, data/sample handling documents, or prior correction requests.
---

# Ethics Precheck

Perform an evidence-first precheck without replacing the ethics committee's judgment. Produce a traceable review package and a template-preserving return Word file only after all quality gates pass.

## Start safely

1. Work inside the user-designated study folder. Never commit or publish study files, extracted text, personal data, or generated review outputs.
2. Preserve every original. Write extraction files to a separate temporary folder and results to a new `99_プレチェック出力/` folder; never overwrite a prior output folder.
3. Read these references in order before reviewing source materials:
   - `references/rules/rules-index.md`
   - `references/rules/r1-cross-document-consistency.md`
   - `references/rules/r2-required-content.md`
   - `references/rules/r3-format-and-word-qa.md`
   - `references/rules/r4-writing-style.md`
   - `references/rules/r5-office-standard-text.md`
   - `references/review_quality_playbook.md`
4. Read `references/workflow.md` for the complete operating contract, required output schemas, resubmission rules, and quality gates.
5. Do not open historical review outputs before independently applying R1-R3. After drafting, read `references/examples/review_examples_policy.md` only to calibrate tone and length. Treat no historical case as Gold unless independently audited and explicitly human-approved.

## Run the review

### 1. Inventory and extract

- List every file, type, version, date, study title, and principal investigator found in the package.
- Use `scripts/extract_research_docs.py <study-root> --out <temporary-output>` for batch PDF/DOCX extraction when appropriate. Inspect extraction errors and visually inspect image-only or layout-dependent pages; extracted text is not proof of visual completeness.
- Create a page-addressable fact table covering people and roles, sites, study design, intervention/invasiveness, participant criteria and counts, periods, consent route, samples versus information, processing category, transfers, storage, disposal, future use, costs, compensation, conflicts, and contact details.
- If legal or regulatory currency affects a finding, verify the current official Japanese source before relying on it. Record the official source and access date; do not rely on memory or a secondary summary alone.

### 2. Apply every rule

- Apply every applicable ID in R1 and R2 to the raw source package, plus the mechanical checks in R3.
- For each rule, record one status: not applicable, satisfied, finding candidate, or human confirmation required.
- Record the source document, page/item/table, and the compared wording for every candidate.
- Keep unsupported inferences out of findings. Put facility-policy questions, unreadable content, missing evidence, and committee judgments into `05_人間確認事項.md`.
- On resubmissions, keep two independent passes. A blind reviewer should complete the zero-based 51-item review without prior findings. A separate pass should disposition every prior finding. If only one reviewer is available, finish and save the zero-based pass before opening the prior review.

### 3. Create the audit trail

Create these files in the study output folder:

1. `01_資料一覧.md`
2. `02_研究種別_仮分類.md`
3. `03_サブAI別レビュー結果.md`
4. `04_指摘事項書ドラフト.md`
5. `05_人間確認事項.md`
6. `06_ファクトチェック結果.md`

Use the schemas in `references/workflow.md`. Cover all seven review domains even when one reviewer handles several domains. If subagents are available, independently assign at least research-plan/ethics review and fact-check/evidence reconciliation. If unavailable, state that fact and the reason in 03 or 06.

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

## Build and verify the Word return

1. Copy `assets/return-template.docx`; never edit the bundled template itself. Its English asset name is internal only and must not become the researcher-facing filename.
2. Before using the builder, require the A-section header `| 項目 | 対応要否 |` and all 52 application rows in exact template order. After the table, require exactly one footer marker `A．倫理審査申請システム質疑事項` that begins the A/B/C instruction block. The same words may appear in a heading before the table; the builder must ignore pre-table occurrences. The builder maps rows positionally rather than by item number and must stop before copying when rows are missing, duplicated, out of order, use unsupported statuses, or the post-table footer marker is missing or duplicated.
3. Save the copied return Word inside `99_プレチェック出力/` or `99_プレチェック出力_<label>/` with the exact Japanese pattern `★倫理申請_修正対応依頼_<研究課題名>_<作成日4桁>.docx`. Never accept `return.docx`, another placeholder, or an English filename as final. Replace Windows-forbidden title characters (`\ / : * ? " < > |`) with their full-width equivalents so the study title remains recognizable. Then use `scripts/build_return_docx_from_template.py --template <template> --draft <04.md> --output "<output-folder>/★倫理申請_修正対応依頼_<研究課題名>_<作成日4桁>.docx"`. The builder validates the folder, filename, template shape, row labels and order, statuses, and footer marker before creating or replacing the output.
4. Verify the requirements in R3 and `references/workflow.md`, including two tables, Table 1 at 52 rows by 2 columns, Table 2 at 1 row by 1 column, preserved page setup and margins, no comments, no tracked changes, correct study title, correct A-item placement, correct B/C filenames, the required Japanese filename, and exact 04-to-Word transfer.
5. Render the DOCX and inspect every page when rendering is available. If only structural checks were possible, state that visual QA remains incomplete.
6. Treat the result as a draft for human approval, never as committee approval or legal advice.

## Reconcile human edits

When a human corrects the returned Word, compare all 52 rows with 04. Update 04, preserve rejected candidates with reasons, and update the appropriate rule or writing example only when the correction is generalizable. Never mechanically restore a candidate the human intentionally deleted.

## Resource map

- `references/workflow.md`: complete workflow and output contract
- `references/review_quality_playbook.md`: omission, overreach, resubmission, and transfer QA
- `references/rules/`: rule IDs, wording patterns, template checks, provenance, and reverse-application evidence
- `references/examples/review_examples_policy.md`: permitted use of historical examples
- `scripts/extract_research_docs.py`: batch extraction and manifest creation
- `scripts/build_return_docx_from_template.py`: template-preserving DOCX builder
- `assets/return-template.docx`: blank institutional return template
