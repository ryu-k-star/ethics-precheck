---
name: ethics-precheck
description: Precheck Japanese clinical-research ethics applications with source-grounded independent review and a template-preserving return Word. Use the compact workflow for initial and resubmitted applications, not committee approval decisions.
---

# Ethics Precheck

Read [the standard workflow](references/precheck_v3_workflow.md). It specifies commands and record schemas. In the installed skill replace tools/ with scripts/, rules/ with references/rules/, and use assets/return-template.docx.

## Output and reading discipline

Persist only the return Word, 確認事項.md when human questions exist, and _internal/review_state.json. Do not create separate inventories, reviewer reports, draft Markdown, factcheck narratives, verification reports, README files or completion reports. Store a finding's final wording once. Programs derive temporary views for the existing validators; those are not extra deliverables or prompts to write.

Use scripts/precheck_scan.py for compact preparation. It caches extraction and compiles assigned rule groups including exceptions. Use scripts/precheck_compact.py:
- view: retrieve one role's input, source excerpts, timeline, ledger summary, editor guidance or factcheck input; do not save views as extra files.
- record: ingest structured reviewer/director/factchecker records. --consume removes only the successfully ingested temporary JSON directly inside _internal; never the state or originals.
- render --date MMDD: build the Word after independent preliminary factcheck; export human questions only if needed.
- verify: check final source reconciliation and Word transfer; store the result inside the state rather than another report.

Keep existing case outputs unchanged. --legacy-layout is for explicitly requested compatibility work, not ordinary new cases.

## Review requirements

- Preserve originals and human edits. Never publish study files or contact researchers without explicit authorization.
- Cover all 51 items, every R1/R2 rule and all seven domains. Read all assigned raw text and original rule wording including exceptions. No keyword match does not prove absence or non-applicability. Retrieve other relevant documents/sections from the complete catalog.
- Use real independent research/ethics and privacy/consent subagents without full history (fork_turns="none"). Give only role, output directory, view commands and record format. Neither primary reviewer sees prior findings, the draft, or the other reviewer's conclusions. A separate factchecker verifies findings/evidence before Word generation and the final Word afterward.
- Mark every rule na/satisfied/candidate/human with a short reason and source IDs. Record every raw source as read, visually checked or explicitly unresolved. PDF/Word text extraction does not establish image/table/selection-state completeness. Check the originals; missing evidence goes to human questions, not a finding of absent content.
- Assess final enrollment, follow-up and analysis through study end. Do not invent missing timing values. Verify official regulatory sources when currency affects a finding.
- Reconcile every primary finding into the ledger. Retain reasons and resolution routes for excluded findings. Secretariat/profile/other-route handling is not evidence that a finding was wrong.
- Preserve the institution's overreach guards: no separate fixed end date solely for future use; no media-specific disposal steps when inclusive disposal suffices; no mandatory copying of an application-system email address into participant information; no extra withdrawal-form submission details when rights, non-disadvantage and required forms are present. Real contradictions remain reviewable.
- Use applicable R4/R5 wording via the editor view only after findings are fixed. The factchecker receives final findings, source evidence, timeline, human questions and actual Word entries; do not repeat full 108-rule and 52-row reports as prose.
- Bind factcheck records to both packet_sha256 and snapshot_hash. Word formatting or content changes require renewed inspection. Set checked flags only after performing the checks.

For application item 15, apply the facility UI guard in R2 C1: only the third option has a free-text field. Never require free text for either of the first two options or switch options merely to reveal it. Review all documents together and request only indispensable corrections at an editable location. Escalate uncertain necessity or routing to human questions.

## Word and final gates

Copy the official template; clear all 52 response cells, including its 27 example/internal-note cells, then populate adopted findings. Preserve two tables (52×2 and 1×1), layout and sections. Save as ★倫理申請_修正対応依頼_<研究課題名>_<MMDD>.docx in a 99_プレチェック出力* folder.

Independent final factcheck, all-page Word visual inspection and verifier PASS are required. The program verifies rule coverage, source freshness, omitted findings/items, A/B text transfer, blank cells, comments/tracked changes, filename and independent confirmation. It does not establish substantive correctness or committee approval. If subagents/rendering are unavailable, record that limitation and leave the relevant gate incomplete.

## Conditional references

- [Quality playbook](references/review_quality_playbook.md): relevant omission/overreach/resubmission sections before review and drafting; do not read historical cases first.
- [Detailed contract](references/workflow.md): institution procedures and human edits (§6.8), only as needed. The current compact output/command workflow supersedes older multi-file instructions.
- [Example policy](references/examples/review_examples_policy.md): after independent findings and initial wording; tone/length only, never issue counts or completeness.

Resubmissions retain independent zero-based review followed by prior-finding reconciliation. Human-edited Word requires all-52-row reconciliation and explicit disposition reasons; never overwrite it or mechanically restore removed findings.

The first passing verification freezes the delivered AI wording in review_state.json under ai_initial. Preserve it permanently within the case; never replace it with edited wording. Humans edit the Word only. Use precheck_compact.py compare-human --out OUTPUT --word EDITED.docx to append before/after changes to human_revision_history in that same file. Review reasons and routing before adopting a change as a training example or updating rules. Comparisons do not train a model, upload data or approve a revised Word. Structural changes and tracked changes require explicit review. See the human revision section in the standard workflow.

Do not produce usage ledgers in ordinary cases. Explicit usage audits alone record final cumulative input, cached input and output; do not sum intermediate totals or infer missing values. File counts and character counts are not measured subscription-usage savings.
