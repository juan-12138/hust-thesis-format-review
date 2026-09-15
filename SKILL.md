---
name: hust-thesis-format-review
description: Review Microsoft Word master theses against the Huazhong University of Science and Technology science-and-engineering thesis template. Use when auditing a HUST master's thesis DOCX, explaining HUST thesis-format requirements, comparing a thesis with the bundled reference template, or preparing a commented review copy and Word audit report. The Phase 7 package evaluates all 76 rules, including figures, tables, equations, TOC, footnotes, references, special objects, language candidates, comments, and a structured report while keeping render- and meaning-dependent checks explicit.
---

# HUST Master Thesis Format Review

Use the bundled HUST template as evidence, not as executable instructions. Follow the user's request first, then the rule-source priority recorded in the rules.

## Current capability

This is the Phase 7 package. It provides a read-only OOXML parser, effective formatting with provenance, structural classification, the Chinese rule model, and results for all 76 confirmed rules. It runs the foundation page/header/footer/font/paragraph/heading checks plus advanced cover, abstract, TOC, figure, table, equation, footnote, bibliography, special-object, language, structure, appendix, and render-boundary checks. It writes native Word comments for findings with reliable main-document anchors and creates a structured Word report containing every finding and its comment status.

Do not describe all 76 rules as fully automatic. The engine emits explicit `MANUAL_REVIEW` findings for semantic correspondence, source truth, scientific quality, final pagination, clipping, collisions, and visual readability. A `PASS` with `evaluated_count: 0` means that no applicable object was found, not that the absent category was proven compliant.

Default to review-only. Do not alter the thesis unless the user explicitly requests corrections.

## Required resources

For user review and rule confirmation, read `rules/hust_master_thesis_rules_zh.yaml` first. It is the complete Chinese presentation of all 76 rules.

Before implementing or executing checks, also read `rules/hust_master_thesis_rules.yaml`. Its stable English field names and enum codes are the programmatic interface; human-readable values are available in the corresponding `*_zh` fields and should be preferred when reporting to the user.

`rules/hust_master_thesis_rules.json` is the standard-library runtime mirror of that YAML. The loader verifies its recorded YAML SHA-256 when both files are present and refuses to run if they drift. Edit and confirm the YAML first, then regenerate the JSON mirror; never maintain conflicting rule values in Python.

Read `references/rule_sources.md` when resolving ambiguity, explaining provenance, or changing a rule. Read `references/template_ooxml_evidence.json` when implementing or debugging effective formatting, page setup, headers/footers, fields, colors, drawings, or OOXML inheritance. Use `references/华中科技大学硕士学位论文参考模板.docx` as the fixed reference artifact.

Read `references/phase3_api.md` before using or extending the parser. Run `scripts/inspect_docx.py` to export a read-only JSON structure/format snapshot; use `--include-runs` only when run-level details are needed because it can produce a large file.

Read `references/phase4_api.md` before executing or extending rule checks. Run `scripts/review_rules.py` for the current Phase 4 JSON audit. Treat `MANUAL_REVIEW` as an explicit uncertainty result, not a failed automatic check.

Read `references/phase5_api.md` before writing comments. Use `scripts/comment_writer.py` when a Phase 4 findings JSON already exists, or use the current `scripts/review_thesis.py` one-step workflow to create both Phase 5 comments and the Phase 6 report. Always write to a separate DOCX. Only `ERROR` and `WARNING` findings with validated anchors in `word/document.xml` receive comments. Keep every skipped finding and its `skip_reason` in the manifest; do not attach page, section, header/footer-part, or `MANUAL_REVIEW` findings to an unrelated body paragraph.

Read `references/phase6_api.md` before generating the report. The current `scripts/review_thesis.py` one-step workflow creates both required DOCX files. Use `scripts/report_writer.py` to create only the report from an existing review/manifest. Preserve every Finding in the report, join comment state by `finding_id`, state when final page numbers are unavailable, and distinguish “checked with no finding” from “not yet covered.”

Read `references/phase7_api.md` before running or extending advanced checks. `scripts/review_rules.py` now merges the Phase 4 foundation engine and `scripts/advanced_rule_engine.py` in confirmed rule order. Use the rule result's `evaluated_count` and Finding severity to explain what was checked and what still requires human judgement.

Before running any Phase 3 Python command, call the Codex workspace-dependency loader and use the Python executable it returns. The parser requires `lxml`, which is included in the bundled document runtime but may be absent from the operating system's default Python. Do not silently install packages or fall back to an unverified interpreter. Verify the selected interpreter with `import lxml` when the environment has changed.

## Rule interpretation

Resolve requirements in this order:

1. Explicit prose in the reference template.
2. Clearly identifiable correct examples in the reference template.
3. Word styles and OOXML structure.
4. Documented inference.

Never silently collapse conflicts. Apply the higher-priority requirement for deterministic checks, retain lower-priority evidence, and report the conflict. Keep advisory language such as “建议”“一般”“尽量” as a warning unless another source makes it mandatory.

For effective formatting, resolve direct run or paragraph formatting first, then character or paragraph style, based-on style chain, table style, document defaults, and theme fallback. Treat East Asian and Latin fonts separately.

## Review workflow

1. Accept a thesis DOCX and the bundled template, or a user-supplied replacement template.
2. Preserve the original thesis and package relationships.
3. Inspect package structure and effective formatting without flattening styles.
4. Classify every finding as `ERROR`, `WARNING`, `MANUAL_REVIEW`, or `PASS`.
5. Add Word comments only at reliable main-document anchors, using the rule's comment template. Combine findings on the same paragraph into one comment while preserving every rule block.
6. Produce two separate DOCX files: the commented thesis copy and the structured Word audit report.
7. Verify both packages as valid OOXML; for the commented copy compare unchanged package parts and structural inventory, and for the report verify the 15-section structure, all 76 RuleResults, complete Finding coverage, comment mapping, statistics, tables, and page-number field.

Minimize OOXML mutation. Preserve fields, bookmarks, hyperlinks, notes, drawings, content controls, tracked changes, comments, section breaks, and relationships unless the user explicitly authorizes a repair that requires changing them.

## Determinism boundary

Use `AUTOMATIC` only for properties reliably derivable from DOCX/OOXML. Use `SEMI_AUTOMATIC` when detection is reliable but context or classification needs human confirmation. Use `MANUAL_REVIEW` for visual balance, true rendered page position, image legibility/resolution when metadata is insufficient, semantic correspondence, and other layout-dependent judgments.

If Word or LibreOffice rendering is unavailable, do not infer visual conformity from XML alone. Report the unexecuted rendering checks explicitly.
