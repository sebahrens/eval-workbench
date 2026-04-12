# Formatter-Family Coverage Matrix

This document defines the supported output families, their representative formatters, canary-placement patterns, noise-applicability, and any intentionally excluded families with rationale. It serves as the reference for `synth-data-8ok.5` (controlled document noise) and `synth-data-8ok.9` (canary-preservation tests).

## Supported Formatter Families

### 1. XLSX (openpyxl)

| Attribute | Value |
|---|---|
| **Library** | openpyxl |
| **Canary location** | Custom document property `description` |
| **Canary embed function** | `embed_canary_xlsx()` in `generator/canaries.py` |
| **Determinism mechanism** | `save_xlsx_deterministic()` — pins ZIP entry dates, `pin_xlsx_dates()` for created/modified timestamps |
| **File count** | ~100+ across all TCs |
| **Representative TCs** | TC-01 (trial balance), TC-05 (AR aging), TC-12 (data room schedules), TC-16 (fee schedule with ERR-013) |
| **Planted errors in family** | ERR-001, ERR-004, ERR-005, ERR-010, ERR-013, ERR-016, ERR-019, ERR-020, ERR-021 |

**Noise applicability**: High priority. XLSX is the dominant output format and carries the most financial data. Noise dimensions:

- **Header perturbation** — column name variations (e.g., "Acct #" vs "Account Number"), extra whitespace, mixed case
- **Date format variation** — cells formatted as `MM/DD/YYYY` vs `YYYY-MM-DD` vs `M/D/YY` vs text-stored dates
- **Number formatting** — parenthetical negatives `(1,234)` vs `-1,234`, inconsistent decimal places
- **Blank rows/columns** — stray empty rows between data regions, merged cells in headers
- **Sheet naming** — inconsistent sheet tab names ("Sheet1" vs descriptive name)

**Canary risk**: Low. Canary is in document properties, which are independent of cell/sheet content. Noise transforms that modify cells or formatting will not affect the canary. Risk exists only if a noise transform recreates the workbook from scratch without copying properties.

---

### 2. DOCX (python-docx)

| Attribute | Value |
|---|---|
| **Library** | python-docx |
| **Canary location** | Core properties `comments` field |
| **Canary embed function** | `embed_canary_docx()` in `generator/canaries.py` |
| **Determinism mechanism** | `save_docx_deterministic()` — pins ZIP entry dates |
| **File count** | ~50+ across all TCs |
| **Representative TCs** | TC-03 (management rep letter), TC-06 (statutory rates), TC-08 (R&D project descriptions), TC-12 (litigation memo), TC-17 (deliverable sections), TC-19 (contracts) |
| **Planted errors in family** | None directly in docx content; some TCs reference docx files in error context |

**Noise applicability**: Medium priority. DOCX files carry narrative content (memos, agreements, descriptions) where noise is structurally different from tabular data. Noise dimensions:

- **Formatting inconsistency** — mixed fonts within a document, inconsistent heading levels, manual bold vs style-based bold
- **Whitespace artifacts** — double spaces, trailing whitespace, inconsistent paragraph spacing
- **OCR-like quirks** — for documents that simulate scanned-then-OCR'd originals (e.g., contracts): `l`/`1` confusion, ligature artifacts, broken hyphens
- **Immaterial missing fields** — optional header fields left blank (e.g., "Prepared by: ___")

**Canary risk**: Low. Canary is in core properties (comments), isolated from document body content. Same risk profile as XLSX — only a full document rebuild without property copy would lose it.

---

### 3. PDF (reportlab + fpdf2)

| Attribute | Value |
|---|---|
| **Libraries** | reportlab (Canvas), fpdf2 (FPDF) |
| **Canary location** | reportlab: PDF metadata `Author`; fpdf2: PDF metadata `Subject` |
| **Canary embed functions** | `embed_canary_pdf_reportlab()`, `embed_canary_pdf_fpdf2()` in `generator/canaries.py` |
| **Determinism mechanism** | reportlab: `invariant=True` + fixed creation date; fpdf2: explicit `creation_date`/`modification_date` |
| **File count** | ~60+ across all TCs |
| **Representative TCs** | TC-01 (signed financials), TC-04 (15 lease PDFs), TC-07 (K-1 forms), TC-09 (TP report), TC-11 (customer contracts), TC-12 (corporate docs, legal agreements, tax returns) |
| **Planted errors in family** | None directly planted in PDF content |

**Noise applicability**: Medium-low priority. PDFs are read-only artifacts simulating signed/scanned documents. Noise dimensions are more limited:

- **Scan-quality simulation** — slightly rotated text, faint background noise, minor alignment shifts (only for documents simulating scanned originals)
- **Font variation** — different fonts between sections of the same document
- **Page numbering inconsistency** — "Page 1 of 5" vs "1/5" vs no page numbers
- **Metadata clutter** — extra metadata fields (producer, creator tool strings) that add noise to metadata extraction

**Canary risk**: Medium. The canary lives in PDF metadata (Author or Subject field). Risks:

- reportlab and fpdf2 use **different metadata fields** (Author vs Subject) — noise transforms must be aware of which library produced the PDF
- A noise transform that rewrites PDF metadata (e.g., adding fake producer strings) could overwrite the canary field
- PDF re-rendering or re-saving through a different library could strip metadata

**Mitigation**: Noise transforms for PDF must treat `Author` and `Subject` fields as reserved. Any metadata-level noise should only touch other fields (e.g., `Producer`, `Creator`, `Keywords`).

---

### 4. CSV (plain text)

| Attribute | Value |
|---|---|
| **Library** | Built-in `csv` module / pandas `to_csv()` |
| **Canary location** | First line as `# CANARY: XXXXXXXX` comment |
| **Canary embed function** | `embed_canary_csv_comment()` in `generator/canaries.py` |
| **Determinism mechanism** | Deterministic by construction (text output, sorted keys) |
| **File count** | ~5 across all TCs |
| **Representative TCs** | TC-02 (bank statement), TC-08 (R&D time records), TC-13 (AP transactions), TC-18 (current-year TB as CSV — format change from prior-year xlsx) |
| **Planted errors in family** | ERR-002 (transposed_digits in TC-02 bank CSV) |

**Noise applicability**: Medium priority despite low file count. CSV is the format most susceptible to real-world messiness. Noise dimensions:

- **Delimiter variation** — some CSVs use semicolons or tabs instead of commas
- **Encoding artifacts** — BOM markers, mixed line endings (`\r\n` vs `\n`)
- **Quoting inconsistency** — some fields unnecessarily quoted, others missing quotes around commas-in-values
- **Header variation** — column names differ from schema (abbreviations, extra spaces, different case)
- **Trailing delimiters** — extra comma at end of rows

**Canary risk**: Medium-high. The canary is a comment line (`# CANARY: ...`) at the top of the file. Risks:

- Many CSV parsers skip comment lines, but not all tools recognize `#` as a comment character
- A noise transform that shuffles lines or strips comments would destroy the canary
- Adding BOM markers before the comment line could make the canary line unparseable

**Mitigation**: Noise transforms must preserve the first line of any CSV file that starts with `# CANARY:`. Comment-line stripping must be excluded from CSV noise profiles.

---

## Intentionally Excluded Families

### JSON (registry and gold standard files)

**Rationale**: JSON files in this project are machine-generated registries (`manifest.json`, `canary_registry.json`, `error_registry.json`) and gold standards (`gold_standards/TC-XX_gold.json`). These are **infrastructure**, not test inputs — the agent under test does not receive them as input files. Applying noise to registries would corrupt the test harness itself. Gold standards must remain pristine for grading.

**Exception**: If a future test case provides JSON as an agent input (e.g., an API response fixture), it would need its own canary/noise strategy. File a bead at that point.

### Markdown (prompt.md and expected_behavior.md)

**Rationale**: Markdown files are the test case prompts and expected-behavior descriptions. They are instructions to the agent, not data for analysis. Adding noise to instructions would test the agent's ability to parse noisy prompts, which is out of scope for this suite (the suite tests professional services workload handling, not prompt robustness).

### YAML (config.yaml, rubrics.yaml)

**Rationale**: YAML files are generator configuration and scoring rubrics — internal infrastructure. Not presented to the agent under test.

---

## Coverage Summary

| Family | File Count | Canary Method | Canary Risk | Noise Priority | Noise Dimensions |
|---|---|---|---|---|---|
| **XLSX** | ~100+ | doc property `description` | Low | **High** | Headers, dates, numbers, blanks, sheet names |
| **DOCX** | ~50+ | core property `comments` | Low | **Medium** | Formatting, whitespace, OCR quirks, missing fields |
| **PDF** | ~60+ | metadata `Author`/`Subject` | **Medium** | Medium-low | Scan quality, fonts, page numbers, metadata clutter |
| **CSV** | ~5 | `# CANARY:` comment line | **Medium-high** | **Medium** | Delimiters, encoding, quoting, headers, trailing delimiters |
| JSON | ~25 | None | N/A | Excluded | Infrastructure, not agent input |
| Markdown | ~42 | None | N/A | Excluded | Instructions, not data |
| YAML | ~3 | None | N/A | Excluded | Configuration, not agent input |

## Recommended Pilot Order for synth-data-8ok.5

Based on file count, noise priority, and canary risk:

1. **XLSX** — highest file count, highest noise priority, lowest canary risk. Start here.
2. **CSV** — highest canary risk, good test of comment-line preservation. Pair with XLSX in the pilot.
3. **DOCX** — medium priority, low canary risk. Second wave.
4. **PDF** — most complex (dual library), medium canary risk. Third wave or deferred if the first three families cover the acceptance criteria.

## Noise Invariants (apply to all families)

These rules must hold for any noise transform, regardless of family:

1. **Canary preservation** — the canary must remain findable after noise is applied. "Findable" means the auto-grader's canary-verification logic can extract it from the noisy file using the same method it uses on clean files.
2. **Model fact preservation** — noise must not alter any value that the gold standard depends on. Noise is cosmetic/structural, never semantic.
3. **Planted error preservation** — if a file contains a planted error, the error must still be detectable at its registered location after noise. Noise must not mask, relocate, or accidentally "fix" a planted error.
4. **Determinism** — noise must be seeded and produce identical output across runs. Use the `ScenarioContext` seed namespace (from `synth-data-ln9.1`) to derive a per-file noise seed.
5. **Canary findability testing** — `synth-data-8ok.9` must verify these invariants for every family in this matrix. A passing noise profile is one where all canary and error checks pass on the noisy output.
