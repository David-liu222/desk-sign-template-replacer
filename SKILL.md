---
name: desk-sign-template-replacer
description: Replace names in the supplied dining or meeting desk-sign templates while preserving the original borders, text boxes, row heights, print settings, and cut-ready dimensions. Use when the user asks for 餐桌签、吃饭桌签、会议桌签、开会桌签, provides names or a Word name list, or asks to reuse these desk-sign templates.
---

# 桌签模板替换器

## Overview

Use the two bundled templates as the format authority. Extract names from direct text or a supplied `.doc`/`.docx`, infer the requested output as dining or meeting, replace only the existing name content, and return an editable print-ready file. Do not redesign the template, add roles, reorder names, or distribute raw name lists into the skill package.

## Template selection

- Treat `吃饭、就餐、用餐、餐桌签` as `dining`; use `assets/templates/餐桌签.docx` and output `.docx`.
- Treat `开会、会议、会议桌签` as `meeting`; use `assets/templates/会议桌签.xlsx` and output `.xlsx`.
- If the user provides an explicit template, use that template and infer the mode from its extension/name unless the user explicitly says otherwise.
- If neither mode nor template is clear, ask one short question before editing. Do not guess because the two templates have different physical layouts.

## Name input

- Accept names typed in the request, a UTF-8 text file, or a Word document containing one name per line/paragraph.
- Before extracting, choose the input branch: if the source is `.xls`/`.xlsx` or a Word document contains a table, use the table branch; otherwise use the established non-table branch.
- In the table branch, read a `姓名`/`名字`/`人员` column when a header exists; otherwise read non-empty cells in row order. Ignore sequence-number cells and stop with a clear error when a table has no valid names.
- In the non-table branch, keep the established behavior: use `scripts/desk_signs.py extract` to read ordinary paragraphs or text boxes in legacy-converted `.docx` files.
- Preserve input order and repeated names. Strip numbering and labels such as `姓名：`; do not silently deduplicate.
- Do not treat the old names in the bundled templates as input names. The templates are only format resources.

## Workflow

1. Resolve the template and names. For direct text, write a temporary UTF-8 file with one name per line. Keep this temporary file outside the skill directory.
2. Run the extractor when the source is a Word document:

   ```bash
   PYTHON_BIN="/Users/davidliu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
   "$PYTHON_BIN" scripts/desk_signs.py extract --source "/absolute/path/names.docx" --output "/tmp/desk-sign-names.txt"
   ```

3. Replace names using the bundled template. The script converts `.doc`/`.xls` in a temporary LibreOffice profile before editing, then writes only the requested output file:

   ```bash
   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode dining \
     --template assets/templates/餐桌签.docx \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/餐桌签-替换版.docx"

   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode meeting \
     --template assets/templates/会议桌签.xlsx \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/会议桌签-替换版.xlsx"
   ```

4. Verify that the output contains exactly the supplied names in order, that no old template name remains in the edited name fields, and that the file opens. For dining output, inspect the patched `w:txbxContent` groups; for meeting output, inspect column A in `Sheet1` and confirm the original row heights, borders, and page setup remain unchanged.
5. Render for visual QA when the user asks for a print-ready deliverable. Use the document renderer for `.docx`; use LibreOffice PDF export or the spreadsheet renderer for `.xlsx`. Check that borders are complete, text is centered, names are not clipped, and the last page remains cuttable. For dining signs, when WPS/Word is available, open the output there and verify the actual text-box position: the glyphs must be vertically centered, and a two-character name's first/second glyphs must align with the first/third glyphs of a three-character name. For meeting signs, verify the corresponding centered spreadsheet cells use the same first/third-position rule. If the legacy Word text boxes render blank in headless LibreOffice, still verify the OOXML text nodes and state the limitation rather than rebuilding the layout.
6. Return the editable `.docx` or `.xlsx`. Offer a PDF only when the user asks for a print-only copy.

## Fidelity and capacity rules

- Dining template: treat two adjacent visible rectangles as one person's pair and write the same formatted name into both. Keep its borders, dimensions, font, and page geometry, while explicitly centering the name inside each box horizontally and vertically. A two-character name must use one full-width space so its first and second characters align with the first and third positions of a three-character name, such as `张　三`; a three-character name is written without an internal space, such as `李佳琦`. The template has 92 people slots; blank unused boxes on the last used page and remove trailing unused template pages so a short list is directly printable.
- Meeting template: patch only the existing name cells in `Sheet1!A1:A34`; keep all other sheets, cell styles, row heights, borders, margins, orientation, and scaling. Blank unused rows after the supplied list. Apply the same name alignment rule as dining signs: use one full-width space for a two-character name, such as `张　三`, so its two glyphs occupy the first and third positions; leave three-character names such as `李佳琦` unchanged. The template has 34 slots and one person per cell.
- Stop with a clear error if the list exceeds the selected template capacity; do not silently add rows or shrink text.
- Keep Chinese internal spacing exactly as supplied when it is meaningful (for example, `李  波`). Do not add `班长：` or other role labels unless the user explicitly includes them.

## Script contract

`scripts/desk_signs.py` uses ZIP/XML-level edits for DOCX/XLSX values so it does not recreate tables or restyle the workbook. It accepts the bundled legacy templates and converts them only in a temporary directory. It requires the bundled Python runtime, `lxml`, and a usable `soffice` for legacy conversion.
