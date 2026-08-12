---
name: desk-sign-template-replacer
description: Replace names in the supplied dining or meeting desk-sign templates while preserving the original borders, text boxes, row heights, print settings, and cut-ready dimensions. Use when the user asks for 餐桌签、吃饭桌签、会议桌签、开会桌签, provides names or a Word name list, or asks to reuse these desk-sign templates.
---

# 桌签模板替换器

## Overview

Use the two bundled templates as the format authority. Extract names from direct text or a supplied `.doc`/`.docx`, route the request to an independent dining branch, meeting branch, or both branches, replace only the existing name content, and return editable print-ready file(s). Do not redesign the template, add roles, reorder names, or distribute raw name lists into the skill package.

## Request routing

Route before editing; never let one template branch format the other type:

- `吃饭、就餐、用餐、餐桌签` => run only the `dining` branch.
- `开会、会议、会议桌签` => run only the `meeting` branch.
- `全部、两种、吃饭和会议、会议和吃饭、各做一份` => run both branches separately with the same extracted names, then return the two output files as one grouped result. The dining file is produced first and the meeting file second.
- If the request contains both types but does not clearly ask for both, ask one short question. If it contains neither type, use an explicitly supplied template as the fallback; otherwise ask.

When the request asks for both, use the batch dispatcher so each branch performs its own capacity and layout checks:

```bash
PYTHON_BIN="/Users/davidliu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
"$PYTHON_BIN" scripts/desk_signs.py batch \
  --mode all \
  --dining-template assets/templates/餐桌签.doc \
  --meeting-template assets/templates/会议桌签.xls \
  --names-file "/tmp/desk-sign-names.txt" \
  --output-dir "/absolute/path/output" \
  --output-stem "桌签-本次名单"
```

The batch command extracts/uses the names once, invokes the two independent replacement functions, and prints both paths. Do not merge the two formats into one document.

## Template selection

- Treat `吃饭、就餐、用餐、餐桌签` as `dining`; use `assets/templates/餐桌签.doc` and output `.docx` after a loss-minimizing legacy conversion.
- Treat `开会、会议、会议桌签` as `meeting`; use `assets/templates/会议桌签.xls` and output `.xlsx` after a loss-minimizing legacy conversion.
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

3. Replace names using the selected branch. The script converts `.doc`/`.xls` in a temporary LibreOffice profile before editing, then writes only the requested output file:

   ```bash
   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode dining \
     --template assets/templates/餐桌签.doc \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/餐桌签-替换版.docx"

   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode meeting \
     --template assets/templates/会议桌签.xls \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/会议桌签-替换版.xlsx"
   ```

   For `全部` requests, use the `batch` command above rather than manually placing dining names into the meeting template or vice versa.

4. Verify that the output contains exactly the supplied names in order, that no old template name remains in the edited name fields, and that the file opens. For dining output, inspect the patched `w:txbxContent` groups; for meeting output, inspect column A in `Sheet1` and confirm the original row heights, borders, and page setup remain unchanged.
5. Render for visual QA when the user asks for a print-ready deliverable. Use the document renderer for `.docx`; use LibreOffice PDF export or the spreadsheet renderer for `.xlsx`. Check that borders are complete, text is centered, names are not clipped, and the last page remains cuttable. For dining signs, when WPS/Word is available, open the output there and verify the actual text-box position: the glyphs must be vertically centered, and a two-character name's first/second glyphs must align with the first/third glyphs of a three-character name. For meeting signs, verify the corresponding centered spreadsheet cells use the same first/third-position rule. If the legacy Word text boxes render blank in headless LibreOffice, still verify the OOXML text nodes and state the limitation rather than rebuilding the layout.
6. Return the editable file for the selected branch. For `全部`, return both editable files together and label which is吃饭桌签、which is会议桌签. Offer PDFs only when the user asks for print-only copies.

## Fidelity and capacity rules

- Dining template: treat two adjacent visible rectangles as one person's pair and write the same formatted name into both. Keep its borders, dimensions, font, and page geometry, while explicitly centering the name inside each box horizontally and vertically. A two-character name must use one full-width space so its first and second characters align with the first and third positions of a three-character name, such as `张　三`; a three-character name is written without an internal space, such as `李佳琦`. The template has 92 people slots; blank unused boxes on the last used page and remove trailing unused template pages so a short list is directly printable.
- Meeting template: patch only the existing name cells in `Sheet1!A1:A34`; keep all other sheets, cell styles, row heights, borders, margins, orientation, and scaling. Blank unused rows after the supplied list. Apply the same name alignment rule as dining signs: use one full-width space for a two-character name, such as `张　三`, so its two glyphs occupy the first and third positions; leave three-character names such as `李佳琦` unchanged. The template has 34 slots and one person per cell.
- Stop with a clear error if the list exceeds the selected template capacity; do not silently add rows or shrink text.
- Keep Chinese internal spacing exactly as supplied when it is meaningful (for example, `李  波`). Do not add `班长：` or other role labels unless the user explicitly includes them.

## Script contract

`scripts/desk_signs.py` uses ZIP/XML-level edits for DOCX/XLSX values so it does not recreate tables or restyle the workbook. It accepts the bundled legacy templates and converts them only in a temporary directory. It requires the bundled Python runtime, `lxml`, and a usable `soffice` for legacy conversion.
