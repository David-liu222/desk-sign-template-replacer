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

For a list that exceeds either template's capacity, add `--split-overflow`. Keep the original name order and create numbered files such as `第01批` and `第02批`; never add rows or shrink the template:

```bash
"$PYTHON_BIN" scripts/desk_signs.py batch \
  --mode all \
  --dining-template assets/templates/餐桌签.doc \
  --meeting-template assets/templates/会议桌签.xls \
  --names-file "/tmp/desk-sign-names.txt" \
  --output-dir "/absolute/path/output" \
  --output-stem "桌签-本次名单" \
  --split-overflow
```

## Template selection

- Treat `吃饭、就餐、用餐、餐桌签` as `dining`; use `assets/templates/餐桌签.doc` and output `.docx` after a loss-minimizing legacy conversion.
- Treat `开会、会议、会议桌签` as `meeting`; use `assets/templates/会议桌签.xls` and output `.xlsx` after a loss-minimizing legacy conversion.
- If the user provides an explicit template, use that template and infer the mode from its extension/name unless the user explicitly says otherwise.
- If neither mode nor template is clear, ask one short question before editing. Do not guess because the two templates have different physical layouts.

## Name input

- Accept names typed in the request, a UTF-8 text file, or a Word document containing one name per line/paragraph.
- Before extracting, choose the input branch: if the source is `.xls`/`.xlsx` or a Word document contains a table, use the table branch; otherwise use the established non-table branch.
- In the table branch, search the first 12 rows for a `姓名`/`名字`/`人员` header and read only that physical column. Preserve blank Excel columns when locating cells. If there is no header, accept a single populated column or one uniquely identifiable name column only; if department, position, phone, or another column could also be selected, stop and ask the user to label the name column. Never flatten a multi-column table into names.
- In the non-table branch, keep the established behavior: use `scripts/desk_signs.py extract` to read ordinary paragraphs or text boxes in legacy-converted `.docx` files.
- Preserve input order and repeated names. Strip numbering and labels such as `姓名：`; do not silently deduplicate.
- Never silently correct a valid-looking name. The skill can prove that output matches the source exactly, but it cannot know whether the source itself wrote `陈立` when the intended person was `陈莉`; expose the extracted list for human review when the source may be wrong.
- Do not treat the old names in the bundled templates as input names. The templates are only format resources.

## Large-list safeguards

Apply these safeguards whenever the source is a table, the list contains more than 20 entries, or the user says the list has previously produced wrong names or formatting:

1. Generate a UTF-8 BOM audit CSV before making desk signs:

   ```bash
   "$PYTHON_BIN" scripts/desk_signs.py audit \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/tmp/桌签名单核对.csv"
   ```

2. Check the reported count against the source and inspect every row whose `检查结果` is not `通过`. Review `重复次数` rather than deleting repeated names: a repeat may be intentional.
3. Stop before generation for a one-character item, digits, unsupported punctuation, overlong text, a likely header/summary field, or a likely department/unit value. Do not guess the intended spelling.
4. If the list exceeds capacity, use `--split-overflow`. Dining and meeting branches split independently because their capacities differ.
5. Generation is transactional: build all requested branches/batches in a staging directory, verify all of them, and only then publish final files. If any name, font, alignment, or slot check fails, return the error and publish none of the newly generated group.
6. The built-in post-write verification must compare every output slot with the formatted input in exact order. Dining verifies both compatibility copies of every visible text box, two boxes per person, `楷体、65号`, and horizontal/vertical centering. Meeting verifies each name cell, `方正楷体_GBK、120号`, and horizontal/vertical centering.

## Workflow

1. Resolve the template and names. For direct text, write a temporary UTF-8 file with one name per line. Keep this temporary file outside the skill directory.
2. Run the extractor when the source is a Word document:

   ```bash
   PYTHON_BIN="/Users/davidliu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
   "$PYTHON_BIN" scripts/desk_signs.py extract --source "/absolute/path/names.docx" --output "/tmp/desk-sign-names.txt"
   ```

3. For a table or a list over 20 entries, run the audit command above. Then replace names using the selected branch. The script converts `.doc`/`.xls` in a temporary LibreOffice profile, builds in a staging directory, performs exact post-write checks, and publishes only verified output:

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

4. Confirm that the command reports successful replacement and verification. The script already checks that output contains exactly the supplied names in order, that no old template name remains in edited name fields, that dining compatibility text boxes agree, and that the required fonts/alignment remain present. Independently confirm that the file opens and that original row heights, borders, page setup, and dimensions remain unchanged.
5. Render for visual QA when the user asks for a print-ready deliverable. Use the document renderer for `.docx`; use LibreOffice PDF export or the spreadsheet renderer for `.xlsx`. Check that borders are complete, text is centered, names are not clipped, and the last page remains cuttable. For dining signs, when WPS/Word is available, open the output there and verify the actual text-box position: the font must be `楷体、65号`, the glyphs must be vertically centered, and a two-character name's first/second glyphs must align with the first/third glyphs of a three-character name. For meeting signs, verify the corresponding centered spreadsheet cells use the font `方正楷体_GBK、120号` and the same first/third-position rule. If the legacy Word text boxes render blank in headless LibreOffice, still verify the OOXML text nodes and state the limitation rather than rebuilding the layout.
6. Return the editable file for the selected branch. For `全部`, return both editable files together and label which is吃饭桌签、which is会议桌签. Offer PDFs only when the user asks for print-only copies.

## Fidelity and capacity rules

- Dining template: treat two adjacent visible rectangles as one person's pair and write the same formatted name into both. Keep its borders, dimensions, page geometry, and explicitly use `楷体、65号` inside each box while centering the name horizontally and vertically. A two-character name must use one full-width space so its first and second characters align with the first and third positions of a three-character name, such as `张　三`; a three-character name is written without an internal space, such as `李佳琦`. The template has 92 people slots; blank unused boxes on the last used page and remove trailing unused template pages so a short list is directly printable.
- Meeting template: patch only the existing name cells in `Sheet1!A1:A34`; keep all other sheets, cell styles, row heights, borders, margins, orientation, and scaling. Preserve the user-confirmed template font `方正楷体_GBK、120号` on the name cells even when legacy conversion reports a different font. Blank unused rows after the supplied list. Apply the same name alignment rule as dining signs: use one full-width space for a two-character name, such as `张　三`, so its two glyphs occupy the first and third positions; leave three-character names such as `李佳琦` unchanged. The template has 34 slots and one person per cell.
- If the list exceeds the selected template capacity, stop with a clear error unless `--split-overflow` is active. With that option, split by the detected template capacity into numbered files without reordering names; do not add rows or shrink text.
- Keep Chinese internal spacing exactly as supplied when it is meaningful (for example, `李  波`). Do not add `班长：` or other role labels unless the user explicitly includes them.

## Script contract

`scripts/desk_signs.py` uses ZIP/XML-level edits for DOCX/XLSX values so it does not recreate tables or restyle the workbook. It accepts the bundled legacy templates, converts them only in a temporary directory, audits risky name values, preserves real spreadsheet column positions, splits overflow lists, and verifies every generated slot before atomic publication. It requires the bundled Python runtime, `lxml`, and a usable `soffice` for legacy conversion.
