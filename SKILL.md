---
name: desk-sign-template-replacer
description: Replace names in the supplied dining, meeting, or pinyin desk-sign templates while preserving the original borders, text boxes, row heights, print settings, and cut-ready dimensions. Use when the user asks for 餐桌签、吃饭桌签、会议桌签、开会桌签、拼音桌签, provides names or a Word/Excel name list, or asks to reuse these desk-sign templates.
---

# 桌签模板替换器

## Overview

Use the three bundled templates as the format authority. Extract names from direct text or a supplied Word/Excel file, route the request to independent dining, meeting, or pinyin branches, replace only the existing name content, and return editable print-ready file(s). Do not redesign the template, add roles, reorder names, or distribute raw name lists into the skill package.

## Request routing

Route before editing; never let one template branch format the other type:

- `吃饭、就餐、用餐、餐桌签` => run only the `dining` branch.
- `开会、会议、会议桌签` => run only the `meeting` branch.
- `拼音、姓名拼音、拼音桌签、英文姓名桌签` => run only the `pinyin` branch. For Chinese names, output Chinese on the first line and pinyin on the second line exactly like the supplied Word file; for already-Latin foreign names, use its original single-line large-name style.
- `全部、两种、吃饭和会议、会议和吃饭、各做一份` => run both branches separately with the same extracted names, then return the two output files as one grouped result. The dining file is produced first and the meeting file second.
- `全部三种、三个桌签、吃饭会议和拼音` => run `all-three`: dining first, meeting second, pinyin third; return three separate editable files.
- If the request contains both types but does not clearly ask for both, ask one short question. If it contains neither type, use an explicitly supplied template as the fallback; otherwise ask.

When the request asks for both, use the batch dispatcher so each branch performs its own capacity and layout checks:

```bash
PYTHON_BIN="/Users/davidliu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
"$PYTHON_BIN" scripts/desk_signs.py batch \
  --mode all \
  --dining-template assets/templates/餐桌签.docx \
  --meeting-template assets/templates/会议桌签.xlsx \
  --names-file "/tmp/desk-sign-names.txt" \
  --output-dir "/absolute/path/output" \
  --output-stem "桌签-本次名单"
```

The batch command extracts/uses the names once, invokes the two independent replacement functions, and prints both paths. Do not merge the two formats into one document.

For all three branches, use `--mode all-three` and add the one-page pinyin template:

```bash
"$PYTHON_BIN" scripts/desk_signs.py batch \
  --mode all-three \
  --dining-template assets/templates/餐桌签.docx \
  --meeting-template assets/templates/会议桌签.xlsx \
  --pinyin-template assets/templates/拼音桌签.docx \
  --names-file "/tmp/desk-sign-names.txt" \
  --output-dir "/absolute/path/output" \
  --output-stem "桌签-本次名单"
```

For a list that exceeds either template's capacity, add `--split-overflow`. Keep the original name order and create numbered files such as `第01批` and `第02批`; never add rows or shrink the template:

```bash
"$PYTHON_BIN" scripts/desk_signs.py batch \
  --mode all \
  --dining-template assets/templates/餐桌签.docx \
  --meeting-template assets/templates/会议桌签.xlsx \
  --names-file "/tmp/desk-sign-names.txt" \
  --output-dir "/absolute/path/output" \
  --output-stem "桌签-本次名单" \
  --split-overflow
```

## Template selection

- Treat `吃饭、就餐、用餐、餐桌签` as `dining`; use `assets/templates/餐桌签.docx`. If the user explicitly supplies a legacy `.doc`, convert it loss-minimally in a temporary directory first.
- Treat `开会、会议、会议桌签` as `meeting`; use `assets/templates/会议桌签.xlsx`. If the user explicitly supplies a legacy `.xls`, convert it loss-minimally in a temporary directory first.
- Treat `拼音、姓名拼音、拼音桌签、英文姓名桌签` as `pinyin`; use `assets/templates/拼音桌签.docx`, which contains only page 1 of the user-confirmed source, and output `.docx`.
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
- In pinyin mode, Chinese names are converted automatically to toneless surname-first pinyin and kept together with the Chinese line: `张三 / Zhang  San`, `李佳琦 / Li Jiaqi`, `欧阳娜娜 / Ouyang Nana`. Match the supplied file's spacing: a two-character Chinese name uses two ordinary spaces between its characters, and its pinyin uses two ordinary spaces between surname and given name; names with three or more Chinese characters use one pinyin space between surname and the joined given name.
- Automatic pronunciation cannot reliably resolve every polyphonic character. The script contains common surname overrides, but when the requested pronunciation differs, supply the desired Latin spelling directly in the name list (for example `Yue Le`). Never silently guess after the user supplies an explicit Latin spelling.

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
4. If the list exceeds capacity, use `--split-overflow`. Dining and meeting branches split independently because their capacities differ. Pinyin mode repeats its four-person, eight-frame page-1 master automatically and remains one DOCX.
5. Generation is transactional: build all requested branches/batches in a staging directory, verify all of them, and only then publish final files. If any name, font, alignment, or slot check fails, return the error and publish none of the newly generated group.
6. The built-in post-write verification must compare every output slot with the formatted input in exact order. Dining verifies both compatibility copies of every visible text box, two boxes per person, `楷体、65号`, and horizontal/vertical centering. Meeting verifies each name cell, `方正楷体_GBK、120号`, horizontal/vertical centering, A4 portrait, and a manual page break after every two used rows. Pinyin verifies two identical adjacent frames per person, four people/eight frames per A4 page, Chinese-plus-pinyin line order, the original two-character spacing, and the source font sizes.

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
     --template assets/templates/餐桌签.docx \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/餐桌签-替换版.docx"

   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode meeting \
     --template assets/templates/会议桌签.xlsx \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/会议桌签-替换版.xlsx"

   "$PYTHON_BIN" scripts/desk_signs.py replace \
     --mode pinyin \
     --template assets/templates/拼音桌签.docx \
     --names-file "/tmp/desk-sign-names.txt" \
     --output "/absolute/path/拼音桌签-替换版.docx"
   ```

   For `全部` requests, use the `batch` command above rather than manually placing dining names into the meeting template or vice versa.

4. Confirm that the command reports successful replacement and verification. The script already checks that output contains exactly the supplied names in order, that no old template name remains in edited name fields, that dining compatibility text boxes agree, and that the required fonts/alignment remain present. Independently confirm that the file opens and that original row heights, borders, page setup, and dimensions remain unchanged.
5. Render for visual QA when the user asks for a print-ready deliverable. Use the document renderer for `.docx`; use LibreOffice PDF export or the spreadsheet renderer for `.xlsx`. Check that borders are complete, text is centered, names are not clipped, and the last page remains cuttable. For dining signs, when WPS/Word is available, open the output there and verify the actual text-box position: the font must be `楷体、65号`, the glyphs must be vertically centered, and a two-character name's first/second glyphs must align with the first/third glyphs of a three-character name. For meeting signs, verify the corresponding centered spreadsheet cells use the font `方正楷体_GBK、120号` and the same first/third-position rule. For pinyin signs, WPS is the visual authority because LibreOffice omits the Chinese line from these legacy text boxes; verify in WPS that each page has four people/eight frames, Chinese is above pinyin, and the status bar page count matches `ceil(人数/4)`. Still run structural OOXML verification for both compatibility copies.
6. Return the editable file for the selected branch. For `全部`, return both editable files together and label which is吃饭桌签、which is会议桌签. Offer PDFs only when the user asks for print-only copies.

## Fidelity and capacity rules

- Dining template: treat two adjacent visible rectangles as one person's pair and write the same formatted name into both. Keep its borders, dimensions, page geometry, and explicitly use `楷体、65号` inside each box while centering the name horizontally and vertically. A two-character name must use one full-width space so its first and second characters align with the first and third positions of a three-character name, such as `张　三`; a three-character name is written without an internal space, such as `李佳琦`. The template has 92 people slots; blank unused boxes on the last used page and remove trailing unused template pages so a short list is directly printable.
- Meeting template: patch only the existing name cells in `Sheet1!A1:A34`; keep all other sheets, cell styles, used row heights, borders, margins, orientation, and scaling. Preserve the user-confirmed template font `方正楷体_GBK、120号` on the name cells even when legacy conversion reports a different font. Remove trailing unused name rows from the printable range and insert a manual page break after every two used rows, so each A4 portrait sheet contains exactly two bordered name frames and a partial final batch does not print blank pages. Apply the same name alignment rule as dining signs: use one full-width space for a two-character name, such as `张　三`, so its two glyphs occupy the first and third positions; leave three-character names such as `李佳琦` unchanged. The template has 34 slots and one person per cell.
- Pinyin template: the bundled DOCX is a sanitized copy of only the first physical page as displayed by WPS in the supplied legacy Word file. Preserve its A4 portrait page, eight bordered rectangles in two columns and four rows, exact anchor positions, and paragraph geometry. Each person occupies two adjacent identical frames; each page contains four people. Chinese names use two centered lines: `楷体 56磅` Chinese above `楷体 36磅` pinyin, both regular weight. Two-character Chinese names and their two-part pinyin use the source's two ordinary spaces; three-character names keep the Chinese characters together and use one pinyin space. Already-Latin foreign names use the source's single centered `楷体 65磅、加粗` line. For more than four names, clone only this first-page layout and insert a page break between copies. Keep the unused pairs on the last page blank without removing their cut borders. Reduce only an overlong Latin/pinyin line enough to prevent clipping, never below 30 points; do not alter the Chinese 56-point line.
- If the list exceeds the selected template capacity, stop with a clear error unless `--split-overflow` is active. With that option, split by the detected template capacity into numbered files without reordering names; do not add rows or shrink text.
- Keep Chinese internal spacing exactly as supplied when it is meaningful (for example, `李  波`). Do not add `班长：` or other role labels unless the user explicitly includes them.

## Script contract

`scripts/desk_signs.py` uses ZIP/XML-level edits for DOCX/XLSX values so it does not recreate tables or restyle the workbook. It accepts the bundled legacy templates, converts them only in a temporary directory, audits risky name values, preserves real spreadsheet column positions, splits overflow lists, repeats the pinyin page-1 master, and verifies every generated slot before atomic publication. Chinese-to-pinyin conversion uses the bundled Swift helper and macOS Foundation, so no network lookup or third-party pinyin package is needed. It requires the bundled Python runtime, `lxml`, macOS `/usr/bin/swift`, and a usable `soffice` for legacy conversion.
