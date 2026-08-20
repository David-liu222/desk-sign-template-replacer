#!/usr/bin/env python3
"""Extract names and replace only name values in the supplied desk-sign templates."""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter
from copy import deepcopy
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPS = f"{{{WPS_NS}}}"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MAIN = f"{{{MAIN_NS}}}"
XML_NS = "http://www.w3.org/XML/1998/namespace"
DINING_FONT = "楷体"
DINING_FONT_HALF_POINTS = "130"
MEETING_FONT = "方正楷体_GBK"
MEETING_FONT_POINTS = "120"
PINYIN_FONT = "楷体"
PINYIN_LATIN_HALF_POINTS = "130"
PINYIN_CHINESE_HALF_POINTS = "112"
PINYIN_ROMAN_HALF_POINTS = "72"
PINYIN_PEOPLE_PER_PAGE = 4
PINYIN_SLOTS_PER_PERSON = 2
PINYIN_INNER_WIDTH_POINTS = 230.0
COMPOUND_SURNAMES = {
    "欧阳",
    "太史",
    "端木",
    "上官",
    "司马",
    "东方",
    "独孤",
    "南宫",
    "万俟",
    "闻人",
    "夏侯",
    "诸葛",
    "尉迟",
    "公羊",
    "赫连",
    "澹台",
    "皇甫",
    "宗政",
    "濮阳",
    "公冶",
    "太叔",
    "申屠",
    "公孙",
    "慕容",
    "仲孙",
    "钟离",
    "长孙",
    "宇文",
    "司徒",
    "鲜于",
    "司空",
    "闾丘",
    "子车",
    "亓官",
    "司寇",
    "巫马",
    "公西",
    "颛孙",
    "壤驷",
    "公良",
    "漆雕",
    "乐正",
    "宰父",
    "谷梁",
    "拓跋",
    "夹谷",
    "轩辕",
    "令狐",
    "段干",
    "百里",
    "呼延",
    "东郭",
    "南门",
    "羊舌",
    "微生",
    "公户",
    "公玉",
    "公仪",
    "梁丘",
    "公仲",
    "公上",
    "公门",
    "公山",
    "公坚",
    "左丘",
    "公伯",
    "西门",
    "公祖",
    "第五",
    "公乘",
    "贯丘",
    "公皙",
    "南荣",
    "东里",
    "东宫",
    "仲长",
    "子书",
    "子桑",
    "即墨",
    "达奚",
    "褚师",
}
SURNAME_PINYIN_OVERRIDES = {
    "单": "Shan",
    "曾": "Zeng",
    "解": "Xie",
    "仇": "Qiu",
    "区": "Ou",
    "查": "Zha",
    "朴": "Piao",
    "乐": "Yue",
    "重": "Chong",
    "翟": "Zhai",
    "折": "She",
    "黑": "He",
}
HEADER_CANDIDATES = {"姓名", "名字", "人员", "人员姓名", "姓名名单"}
NON_NAME_TERMS = {
    "序号",
    "编号",
    "部门",
    "单位",
    "岗位",
    "职务",
    "备注",
    "电话",
    "手机号",
    "身份证号",
    "合计",
    "总计",
    "人数",
}


def soffice_path(explicit: str | None = None) -> str:
    candidates = [
        explicit,
        os.environ.get("SOFFICE"),
        shutil.which("soffice"),
        "/Users/davidliu/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    raise RuntimeError("找不到 soffice；请设置 SOFFICE 或传入 --soffice。")


def convert_legacy(source: Path, target_suffix: str, explicit_soffice: str | None) -> Path:
    if source.suffix.lower() == target_suffix:
        return source
    if source.suffix.lower() not in {".doc", ".xls"}:
        raise ValueError(f"不支持的模板格式：{source}")
    temp_root = Path(tempfile.mkdtemp(prefix="desk-sign-convert-"))
    out_dir = temp_root / "out"
    out_dir.mkdir()
    profile = temp_root / "lo-profile"
    profile.mkdir()
    cmd = [
        soffice_path(explicit_soffice),
        "--headless",
        f"-env:UserInstallation={profile.as_uri()}",
        "--convert-to",
        "docx" if target_suffix == ".docx" else "xlsx",
        "--outdir",
        str(out_dir),
        str(source),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "LibreOffice 转换失败")
    converted = out_dir / (source.stem + target_suffix)
    if not converted.exists():
        raise RuntimeError(f"LibreOffice 未生成预期文件：{converted}")
    return converted


def clean_name(value: str) -> str:
    value = value.replace("\u3000", " ").strip()
    value = re.sub(r"^\s*(?:\d+[.、)）]|[-•])\s*", "", value)
    value = re.sub(r"^\s*(?:姓名|名单|人员名单)\s*[:：]\s*", "", value)
    if re.search(r"^(?:姓名|名单|人员名单)$", value):
        return ""
    if re.fullmatch(r"\d+", value):
        return ""
    return value.strip("，,；;、|\t ")


def name_issues(value: str) -> list[str]:
    """Return reasons why a value is unsafe to write as a person's name."""
    compact = re.sub(r"\s+", "", value)
    issues: list[str] = []
    if not compact:
        return ["空姓名"]
    if compact in HEADER_CANDIDATES or compact in NON_NAME_TERMS:
        issues.append("疑似表头或汇总字段")
    if re.search(r"\d", compact):
        issues.append("含数字")
    if re.search(r"[:：,，;；|/\\]", value):
        issues.append("含分隔符或字段标记")
    if not re.fullmatch(r"[\u3400-\u9fffA-Za-z·•'’\-\s]+", value):
        issues.append("含不支持的字符")
    if len(compact) == 1:
        issues.append("仅1个字符，需人工确认")
    if re.search(r"[\u3400-\u9fff]", compact) and len(compact) > 12:
        issues.append("中文姓名长度超过12个字符")
    if not re.search(r"[\u3400-\u9fff]", compact) and len(compact) > 30:
        issues.append("拼音或拉丁字母姓名长度超过30个字符")
    if len(compact) >= 3 and compact.endswith(
        ("部", "科", "处", "室", "组", "队", "班", "部门", "公司", "煤业", "矿业", "办公室", "中心")
    ):
        issues.append("疑似部门或单位名称")
    return list(dict.fromkeys(issues))


def validate_names(names: list[str]) -> None:
    problems = [(index + 1, name, name_issues(name)) for index, name in enumerate(names)]
    problems = [item for item in problems if item[2]]
    if problems:
        detail = "；".join(
            f"第{index}项“{name}”（{'、'.join(issues)}）" for index, name, issues in problems[:12]
        )
        suffix = f"；另有 {len(problems) - 12} 项" if len(problems) > 12 else ""
        raise ValueError(f"名单预检未通过：{detail}{suffix}。请修正名单后再生成，系统不会猜测姓名。")


def write_name_audit(names: list[str], output: Path) -> None:
    """Write a deterministic review sheet without modifying or deduplicating names."""
    output.parent.mkdir(parents=True, exist_ok=True)
    counts = Counter(names)
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["序号", "原始姓名", "排版姓名", "字符数", "重复次数", "检查结果"])
        for index, name in enumerate(names, start=1):
            issues = name_issues(name)
            writer.writerow(
                [
                    index,
                    name,
                    format_aligned_name(name),
                    len(re.sub(r"\s+", "", name)),
                    counts[name],
                    "；".join(issues) if issues else "通过",
                ]
            )


def format_aligned_name(value: str) -> str:
    """Align two-character names to the first/third positions of three-character names."""
    compact = re.sub(r"\s+", "", value)
    if len(compact) == 2:
        # Use one full-width ideographic space so the two glyphs occupy the
        # first and third character positions of a three-character name.
        return f"{compact[0]}\u3000{compact[1]}"
    if len(compact) == 3:
        return compact
    return value.strip()


def split_text(text: str) -> list[str]:
    parts: list[str] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Preserve internal spaces in names; split only on list delimiters.
        for item in re.split(r"[、,，;；|]+", line):
            item = clean_name(item)
            if item:
                parts.append(item)
    return parts


def visual_order(items: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]:
    """Order anchored shapes top-to-bottom, then left-to-right within each paragraph."""
    ordered: list[tuple[int, int, int, object]] = []
    for paragraph_index in sorted({item[0] for item in items}):
        paragraph_items = sorted(
            (item for item in items if item[0] == paragraph_index),
            key=lambda item: (item[1], item[2]),
        )
        rows: list[list[tuple[int, int, int, object]]] = []
        for item in paragraph_items:
            if not rows or item[1] - rows[-1][-1][1] > 100000:
                rows.append([item])
            else:
                rows[-1].append(item)
        for row in rows:
            ordered.extend(sorted(row, key=lambda item: item[2]))
    return ordered


def table_rows_from_docx(root: etree._Element, ns: dict[str, str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for table in root.xpath("./w:body/w:tbl", namespaces=ns):
        for row in table.xpath("./w:tr", namespaces=ns):
            cells = []
            for cell in row.xpath("./w:tc", namespaces=ns):
                cells.append("".join(cell.xpath(".//w:t/text()", namespaces=ns)).strip())
            if any(cells):
                rows.append(cells)
    return rows


def names_from_table_rows(rows: list[list[str]]) -> list[str]:
    if not rows:
        return []
    header_index: int | None = None
    name_column: int | None = None
    for row_index, row in enumerate(rows[:12]):
        for column_index, value in enumerate(row):
            normalized = re.sub(r"\s+", "", value)
            if normalized in HEADER_CANDIDATES:
                header_index = row_index
                name_column = column_index
                break
        if name_column is not None:
            break
    values: list[str] = []
    if name_column is not None:
        for row in rows[header_index + 1 :]:
            if name_column < len(row):
                values.extend(split_text(row[name_column]))
    else:
        populated_rows = [
            [(column_index, value) for column_index, value in enumerate(row) if value.strip()]
            for row in rows
        ]
        populated_rows = [row for row in populated_rows if row]
        if populated_rows and all(len(row) == 1 for row in populated_rows):
            for row in populated_rows:
                values.extend(split_text(row[0][1]))
        else:
            max_columns = max((len(row) for row in rows), default=0)
            candidates: list[tuple[int, int, float]] = []
            for column_index in range(max_columns):
                column_values = [
                    row[column_index]
                    for row in rows
                    if column_index < len(row) and row[column_index].strip()
                ]
                if not column_values:
                    continue
                plausible = 0
                for raw_value in column_values:
                    items = split_text(raw_value)
                    if len(items) == 1 and not name_issues(items[0]):
                        plausible += 1
                ratio = plausible / len(column_values)
                if plausible >= 2 and ratio >= 0.8:
                    candidates.append((column_index, plausible, ratio))
            if len(candidates) != 1:
                raise ValueError(
                    "表格没有明确的“姓名”表头，且无法唯一判断姓名列；"
                    "请把姓名列表头改为“姓名”，系统不会把部门、职务等列当作姓名。"
                )
            selected_column = candidates[0][0]
            for row in rows:
                if selected_column < len(row):
                    values.extend(split_text(row[selected_column]))
    return [name for name in values if name]


def names_from_docx(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": W_NS, "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
    table_names = names_from_table_rows(table_rows_from_docx(root, ns))
    if root.xpath("./w:body/w:tbl", namespaces=ns):
        if not table_names:
            raise ValueError("检测到表格，但未检测到有效姓名；请保留姓名列或每格一个姓名")
        return table_names
    anchors: list[tuple[int, int, int, str]] = []
    ordinary_names: list[str] = []
    for paragraph_index, paragraph in enumerate(root.xpath("./w:body/w:p", namespaces=ns)):
        for anchor in paragraph.xpath(".//wp:anchor", namespaces=ns):
            text = clean_name("".join(anchor.xpath(".//w:txbxContent//w:t/text()", namespaces=ns)))
            if text:
                vertical = int(anchor.xpath("string(wp:positionV/wp:posOffset)", namespaces=ns) or 0)
                horizontal = int(anchor.xpath("string(wp:positionH/wp:posOffset)", namespaces=ns) or 0)
                anchors.append((paragraph_index, vertical, horizontal, text))
        if not paragraph.xpath(".//w:txbxContent", namespaces=ns):
            text = "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
            ordinary_names.extend(split_text(text))
    if anchors:
        return [item[3] for item in visual_order(anchors)]
    return ordinary_names


def names_from_xlsx(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    shared: list[str] = []
    if "xl/sharedStrings.xml" in members:
        shared_root = etree.fromstring(members["xl/sharedStrings.xml"])
        shared = ["".join(item.xpath(".//main:t/text()", namespaces={"main": MAIN_NS})) for item in shared_root]
    worksheet_name = next((name for name in members if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)), None)
    if worksheet_name is None:
        raise ValueError("表格中没有可读取的工作表")
    root = etree.fromstring(members[worksheet_name])
    rows: list[list[str]] = []
    for row in root.xpath(".//main:sheetData/main:row", namespaces={"main": MAIN_NS}):
        indexed_values: dict[int, str] = {}
        for cell in row.xpath("./main:c", namespaces={"main": MAIN_NS}):
            value = "".join(cell.xpath(".//main:t/text()", namespaces={"main": MAIN_NS}))
            if not value:
                numeric = cell.find(f"{MAIN}v")
                value = numeric.text if numeric is not None and numeric.text else ""
            if cell.get("t") == "s" and value.isdigit() and int(value) < len(shared):
                value = shared[int(value)]
            reference = cell.get("r", "")
            match = re.match(r"([A-Z]+)", reference)
            if match:
                column_index = 0
                for letter in match.group(1):
                    column_index = column_index * 26 + ord(letter) - ord("A") + 1
                indexed_values[column_index - 1] = value
        values = [""] * (max(indexed_values, default=-1) + 1)
        for column_index, value in indexed_values.items():
            values[column_index] = value
        if any(values):
            rows.append(values)
    names = names_from_table_rows(rows)
    if not names:
        raise ValueError("检测到表格，但未检测到有效姓名；请保留姓名列或每格一个姓名")
    return names


def extract_names(source: Path, explicit_soffice: str | None = None) -> list[str]:
    suffix = source.suffix.lower()
    if suffix == ".doc":
        converted = convert_legacy(source, ".docx", explicit_soffice)
        return names_from_docx(converted)
    if suffix == ".docx":
        return names_from_docx(source)
    if suffix == ".xls":
        converted = convert_legacy(source, ".xlsx", explicit_soffice)
        return names_from_xlsx(converted)
    if suffix == ".xlsx":
        return names_from_xlsx(source)
    if suffix in {".txt", ".csv", ".tsv"}:
        return split_text(source.read_text(encoding="utf-8"))
    raise ValueError(f"姓名来源必须是 Word、Excel 或 UTF-8 文本文件：{source}")


def read_names(args: argparse.Namespace, *, validate: bool = True) -> list[str]:
    if args.names_file:
        names = extract_names(Path(args.names_file), args.soffice)
    elif args.names is not None:
        names = split_text(args.names)
    else:
        raise ValueError("必须提供 --names-file 或 --names")
    names = [clean_name(name) for name in names]
    names = [name for name in names if name]
    if not names:
        raise ValueError("未检测到有效姓名")
    if validate:
        validate_names(names)
    return names


def center_textbox(shape: etree._Element, box: etree._Element) -> None:
    """Center the name inside the legacy text box in both directions."""
    for paragraph in box.xpath("./w:p", namespaces={"w": W_NS}):
        ppr = paragraph.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            paragraph.insert(0, ppr)
        jc = ppr.find(f"{W}jc")
        if jc is None:
            jc = etree.SubElement(ppr, f"{W}jc")
        jc.set(f"{W}val", "center")
        text_alignment = ppr.find(f"{W}textAlignment")
        if text_alignment is None:
            text_alignment = etree.SubElement(ppr, f"{W}textAlignment")
        text_alignment.set(f"{W}val", "center")
        # WPS keeps legacy text-box paragraphs slightly above the visual center
        # even when bodyPr is set to ctr. Add the template-safe top spacing that
        # brings the glyphs to the visual center without changing the box.
        spacing = ppr.find(f"{W}spacing")
        if spacing is None:
            spacing = etree.SubElement(ppr, f"{W}spacing")
        spacing.set(f"{W}before", "600")
        spacing.set(f"{W}after", "0")
    for body_pr in shape.xpath(".//wps:bodyPr", namespaces={"wps": WPS_NS}):
        body_pr.set("anchor", "ctr")


def set_textbox_text(shape: etree._Element, box: etree._Element, value: str) -> None:
    center_textbox(shape, box)
    for run in box.xpath(".//w:r", namespaces={"w": W_NS}):
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            run.insert(0, rpr)
        fonts = rpr.find(f"{W}rFonts")
        if fonts is None:
            fonts = etree.SubElement(rpr, f"{W}rFonts")
        for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(f"{W}{attribute}", DINING_FONT)
        for tag in ("sz", "szCs"):
            size = rpr.find(f"{W}{tag}")
            if size is None:
                size = etree.SubElement(rpr, f"{W}{tag}")
            size.set(f"{W}val", DINING_FONT_HALF_POINTS)
    text_nodes = box.xpath(".//w:t", namespaces={"w": W_NS})
    if not text_nodes:
        paragraph = box.find(f".//{W}p")
        if paragraph is None:
            return
        run = etree.SubElement(paragraph, f"{W}r")
        text_nodes = [etree.SubElement(run, f"{W}t")]
    text_nodes[0].text = value
    if value.startswith(" ") or value.endswith(" "):
        text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
    else:
        text_nodes[0].attrib.pop(f"{{{XML_NS}}}space", None)
    for node in text_nodes[1:]:
        node.text = ""
        node.attrib.pop(f"{{{XML_NS}}}space", None)


def collect_dining_slots(root: etree._Element) -> list[tuple[int, int, int, list[etree._Element]]]:
    """Pair each modern shape with its legacy fallback and return visual slot order."""
    ns = {"w": W_NS, "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
    slots: list[tuple[int, int, int, list[etree._Element]]] = []
    for paragraph_index, paragraph in enumerate(root.xpath("./w:body/w:p", namespaces=ns)):
        current_anchor: etree._Element | None = None
        shape_items: list[tuple[str, etree._Element]] = []
        for element in paragraph.iter():
            if element.tag == f"{{{ns['wp']}}}anchor":
                shape_items.append(("anchor", element))
            elif element.tag == f"{W}pict":
                shape_items.append(("pict", element))
        for kind, shape in shape_items:
            if kind == "anchor":
                current_anchor = shape
                vertical = int(shape.xpath("string(wp:positionV/wp:posOffset)", namespaces=ns) or 0)
                horizontal = int(shape.xpath("string(wp:positionH/wp:posOffset)", namespaces=ns) or 0)
                slots.append((paragraph_index, vertical, horizontal, [shape]))
            elif current_anchor is not None and slots and slots[-1][3][0] is current_anchor:
                slots[-1][3].append(shape)
    return visual_order(slots)


def set_textbox_text_preserving_style(box: etree._Element, value: str) -> None:
    """Replace only text nodes, leaving the template's paragraph/run geometry intact."""
    text_nodes = box.xpath(".//w:t", namespaces={"w": W_NS})
    if not text_nodes:
        paragraph = box.find(f".//{W}p")
        if paragraph is None:
            raise RuntimeError("拼音桌签模板中的文本框没有可写段落")
        run = etree.SubElement(paragraph, f"{W}r")
        text_nodes = [etree.SubElement(run, f"{W}t")]
    text_nodes[0].text = value
    if value.startswith(" ") or value.endswith(" "):
        text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
    else:
        text_nodes[0].attrib.pop(f"{{{XML_NS}}}space", None)
    for node in text_nodes[1:]:
        node.text = ""
        node.attrib.pop(f"{{{XML_NS}}}space", None)


def replace_textbox_content(shape: etree._Element, prototype_box: etree._Element) -> etree._Element:
    boxes = shape.xpath(".//w:txbxContent", namespaces={"w": W_NS})
    if not boxes:
        raise RuntimeError("拼音桌签框缺少文本框内容")
    old_box = boxes[0]
    new_box = deepcopy(prototype_box)
    parent = old_box.getparent()
    if parent is None:
        raise RuntimeError("拼音桌签文本框结构无效")
    parent.replace(old_box, new_box)
    return new_box


def set_run_half_points(run: etree._Element, half_points: str) -> None:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        rpr = etree.Element(f"{W}rPr")
        run.insert(0, rpr)
    for tag in ("sz", "szCs"):
        size = rpr.find(f"{W}{tag}")
        if size is None:
            size = etree.SubElement(rpr, f"{W}{tag}")
        size.set(f"{W}val", half_points)


def format_bilingual_chinese_name(value: str) -> str:
    compact = re.sub(r"\s+", "", value)
    if len(compact) == 2:
        # The supplied pinyin template uses two ordinary spaces, not an
        # ideographic space, between two Chinese characters.
        return f"{compact[0]}  {compact[1]}"
    return compact


def set_bilingual_box_text(
    box: etree._Element,
    chinese_name: str,
    pinyin_name: str,
    pinyin_half_points: str,
) -> None:
    paragraphs = box.xpath("./w:p", namespaces={"w": W_NS})
    if len(paragraphs) < 2:
        raise RuntimeError("拼音桌签双语框必须包含中文行和拼音行")
    values = [format_bilingual_chinese_name(chinese_name), pinyin_name]
    sizes = [PINYIN_CHINESE_HALF_POINTS, pinyin_half_points]
    for paragraph, value, size in zip(paragraphs[:2], values, sizes, strict=True):
        text_nodes = paragraph.xpath(".//w:t", namespaces={"w": W_NS})
        if not text_nodes:
            run = paragraph.find(f"{W}r")
            if run is None:
                run = etree.SubElement(paragraph, f"{W}r")
            text_nodes = [etree.SubElement(run, f"{W}t")]
        text_nodes[0].text = value
        if "  " in value:
            text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
        else:
            text_nodes[0].attrib.pop(f"{{{XML_NS}}}space", None)
        for node in text_nodes[1:]:
            node.text = ""
            node.attrib.pop(f"{{{XML_NS}}}space", None)
        runs = paragraph.xpath("./w:r", namespaces={"w": W_NS})
        if runs:
            set_run_half_points(runs[0], size)
    for paragraph in paragraphs[2:]:
        for node in paragraph.xpath(".//w:t", namespaces={"w": W_NS}):
            node.text = ""


def set_single_line_box_text(box: etree._Element, value: str, half_points: str) -> None:
    set_textbox_text_preserving_style(box, value)
    runs = box.xpath(".//w:r[w:t]", namespaces={"w": W_NS})
    if not runs:
        raise RuntimeError("拼音桌签外籍姓名框缺少文字运行")
    set_run_half_points(runs[0], half_points)


def renumber_cloned_shapes(root: etree._Element) -> None:
    """Give cloned modern/VML shapes unique IDs so Word can open repeated pages safely."""
    ns = {
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "v": "urn:schemas-microsoft-com:vml",
    }
    for index, element in enumerate(root.xpath(".//wp:docPr", namespaces=ns), start=1):
        element.set("id", str(index))
    for index, element in enumerate(root.xpath(".//*[local-name()='cNvPr']"), start=1):
        element.set("id", str(index))
    for index, element in enumerate(root.xpath(".//v:shape", namespaces=ns), start=1):
        element.set("id", f"pinyin_shape_{index:04d}")


def trim_pinyin_template(
    source: Path,
    output: Path,
    explicit_soffice: str | None,
) -> None:
    """Keep WPS page 1 (eight frames/four people) and sanitize its names."""
    converted = convert_legacy(source, ".docx", explicit_soffice)
    with zipfile.ZipFile(converted) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    root = etree.fromstring(members["word/document.xml"])
    ns = {"w": W_NS}
    slots = collect_dining_slots(root)
    expected_slots = PINYIN_PEOPLE_PER_PAGE * PINYIN_SLOTS_PER_PERSON
    if len(slots) < expected_slots:
        raise RuntimeError(f"拼音桌签第一页少于{expected_slots}个姓名框，无法建立模板")
    first_page_slots = slots[:expected_slots]
    latin_prototypes = [
        deepcopy(shape.xpath(".//w:txbxContent", namespaces=ns)[0])
        for shape in first_page_slots[0][3]
    ]
    bilingual_prototypes = [
        deepcopy(shape.xpath(".//w:txbxContent", namespaces=ns)[0])
        for shape in first_page_slots[6][3]
    ]
    last_paragraph_index = first_page_slots[-1][0]
    body = root.find(f"{W}body")
    if body is None:
        raise RuntimeError("拼音桌签模板缺少 Word 正文")
    paragraphs = body.findall(f"{W}p")
    for paragraph in paragraphs[last_paragraph_index + 1 :]:
        body.remove(paragraph)
    for page_break in body.xpath('.//w:br[@w:type="page"]', namespaces=ns):
        parent = page_break.getparent()
        if parent is not None:
            parent.remove(page_break)
    retained_slots = collect_dining_slots(root)
    placeholder_people = ["Foreign", "Guest", "张三", "李佳琦"]
    placeholder_pinyin = romanize_names(placeholder_people)
    for slot_index, slot in enumerate(retained_slots):
        person_index = slot_index // PINYIN_SLOTS_PER_PERSON
        name = placeholder_people[person_index]
        for shape_index, shape in enumerate(slot[3]):
            if re.search(r"[\u3400-\u9fff]", name):
                box = replace_textbox_content(shape, bilingual_prototypes[shape_index])
                set_bilingual_box_text(
                    box,
                    name,
                    placeholder_pinyin[person_index],
                    pinyin_line_half_points(placeholder_pinyin[person_index]),
                )
            else:
                box = replace_textbox_content(shape, latin_prototypes[shape_index])
                set_single_line_box_text(
                    box,
                    placeholder_pinyin[person_index],
                    latin_line_half_points(placeholder_pinyin[person_index]),
                )
    renumber_cloned_shapes(root)
    members["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    write_zip(members, output)


def normalize_latin_name(value: str) -> str:
    words = re.findall(r"[A-Za-z]+(?:['’-][A-Za-z]+)*", value)
    if not words:
        raise ValueError(f"无法识别拼音姓名：{value}")
    return " ".join(word[:1].upper() + word[1:].lower() for word in words)


def romanize_names(names: list[str]) -> list[str]:
    """Convert Chinese names to surname-first, toneless title-case pinyin."""
    chinese_indexes = [index for index, name in enumerate(names) if re.search(r"[\u3400-\u9fff]", name)]
    raw_pinyin: dict[int, list[str]] = {}
    if chinese_indexes:
        swift_script = Path(__file__).with_name("han_to_pinyin.swift")
        if not swift_script.exists():
            raise RuntimeError("缺少中文转拼音脚本 han_to_pinyin.swift")
        payload = "\n".join(names[index] for index in chinese_indexes) + "\n"
        result = subprocess.run(
            ["/usr/bin/swift", str(swift_script)],
            input=payload,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "中文转拼音失败")
        lines = result.stdout.splitlines()
        if len(lines) != len(chinese_indexes):
            raise RuntimeError("中文转拼音返回数量与姓名数量不一致")
        for index, line in zip(chinese_indexes, lines, strict=True):
            tokens = re.findall(r"[A-Za-z]+", line)
            if not tokens:
                raise ValueError(f"无法转换姓名“{names[index]}”，请直接输入目标拼音")
            raw_pinyin[index] = tokens

    output: list[str] = []
    for index, name in enumerate(names):
        if index not in raw_pinyin:
            output.append(normalize_latin_name(name))
            continue
        compact = re.sub(r"\s+", "", name)
        tokens = raw_pinyin[index]
        surname_length = 2 if compact[:2] in COMPOUND_SURNAMES else 1
        if len(tokens) <= surname_length:
            output.append(normalize_latin_name(" ".join(tokens)))
            continue
        if surname_length == 1 and compact[:1] in SURNAME_PINYIN_OVERRIDES:
            surname = SURNAME_PINYIN_OVERRIDES[compact[:1]]
        else:
            surname = "".join(token.lower() for token in tokens[:surname_length]).capitalize()
        given_name = "".join(token.lower() for token in tokens[surname_length:]).capitalize()
        separator = "  " if len(compact) == 2 and surname_length == 1 else " "
        output.append(f"{surname}{separator}{given_name}")
    return output


def pinyin_em_width(value: str) -> float:
    """Approximate bold serif Latin width in em units for deterministic fitting."""
    narrow = set("Iijlrtf1'’")
    wide = set("MWmwQG@")
    total = 0.0
    for character in value:
        if character in {" ", "\u00a0"}:
            total += 0.28
        elif character in narrow:
            total += 0.30
        elif character in wide:
            total += 0.82
        elif character.isupper():
            total += 0.64
        else:
            total += 0.52
    return max(total, 1.0)


def fitted_half_points(value: str, maximum_points: int) -> str:
    fitted_points = int((PINYIN_INNER_WIDTH_POINTS * 0.94) // pinyin_em_width(value))
    if fitted_points < 30:
        raise ValueError(f"姓名“{value}”过长，30磅仍可能超出原模板框")
    return str(min(maximum_points, fitted_points) * 2)


def latin_line_half_points(value: str) -> str:
    return fitted_half_points(value, int(PINYIN_LATIN_HALF_POINTS) // 2)


def pinyin_line_half_points(value: str) -> str:
    return fitted_half_points(value, int(PINYIN_ROMAN_HALF_POINTS) // 2)


def page_break_paragraph() -> etree._Element:
    paragraph = etree.Element(f"{W}p")
    run = etree.SubElement(paragraph, f"{W}r")
    page_break = etree.SubElement(run, f"{W}br")
    page_break.set(f"{W}type", "page")
    return paragraph


def replace_pinyin_docx(
    template: Path,
    names: list[str],
    output: Path,
    explicit_soffice: str | None,
) -> None:
    converted = convert_legacy(template, ".docx", explicit_soffice)
    with zipfile.ZipFile(converted) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    root = etree.fromstring(members["word/document.xml"])
    body = root.find(f"{W}body")
    if body is None:
        raise RuntimeError("拼音桌签模板缺少 Word 正文")
    template_slots = collect_dining_slots(root)
    expected_template_slots = PINYIN_PEOPLE_PER_PAGE * PINYIN_SLOTS_PER_PERSON
    if len(template_slots) != expected_template_slots:
        raise RuntimeError(
            f"拼音桌签模板必须只保留第一页的 {expected_template_slots} 个框，实际为 {len(template_slots)} 个"
        )
    ns = {"w": W_NS}
    latin_prototypes = [
        deepcopy(shape.xpath(".//w:txbxContent", namespaces=ns)[0])
        for shape in template_slots[0][3]
    ]
    bilingual_prototypes = [
        deepcopy(shape.xpath(".//w:txbxContent", namespaces=ns)[0])
        for shape in template_slots[4][3]
    ]
    section = body.find(f"{W}sectPr")
    template_children = [deepcopy(child) for child in body if child is not section]
    if not template_children:
        raise RuntimeError("拼音桌签模板第一页没有可复制内容")
    for child in list(body):
        body.remove(child)
    page_count = (len(names) + PINYIN_PEOPLE_PER_PAGE - 1) // PINYIN_PEOPLE_PER_PAGE
    for page_index in range(page_count):
        if page_index:
            body.append(page_break_paragraph())
        for child in template_children:
            body.append(deepcopy(child))
    if section is not None:
        body.append(section)
    renumber_cloned_shapes(root)
    pinyin_names = romanize_names(names)
    slots = collect_dining_slots(root)
    expected_slot_count = page_count * expected_template_slots
    if len(slots) != expected_slot_count:
        raise RuntimeError(
            f"拼音桌签扩页失败：应有 {expected_slot_count} 个框，实际为 {len(slots)} 个"
        )
    for slot_index, slot in enumerate(slots):
        person_index = slot_index // PINYIN_SLOTS_PER_PERSON
        for shape_index, shape in enumerate(slot[3]):
            if person_index >= len(names):
                box = replace_textbox_content(shape, bilingual_prototypes[shape_index])
                set_bilingual_box_text(box, "\u00a0", "\u00a0", PINYIN_ROMAN_HALF_POINTS)
                continue
            name = names[person_index]
            pinyin_name = pinyin_names[person_index]
            if re.search(r"[\u3400-\u9fff]", name):
                box = replace_textbox_content(shape, bilingual_prototypes[shape_index])
                set_bilingual_box_text(
                    box,
                    name,
                    pinyin_name,
                    pinyin_line_half_points(pinyin_name),
                )
            else:
                box = replace_textbox_content(shape, latin_prototypes[shape_index])
                set_single_line_box_text(box, pinyin_name, latin_line_half_points(pinyin_name))
    members["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    write_zip(members, output)


def replace_docx(template: Path, names: list[str], output: Path, explicit_soffice: str | None) -> None:
    converted = convert_legacy(template, ".docx", explicit_soffice)
    with zipfile.ZipFile(converted) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    root = etree.fromstring(members["word/document.xml"])
    ns = {"w": W_NS, "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
    slots = collect_dining_slots(root)
    capacity = len(slots)
    person_capacity = capacity // 2
    if len(names) > person_capacity:
        raise ValueError(f"餐桌签模板最多 {person_capacity} 个人（每人两个框），本次有 {len(names)} 人")
    used_slots = len(names) * 2
    for index, (_, _, _, shapes) in enumerate(slots):
        # Keep an invisible glyph in unused boxes so legacy anchored shapes do not
        # collapse and change pagination after the old name is removed.
        value = format_aligned_name(names[index // 2]) if index < used_slots else "\u00a0"
        for shape in shapes:
            for box in shape.xpath(".//w:txbxContent", namespaces=ns):
                set_textbox_text(shape, box, value)
    # Keep the original card geometry on the last used page, but remove the
    # many unused template pages that follow it. This makes a short list
    # genuinely print-ready instead of returning dozens of blank pages.
    last_paragraph_index = slots[used_slots - 1][0]
    body = root.find(f"{W}body")
    if body is not None:
        body_paragraphs = body.findall(f"{W}p")
        for paragraph in body_paragraphs[last_paragraph_index + 1 :]:
            body.remove(paragraph)
        kept_paragraphs = body.findall(f"{W}p")
        if kept_paragraphs:
            for page_break in kept_paragraphs[-1].xpath('.//w:br[@w:type="page"]', namespaces=ns):
                parent = page_break.getparent()
                if parent is not None:
                    parent.remove(page_break)
    members["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_zip(members, output)


def set_inline_string(cell: etree._Element, value: str) -> None:
    for child in list(cell):
        cell.remove(child)
    cell.set("t", "inlineStr")
    inline = etree.SubElement(cell, f"{MAIN}is")
    text = etree.SubElement(inline, f"{MAIN}t")
    if value.startswith(" ") or value.endswith(" "):
        text.set(f"{{{XML_NS}}}space", "preserve")
    text.text = value


def restore_meeting_template_font(members: dict[str, bytes], cells: list[etree._Element]) -> None:
    """Undo LibreOffice's legacy-XLS font mapping for the template's name cells."""
    styles_data = members.get("xl/styles.xml")
    if styles_data is None:
        return
    styles = etree.fromstring(styles_data)
    fonts = styles.find(f"{MAIN}fonts")
    cell_xfs = styles.find(f"{MAIN}cellXfs")
    if fonts is None or cell_xfs is None:
        return
    font_ids: set[int] = set()
    for cell in cells:
        style_id = cell.get("s")
        if style_id is None or not style_id.isdigit() or int(style_id) >= len(cell_xfs):
            continue
        font_id = cell_xfs[int(style_id)].get("fontId")
        if font_id is not None and font_id.isdigit() and int(font_id) < len(fonts):
            font_ids.add(int(font_id))
    for font_id in font_ids:
        font = fonts[font_id]
        name = font.find(f"{MAIN}name")
        if name is None:
            name = etree.SubElement(font, f"{MAIN}name")
        name.set("val", MEETING_FONT)
        size = font.find(f"{MAIN}sz")
        if size is None:
            size = etree.SubElement(font, f"{MAIN}sz")
        size.set("val", MEETING_FONT_POINTS)
    members["xl/styles.xml"] = etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def set_meeting_print_layout(root: etree._Element, used_count: int) -> None:
    """Keep only used rows and force exactly two name frames on every printed page."""
    sheet_data = root.find(f"{MAIN}sheetData")
    if sheet_data is None:
        raise RuntimeError("会议桌签模板缺少工作表数据")
    for row in list(sheet_data):
        row_number = row.get("r", "")
        if row_number.isdigit() and int(row_number) > used_count:
            sheet_data.remove(row)
    dimension = root.find(f"{MAIN}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:A{used_count}")
    existing_breaks = root.find(f"{MAIN}rowBreaks")
    if existing_breaks is not None:
        root.remove(existing_breaks)
    break_rows = list(range(2, used_count, 2))
    if not break_rows:
        return
    row_breaks = etree.Element(f"{MAIN}rowBreaks")
    row_breaks.set("count", str(len(break_rows)))
    row_breaks.set("manualBreakCount", str(len(break_rows)))
    for row_number in break_rows:
        page_break = etree.SubElement(row_breaks, f"{MAIN}brk")
        page_break.set("id", str(row_number))
        page_break.set("min", "0")
        page_break.set("max", "16383")
        page_break.set("man", "1")
    predecessor = root.find(f"{MAIN}headerFooter")
    if predecessor is None:
        predecessor = root.find(f"{MAIN}pageSetup")
    if predecessor is None:
        root.append(row_breaks)
    else:
        root.insert(root.index(predecessor) + 1, row_breaks)


def replace_xlsx(template: Path, names: list[str], output: Path, explicit_soffice: str | None) -> None:
    converted = convert_legacy(template, ".xlsx", explicit_soffice)
    with zipfile.ZipFile(converted) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    worksheet_name = next((name for name in members if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)), None)
    if worksheet_name is None:
        raise RuntimeError("会议桌签模板没有可编辑工作表")
    root = etree.fromstring(members[worksheet_name])
    cells = []
    for cell in root.xpath(".//main:sheetData/main:row/main:c", namespaces={"main": MAIN_NS}):
        ref = cell.get("r", "")
        if re.fullmatch(r"A\d+", ref):
            cells.append(cell)
    if len(names) > len(cells):
        raise ValueError(f"会议桌签模板最多 {len(cells)} 个姓名，本次有 {len(names)} 个")
    for index, cell in enumerate(cells):
        set_inline_string(cell, format_aligned_name(names[index]) if index < len(names) else "")
    restore_meeting_template_font(members, cells)
    set_meeting_print_layout(root, len(names))
    members[worksheet_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_zip(members, output)


def write_zip(members: dict[str, bytes], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def verify_dining_output(output: Path, names: list[str]) -> None:
    """Verify every retained visible/fallback box, font, spacing, and name order."""
    with zipfile.ZipFile(output) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": W_NS, "wps": WPS_NS}
    slots = collect_dining_slots(root)
    expected = [format_aligned_name(name) for name in names for _ in range(2)]
    if len(slots) < len(expected):
        raise RuntimeError(f"餐桌签验收失败：应有 {len(expected)} 个姓名框，实际仅 {len(slots)} 个")
    for slot_index, (_, _, _, shapes) in enumerate(slots):
        expected_text = expected[slot_index] if slot_index < len(expected) else "\u00a0"
        boxes = [box for shape in shapes for box in shape.xpath(".//w:txbxContent", namespaces=ns)]
        if not boxes:
            raise RuntimeError(f"餐桌签验收失败：第 {slot_index + 1} 个框没有文本框内容")
        for box in boxes:
            actual_text = "".join(box.xpath(".//w:t/text()", namespaces=ns))
            if actual_text != expected_text:
                raise RuntimeError(
                    f"餐桌签验收失败：第 {slot_index + 1} 个框应为“{expected_text}”，"
                    f"实际为“{actual_text}”"
                )
            for paragraph in box.xpath("./w:p", namespaces=ns):
                if paragraph.xpath("string(w:pPr/w:jc/@w:val)", namespaces=ns) != "center":
                    raise RuntimeError(f"餐桌签验收失败：第 {slot_index + 1} 个框未水平居中")
                before = paragraph.xpath("string(w:pPr/w:spacing/@w:before)", namespaces=ns)
                after = paragraph.xpath("string(w:pPr/w:spacing/@w:after)", namespaces=ns)
                if before != "600" or after != "0":
                    raise RuntimeError(f"餐桌签验收失败：第 {slot_index + 1} 个框垂直位置不符合模板规则")
            for run in box.xpath(".//w:r", namespaces=ns):
                for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                    font = run.xpath(f"string(w:rPr/w:rFonts/@w:{attribute})", namespaces=ns)
                    if font != DINING_FONT:
                        raise RuntimeError(
                            f"餐桌签验收失败：第 {slot_index + 1} 个框字体应为 {DINING_FONT}，实际为 {font or '空'}"
                        )
                for tag in ("sz", "szCs"):
                    size = run.xpath(f"string(w:rPr/w:{tag}/@w:val)", namespaces=ns)
                    if size != DINING_FONT_HALF_POINTS:
                        raise RuntimeError(
                            f"餐桌签验收失败：第 {slot_index + 1} 个框字号应为 65，XML 实际值为 {size or '空'}"
                        )
        for shape in shapes:
            for body_pr in shape.xpath(".//wps:bodyPr", namespaces=ns):
                if body_pr.get("anchor") != "ctr":
                    raise RuntimeError(f"餐桌签验收失败：第 {slot_index + 1} 个框未垂直居中")


def meeting_cell_text(cell: etree._Element) -> str:
    return "".join(cell.xpath(".//main:t/text()", namespaces={"main": MAIN_NS}))


def verify_meeting_output(output: Path, names: list[str]) -> None:
    """Verify exact meeting-name order plus the template font/alignment style."""
    with zipfile.ZipFile(output) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    worksheet_name = next((name for name in members if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)), None)
    if worksheet_name is None or "xl/styles.xml" not in members:
        raise RuntimeError("会议桌签验收失败：工作表或样式表缺失")
    root = etree.fromstring(members[worksheet_name])
    cells = [
        cell
        for cell in root.xpath(".//main:sheetData/main:row/main:c", namespaces={"main": MAIN_NS})
        if re.fullmatch(r"A\d+", cell.get("r", ""))
    ]
    cells.sort(key=lambda cell: int(cell.get("r", "A0")[1:]))
    expected = [format_aligned_name(name) for name in names]
    actual = [meeting_cell_text(cell) for cell in cells]
    if actual != expected:
        raise RuntimeError("会议桌签验收失败：姓名、顺序或空白框与输入名单不一致")
    expected_breaks = list(range(2, len(expected), 2))
    actual_breaks = [
        int(value)
        for value in root.xpath("./main:rowBreaks/main:brk/@id", namespaces={"main": MAIN_NS})
        if value.isdigit()
    ]
    if actual_breaks != expected_breaks:
        raise RuntimeError("会议桌签验收失败：未按每张纸两个框设置分页")
    page_setup = root.find(f"{MAIN}pageSetup")
    if page_setup is None or page_setup.get("paperSize") != "9" or page_setup.get("orientation") != "portrait":
        raise RuntimeError("会议桌签验收失败：打印纸张必须保持 A4 纵向")
    styles = etree.fromstring(members["xl/styles.xml"])
    fonts = styles.find(f"{MAIN}fonts")
    cell_xfs = styles.find(f"{MAIN}cellXfs")
    if fonts is None or cell_xfs is None:
        raise RuntimeError("会议桌签验收失败：字体或单元格样式缺失")
    for index, cell in enumerate(cells[: len(expected)], start=1):
        style_id = cell.get("s", "")
        if not style_id.isdigit() or int(style_id) >= len(cell_xfs):
            raise RuntimeError(f"会议桌签验收失败：第 {index} 个姓名框样式无效")
        style = cell_xfs[int(style_id)]
        font_id = style.get("fontId", "")
        if not font_id.isdigit() or int(font_id) >= len(fonts):
            raise RuntimeError(f"会议桌签验收失败：第 {index} 个姓名框字体引用无效")
        font = fonts[int(font_id)]
        font_name = font.xpath("string(main:name/@val)", namespaces={"main": MAIN_NS})
        font_size = font.xpath("string(main:sz/@val)", namespaces={"main": MAIN_NS})
        if font_name != MEETING_FONT or font_size != MEETING_FONT_POINTS:
            raise RuntimeError(
                f"会议桌签验收失败：第 {index} 个姓名框应为 {MEETING_FONT} {MEETING_FONT_POINTS} 磅"
            )
        alignment = style.find(f"{MAIN}alignment")
        if alignment is None or alignment.get("horizontal") != "center" or alignment.get("vertical") != "center":
            raise RuntimeError(f"会议桌签验收失败：第 {index} 个姓名框未在框内居中")


def verify_pinyin_output(output: Path, names: list[str]) -> None:
    """Verify WPS page-1 layout: four people, paired frames, Chinese plus pinyin."""
    with zipfile.ZipFile(output) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": W_NS}
    slots = collect_dining_slots(root)
    page_count = (len(names) + PINYIN_PEOPLE_PER_PAGE - 1) // PINYIN_PEOPLE_PER_PAGE
    expected_slots = page_count * PINYIN_PEOPLE_PER_PAGE * PINYIN_SLOTS_PER_PERSON
    if len(slots) != expected_slots:
        raise RuntimeError(f"拼音桌签验收失败：应有 {expected_slots} 个框，实际为 {len(slots)} 个")
    pinyin_names = romanize_names(names)
    for slot_index, (_, _, _, shapes) in enumerate(slots):
        person_index = slot_index // PINYIN_SLOTS_PER_PERSON
        if person_index < len(names):
            source_name = names[person_index]
            expected_pinyin = pinyin_names[person_index]
            chinese_mode = bool(re.search(r"[\u3400-\u9fff]", source_name))
            expected_texts = (
                [format_bilingual_chinese_name(source_name), expected_pinyin]
                if chinese_mode
                else [expected_pinyin]
            )
        else:
            source_name = ""
            expected_pinyin = ""
            chinese_mode = True
            expected_texts = []
        for shape in shapes:
            boxes = shape.xpath(".//w:txbxContent", namespaces=ns)
            if not boxes:
                raise RuntimeError(f"拼音桌签验收失败：第 {slot_index + 1} 个框没有文本内容")
            for box in boxes:
                paragraphs = box.xpath("./w:p", namespaces=ns)
                actual_texts = ["".join(p.xpath(".//w:t/text()", namespaces=ns)) for p in paragraphs]
                actual_texts = [value for value in actual_texts if value]
                if not source_name:
                    if any(value.strip(" \u00a0") for value in actual_texts):
                        raise RuntimeError(f"拼音桌签验收失败：第 {slot_index + 1} 个未用框含有旧姓名")
                    continue
                if actual_texts != expected_texts:
                    raise RuntimeError(
                        f"拼音桌签验收失败：第 {slot_index + 1} 个框应为{expected_texts}，实际为{actual_texts}"
                    )
                runs: list[etree._Element] = []
                for paragraph in paragraphs[: len(expected_texts)]:
                    line_runs = paragraph.xpath("./w:r[w:t]", namespaces=ns)
                    if not line_runs:
                        raise RuntimeError(f"拼音桌签验收失败：第 {slot_index + 1} 个框缺少文字运行")
                    runs.append(line_runs[0])
                expected_sizes = (
                    [
                        PINYIN_CHINESE_HALF_POINTS,
                        PINYIN_ROMAN_HALF_POINTS if not source_name else pinyin_line_half_points(expected_pinyin),
                    ]
                    if chinese_mode
                    else [latin_line_half_points(expected_pinyin)]
                )
                for line_index, (run, expected_size) in enumerate(zip(runs, expected_sizes, strict=True), start=1):
                    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                        font = run.xpath(f"string(w:rPr/w:rFonts/@w:{attribute})", namespaces=ns)
                        if font != PINYIN_FONT:
                            raise RuntimeError(
                                f"拼音桌签验收失败：第 {slot_index + 1} 框第 {line_index} 行字体"
                                f"应为 {PINYIN_FONT}，实际为 {font or '空'}"
                            )
                    size = run.xpath("string(w:rPr/w:sz/@w:val)", namespaces=ns)
                    if size != expected_size:
                        raise RuntimeError(
                            f"拼音桌签验收失败：第 {slot_index + 1} 框第 {line_index} 行字号"
                            f"应为 {int(expected_size) / 2:g}，实际为 {size or '空'}"
                        )
                if chinese_mode:
                    if any(run.xpath("w:rPr/w:b", namespaces=ns) for run in runs):
                        raise RuntimeError(f"拼音桌签验收失败：第 {slot_index + 1} 个双语框不应加粗")
                elif not runs[0].xpath("w:rPr/w:b", namespaces=ns):
                    raise RuntimeError(f"拼音桌签验收失败：第 {slot_index + 1} 个外籍姓名框未保留加粗")
    actual_breaks = root.xpath('count(./w:body/w:p/w:r/w:br[@w:type="page"])', namespaces=ns)
    if int(actual_breaks) != page_count - 1:
        raise RuntimeError("拼音桌签验收失败：分页数量不符合每页四人的规则")
    page_width = root.xpath("string(./w:body/w:sectPr/w:pgSz/@w:w)", namespaces=ns)
    page_height = root.xpath("string(./w:body/w:sectPr/w:pgSz/@w:h)", namespaces=ns)
    if page_width != "11906" or page_height != "16838":
        raise RuntimeError("拼音桌签验收失败：纸张尺寸不是模板的 A4 纵向")


def verify_output(mode: str, output: Path, names: list[str]) -> None:
    if mode == "dining":
        verify_dining_output(output, names)
    elif mode == "meeting":
        verify_meeting_output(output, names)
    elif mode == "pinyin":
        verify_pinyin_output(output, names)
    else:
        raise ValueError(f"不支持的桌签类型：{mode}")


def build_raw(
    mode: str,
    template: Path,
    names: list[str],
    output: Path,
    explicit_soffice: str | None,
) -> None:
    if mode == "dining":
        replace_docx(template, names, output, explicit_soffice)
    elif mode == "meeting":
        replace_xlsx(template, names, output, explicit_soffice)
    elif mode == "pinyin":
        replace_pinyin_docx(template, names, output, explicit_soffice)
    else:
        raise ValueError(f"不支持的桌签类型：{mode}")
    verify_output(mode, output, names)


def replace_one(mode: str, template: Path, names: list[str], output: Path, explicit_soffice: str | None) -> None:
    """Build and verify one branch before atomically replacing the final file."""
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".desk-sign-stage-", dir=output.parent) as stage_dir:
        staged = Path(stage_dir) / output.name
        build_raw(mode, template, names, staged, explicit_soffice)
        os.replace(staged, output)


def output_stem(value: str) -> str:
    """Keep batch output inside the requested directory."""
    stem = Path(value).name.strip()
    stem = re.sub(r"[\\/:*?\"<>|]+", "_", stem)
    return stem or "桌签-替换版"


def template_capacity(mode: str, template: Path, explicit_soffice: str | None) -> int:
    if mode == "dining":
        converted = convert_legacy(template, ".docx", explicit_soffice)
        with zipfile.ZipFile(converted) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
        capacity = len(collect_dining_slots(root)) // 2
    elif mode == "meeting":
        converted = convert_legacy(template, ".xlsx", explicit_soffice)
        with zipfile.ZipFile(converted) as archive:
            worksheet_name = next(
                (info.filename for info in archive.infolist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", info.filename)),
                None,
            )
            if worksheet_name is None:
                raise RuntimeError("会议桌签模板没有可编辑工作表")
            root = etree.fromstring(archive.read(worksheet_name))
        capacity = sum(
            1
            for cell in root.xpath(".//main:sheetData/main:row/main:c", namespaces={"main": MAIN_NS})
            if re.fullmatch(r"A\d+", cell.get("r", ""))
        )
    elif mode == "pinyin":
        converted = convert_legacy(template, ".docx", explicit_soffice)
        with zipfile.ZipFile(converted) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
        capacity = len(collect_dining_slots(root)) // PINYIN_SLOTS_PER_PERSON
    else:
        raise ValueError(f"不支持的桌签类型：{mode}")
    if capacity <= 0:
        raise RuntimeError("模板中没有检测到可替换的桌签框")
    return capacity


def split_name_batches(names: list[str], capacity: int, split_overflow: bool) -> list[list[str]]:
    if len(names) <= capacity:
        return [names]
    if not split_overflow:
        raise ValueError(
            f"模板每个文件最多 {capacity} 人，本次有 {len(names)} 人；"
            "请使用 --split-overflow 自动按原顺序分批。"
        )
    return [names[index : index + capacity] for index in range(0, len(names), capacity)]


def numbered_output(base: Path, index: int, total: int) -> Path:
    if total == 1:
        return base
    return base.with_name(f"{base.stem}-第{index:02d}批{base.suffix}")


def publish_plans(
    plans: list[tuple[str, Path, list[str], Path]],
    output_dir: Path,
    explicit_soffice: str | None,
) -> list[Path]:
    """Build every planned file first; publish none unless every verification passes."""
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".desk-sign-stage-", dir=output_dir) as stage_dir:
        stage_root = Path(stage_dir)
        staged_outputs: list[tuple[Path, Path]] = []
        for plan_index, (mode, template, batch_names, final_output) in enumerate(plans, start=1):
            staged = stage_root / f"{plan_index:03d}-{final_output.name}"
            build_raw(mode, template, batch_names, staged, explicit_soffice)
            staged_outputs.append((staged, final_output))
        for staged, final_output in staged_outputs:
            os.replace(staged, final_output)
    return [plan[3] for plan in plans]


def replace_batch(
    mode: str,
    dining_template: Path | None,
    meeting_template: Path | None,
    pinyin_template: Path | None,
    names: list[str],
    output_dir: Path,
    stem: str,
    explicit_soffice: str | None,
    split_overflow: bool = False,
) -> list[Path]:
    """Plan branches separately, verify all staged files, then publish the group."""
    safe_stem = output_stem(stem)
    plans: list[tuple[str, Path, list[str], Path]] = []
    if mode in {"dining", "all", "all-three"}:
        if dining_template is None:
            raise ValueError("当前模式需要 --dining-template")
        capacity = template_capacity("dining", dining_template, explicit_soffice)
        name_batches = split_name_batches(names, capacity, split_overflow)
        base = output_dir / f"{safe_stem}-吃饭桌签.docx"
        for index, batch_names in enumerate(name_batches, start=1):
            plans.append(("dining", dining_template, batch_names, numbered_output(base, index, len(name_batches))))
    if mode in {"meeting", "all", "all-three"}:
        if meeting_template is None:
            raise ValueError("当前模式需要 --meeting-template")
        capacity = template_capacity("meeting", meeting_template, explicit_soffice)
        name_batches = split_name_batches(names, capacity, split_overflow)
        base = output_dir / f"{safe_stem}-会议桌签.xlsx"
        for index, batch_names in enumerate(name_batches, start=1):
            plans.append(("meeting", meeting_template, batch_names, numbered_output(base, index, len(name_batches))))
    if mode in {"pinyin", "all-three"}:
        if pinyin_template is None:
            raise ValueError("当前模式需要 --pinyin-template")
        base = output_dir / f"{safe_stem}-拼音桌签.docx"
        plans.append(("pinyin", pinyin_template, names, base))
    return publish_plans(plans, output_dir, explicit_soffice)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    extract = sub.add_parser("extract", help="从 Word 或文本文件提取姓名")
    extract.add_argument("--source", required=True)
    extract.add_argument("--output", required=True)
    extract.add_argument("--soffice")
    prepare = sub.add_parser("prepare-pinyin-template", help="只保留原 Word 第一页并建立拼音桌签模板")
    prepare.add_argument("--source", required=True)
    prepare.add_argument("--output", required=True)
    prepare.add_argument("--soffice")
    replace = sub.add_parser("replace", help="替换一个桌签模板中的姓名")
    replace.add_argument("--mode", choices=["dining", "meeting", "pinyin"], required=True)
    replace.add_argument("--template", required=True)
    replace.add_argument("--names-file")
    replace.add_argument("--names")
    replace.add_argument("--output", required=True)
    replace.add_argument("--split-overflow", action="store_true")
    replace.add_argument("--soffice")
    batch = sub.add_parser("batch", help="分别制作吃饭/会议/拼音桌签，并统一返回输出文件")
    batch.add_argument("--mode", choices=["dining", "meeting", "pinyin", "all", "all-three"], required=True)
    batch.add_argument("--dining-template")
    batch.add_argument("--meeting-template")
    batch.add_argument("--pinyin-template")
    batch.add_argument("--names-file")
    batch.add_argument("--names")
    batch.add_argument("--output-dir", required=True)
    batch.add_argument("--output-stem", default="桌签-替换版")
    batch.add_argument("--split-overflow", action="store_true")
    batch.add_argument("--soffice")
    audit = sub.add_parser("audit", help="生成名单核对 CSV，不制作桌签")
    audit.add_argument("--names-file")
    audit.add_argument("--names")
    audit.add_argument("--output", required=True)
    audit.add_argument("--soffice")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "prepare-pinyin-template":
        output = Path(args.output)
        trim_pinyin_template(Path(args.source), output, args.soffice)
        verify_pinyin_output(output, ["Foreign", "Guest", "张三", "李佳琦"])
        print(f"已只保留第一页并建立拼音桌签模板：{output}")
        return 0
    if args.command == "extract":
        names = extract_names(Path(args.source), args.soffice)
        Path(args.output).write_text("\n".join(names) + "\n", encoding="utf-8")
        print(f"提取 {len(names)} 个姓名")
        return 0
    if args.command == "audit":
        names = read_names(args, validate=False)
        output = Path(args.output)
        write_name_audit(names, output)
        issue_count = sum(bool(name_issues(name)) for name in names)
        duplicate_count = sum(count - 1 for count in Counter(names).values() if count > 1)
        print(
            f"已核对 {len(names)} 项：异常 {issue_count} 项，重复出现 {duplicate_count} 次；"
            f"核对表：{output}"
        )
        return 0
    names = read_names(args)
    if args.command == "batch":
        outputs = replace_batch(
            args.mode,
            Path(args.dining_template) if args.dining_template else None,
            Path(args.meeting_template) if args.meeting_template else None,
            Path(args.pinyin_template) if args.pinyin_template else None,
            names,
            Path(args.output_dir),
            args.output_stem,
            args.soffice,
            args.split_overflow,
        )
        print(f"已分别制作 {len(outputs)} 个桌签文件，共使用 {len(names)} 个姓名：")
        for output in outputs:
            print(output)
        return 0
    template = Path(args.template)
    output = Path(args.output)
    if args.mode == "pinyin":
        name_batches = [names]
    else:
        capacity = template_capacity(args.mode, template, args.soffice)
        name_batches = split_name_batches(names, capacity, args.split_overflow)
    plans = [
        (args.mode, template, batch_names, numbered_output(output, index, len(name_batches)))
        for index, batch_names in enumerate(name_batches, start=1)
    ]
    outputs = publish_plans(plans, output.parent, args.soffice)
    print(f"已替换并验收 {len(names)} 个姓名，生成 {len(outputs)} 个文件：")
    for result in outputs:
        print(result)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ValueError, RuntimeError, OSError) as exc:
        raise SystemExit(f"错误：{exc}")
