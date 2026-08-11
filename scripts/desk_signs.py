#!/usr/bin/env python3
"""Extract names and replace only name values in the supplied desk-sign templates."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPS = f"{{{WPS_NS}}}"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MAIN = f"{{{MAIN_NS}}}"
XML_NS = "http://www.w3.org/XML/1998/namespace"


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
    header_candidates = {"姓名", "名字", "人员", "人员姓名", "姓名名单"}
    header_index: int | None = None
    name_column: int | None = None
    for row_index, row in enumerate(rows[:3]):
        for column_index, value in enumerate(row):
            normalized = re.sub(r"\s+", "", value)
            if normalized in header_candidates:
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
        for row in rows:
            for value in row:
                values.extend(split_text(value))
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
        values: list[str] = []
        for cell in row.xpath("./main:c", namespaces={"main": MAIN_NS}):
            value = "".join(cell.xpath(".//main:t/text()", namespaces={"main": MAIN_NS}))
            if not value:
                numeric = cell.find(f"{MAIN}v")
                value = numeric.text if numeric is not None and numeric.text else ""
            if cell.get("t") == "s" and value.isdigit() and int(value) < len(shared):
                value = shared[int(value)]
            values.append(value)
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


def read_names(args: argparse.Namespace) -> list[str]:
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


def replace_docx(template: Path, names: list[str], output: Path, explicit_soffice: str | None) -> None:
    converted = convert_legacy(template, ".docx", explicit_soffice)
    with zipfile.ZipFile(converted) as archive:
        members = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    root = etree.fromstring(members["word/document.xml"])
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
    slots = visual_order(slots)
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
    members[worksheet_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_zip(members, output)


def write_zip(members: dict[str, bytes], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    extract = sub.add_parser("extract", help="从 Word 或文本文件提取姓名")
    extract.add_argument("--source", required=True)
    extract.add_argument("--output", required=True)
    extract.add_argument("--soffice")
    replace = sub.add_parser("replace", help="替换模板中的姓名")
    replace.add_argument("--mode", choices=["dining", "meeting"], required=True)
    replace.add_argument("--template", required=True)
    replace.add_argument("--names-file")
    replace.add_argument("--names")
    replace.add_argument("--output", required=True)
    replace.add_argument("--soffice")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "extract":
        names = extract_names(Path(args.source), args.soffice)
        Path(args.output).write_text("\n".join(names) + "\n", encoding="utf-8")
        print(f"提取 {len(names)} 个姓名")
        return 0
    names = read_names(args)
    template = Path(args.template)
    output = Path(args.output)
    if args.mode == "dining":
        replace_docx(template, names, output, args.soffice)
    else:
        replace_xlsx(template, names, output, args.soffice)
    print(f"已替换 {len(names)} 个姓名：{output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ValueError, RuntimeError, OSError) as exc:
        raise SystemExit(f"错误：{exc}")
