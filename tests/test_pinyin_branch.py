from __future__ import annotations

import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree


SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = SKILL_ROOT / "scripts" / "desk_signs.py"
SPEC = importlib.util.spec_from_file_location("desk_signs", SCRIPT)
assert SPEC and SPEC.loader
desk_signs = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(desk_signs)


class PinyinBranchRegressionTest(unittest.TestCase):
    def test_uses_four_people_and_bilingual_source_layout(self) -> None:
        names = ["张三", "张起", "张八", "李佳琦", "王一涵"]
        template = SKILL_ROOT / "assets" / "templates" / "拼音桌签.docx"
        with tempfile.TemporaryDirectory(prefix="pinyin-regression-") as temp_dir:
            output = Path(temp_dir) / "result.docx"
            desk_signs.replace_one("pinyin", template, names, output, None)
            with zipfile.ZipFile(output) as archive:
                root = etree.fromstring(archive.read("word/document.xml"))
        slots = desk_signs.collect_dining_slots(root)
        self.assertEqual(len(slots), 16)  # 2 pages × 4 people × 2 frames
        ns = {"w": desk_signs.W_NS}

        def lines(slot_index: int) -> list[str]:
            box = slots[slot_index][3][0].xpath(".//w:txbxContent", namespaces=ns)[0]
            values = ["".join(p.xpath(".//w:t/text()", namespaces=ns)) for p in box.xpath("./w:p", namespaces=ns)]
            return [value for value in values if value and value.strip(" \u00a0")]

        self.assertEqual(lines(0), ["张  三", "Zhang  San"])
        self.assertEqual(lines(6), ["李佳琦", "Li Jiaqi"])
        self.assertEqual(lines(8), ["王一涵", "Wang Yihan"])
        self.assertEqual(lines(10), [])


if __name__ == "__main__":
    unittest.main()
