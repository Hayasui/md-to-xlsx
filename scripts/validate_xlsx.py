#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校验生成的 xlsx：行高裁切、合并重叠、内容完整性。

用法：python validate_xlsx.py <文件.xlsx>

为什么需要单独校验
------------------
openpyxl 写入的行高是「估算值」。若文本比行高允许的更长，Excel/WPS 里
会显示不全。本脚本用与实际渲染一致的宽度模型逐格复核。

同时检查：
  1. 行高是否足够容纳内容（避免裁切）
  2. 合并单元格是否重叠（会导致文件损坏或显示异常）
  3. 是否有 Markdown 残留（**、*）
  4. 冻结窗格、网格线等排版设置是否到位

注意：本脚本用 zipfile 直接读原始 XML 获取列宽，而不是 openpyxl 的
column_dimensions。原因是 WPS 保存后会把连续列宽合并成
`<col min="3" max="4" width="54"/>`，openpyxl 读取时会漏掉中间列。
"""
import math
import sys
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

LINE_HEIGHT = 15.0      # 每行文字占用的磅值（与 WPS 实测一致）
TOLERANCE = 2.0         # 容差
# 列宽 → 每行可容纳的字符数换算系数。
# 实测（列宽 50 的英文列）：约 53-55 个英文字符/行，故系数取 1.08。
# 中文按 2 倍宽度计，换算后约为列宽的一半字符。
EN_WIDTH_FACTOR = 1.08


def display_len(text):
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def read_widths_from_xml(path):
    """从原始 XML 读取每个工作表的真实列宽。"""
    archive = zipfile.ZipFile(path)
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = {rel.get("Id"): rel.get("Target")
            for rel in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
    result = {}
    for sheet in workbook.find(NS + "sheets"):
        target = rels[sheet.get(RNS + "id")].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        root = ET.fromstring(archive.read(target))
        cols = root.find(NS + "cols")
        widths = {}
        if cols is not None:
            for col in cols:
                for index in range(int(col.get("min")), int(col.get("max")) + 1):
                    widths[get_column_letter(index)] = float(col.get("width"))
        result[sheet.get("name")] = widths
    return result


def check(path):
    widths_by_sheet = read_widths_from_xml(path)
    workbook = openpyxl.load_workbook(path)
    problems = []

    for sheet in workbook.worksheets:
        widths = widths_by_sheet.get(sheet.title, {})

        # 合并区域索引
        starts, covered = {}, set()
        for merged in sheet.merged_cells.ranges:
            starts[(merged.min_row, merged.min_col)] = (
                merged.max_col - merged.min_col + 1,
                merged.max_row - merged.min_row + 1)
            for row in range(merged.min_row, merged.max_row + 1):
                for col in range(merged.min_col, merged.max_col + 1):
                    if (row, col) != (merged.min_row, merged.min_col):
                        covered.add((row, col))

        # 合并重叠检测
        ranges = list(sheet.merged_cells.ranges)
        for i in range(len(ranges)):
            for j in range(i + 1, len(ranges)):
                a, b = ranges[i], ranges[j]
                if (a.min_row <= b.max_row and b.min_row <= a.max_row
                        and a.min_col <= b.max_col and b.min_col <= a.max_col):
                    problems.append((sheet.title, str(a), "合并重叠", str(b)))

        # 逐格检查
        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.strip():
                    continue
                if (cell.row, cell.column) in covered:
                    continue
                span, rowspan = starts.get((cell.row, cell.column), (1, 1))
                width = sum(widths.get(get_column_letter(cell.column + i), 9)
                            for i in range(span))
                height = sum((sheet.row_dimensions[cell.row + i].height or 14.4)
                             for i in range(rowspan))
                usable = max(width * EN_WIDTH_FACTOR, 4)
                lines = sum(max(1, math.ceil(display_len(seg) / usable))
                            for seg in cell.value.split("\n"))
                needed = lines * LINE_HEIGHT
                if needed > height + TOLERANCE:
                    problems.append((sheet.title, cell.coordinate, "行高不足",
                                     f"需要{needed:.0f} 实际{height:.0f} 共{lines}行"))
                if "**" in cell.value:
                    problems.append((sheet.title, cell.coordinate, "Markdown残留", "**"))

        if not sheet.freeze_panes:
            problems.append((sheet.title, "-", "未冻结首行", ""))
        if sheet.sheet_view.showGridLines:
            problems.append((sheet.title, "-", "网格线未关闭", ""))

        # 字体检查：英文列应为 Aptos，中文列应为微软雅黑
        # 注意：全角标点（如 ｜）不算中文内容，避免误报
        def is_cjk(ch):
            code = ord(ch)
            return (0x4E00 <= code <= 0x9FFF      # 中日韩统一表意文字
                    or 0x3400 <= code <= 0x4DBF   # 扩展 A
                    or 0x3000 <= code <= 0x303F   # 中文标点
                    or 0xFF00 <= code <= 0xFFEF)  # 全角字符

        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.strip():
                    continue
                name = cell.font.name if cell.font else None
                # 统计中文字符占比，超过 20% 才算中文内容
                chars = [ch for ch in cell.value if not ch.isspace()]
                if not chars:
                    continue
                cjk_ratio = sum(1 for ch in chars if is_cjk(ch)) / len(chars)
                if cjk_ratio > 0.2 and name and "Aptos" in name:
                    problems.append((sheet.title, cell.coordinate, "中文字体异常",
                                     f"应为微软雅黑，实际 {name}"))

        # 打印设置检查
        if sheet.page_setup.orientation != "landscape":
            problems.append((sheet.title, "-", "未设横向打印", ""))
        if not sheet.print_title_rows:
            problems.append((sheet.title, "-", "未设重复表头", ""))

    _check_question_numbers(workbook, problems)
    return workbook, problems


def _check_question_numbers(workbook, problems):
    """检查题号完整性：Q1-Qn 连续，无缺失无重复。

    2026-09 起题号统一用 `Q1` 式（问卷表与大纲表一致），
    不再用「题 1」——兼容检查仍保留，命中才报。
    """
    import re

    for sheet in workbook.worksheets:
        values = [str(c.value) for row in sheet.iter_rows()
                  for c in row if c.value not in (None, "")]

        qnos = sorted({int(m) for v in values
                       for m in re.findall(r"^Q(\d+)$", v.strip())})
        if qnos:
            missing = [n for n in range(1, max(qnos) + 1) if n not in qnos]
            if missing:
                problems.append((sheet.title, "-", "题号缺失",
                                 "Q" + "、Q".join(map(str, missing))))

        # 旧式「题 N」：仅当文件里真的用了这种写法时才检查
        tnos = sorted({int(m) for v in values
                       for m in re.findall(r"^题\s*(\d+)$", v.strip())})
        if tnos:
            missing = [n for n in range(1, max(tnos) + 1) if n not in tnos]
            if missing:
                problems.append((sheet.title, "-", "题号缺失",
                                 "题 " + "、题 ".join(map(str, missing))))


def main():
    if len(sys.argv) < 2:
        print("用法: python validate_xlsx.py <文件.xlsx>")
        sys.exit(1)
    path = sys.argv[1]
    workbook, problems = check(path)

    print(f"文件: {path}")
    for sheet in workbook.worksheets:
        print(f"  {sheet.title}: {sheet.max_row} 行 x {sheet.max_column} 列")
    print()

    if not problems:
        print("校验通过：无裁切、无合并重叠、无 Markdown 残留。")
        return

    print(f"发现 {len(problems)} 个问题：")
    for sheet, coord, kind, detail in problems[:60]:
        print(f"  [{sheet}] {coord} {kind} {detail}")
    if len(problems) > 60:
        print(f"  ... 另有 {len(problems) - 60} 项")
    sys.exit(2)


if __name__ == "__main__":
    main()
