#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校验生成的 xlsx：行高裁切、合并重叠、内容完整性。

用法：python validate_xlsx.py <文件.xlsx>

为什么需要单独校验
------------------
openpyxl 写入的行高是「估算值」。若文本比行高允许的更长，Excel/WPS 里
会显示不全。本脚本用与实际渲染一致的宽度模型逐格复核。

问题分两级（2026-09-10 起）：
  error   客观错误，必须修复：行高裁切、合并重叠、Markdown 残留、
          排版设置缺失、字体错列、题号缺失/重复、跳转目标不存在
  warning 疑似问题，由人裁决：删减关键词命中、层级称呼混用、
          末板块跳过项的写法

单独运行时只有 error 影响退出码（exit 2）；经 build 脚本运行时
两类都只打印，不挡生成——保持「报出来、人来修」的工作流。

注意：本脚本用 zipfile 直接读原始 XML 获取列宽，而不是 openpyxl 的
column_dimensions。原因是 WPS 保存后会把连续列宽合并成
`<col min="3" max="4" width="54"/>`，openpyxl 读取时会漏掉中间列。
"""
import math
import re
import sys
import zipfile
from collections import Counter
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

# ---------------------------------------------------------------------------
# 内容层面的检查（2026-09-10 新增）
# ---------------------------------------------------------------------------

# 「问卷逻辑」列里的跳转写法：「跳转到 Q6」「跳过……跳到 Q4」等。
# 「跳过这个模块」这类不带题号的写法不命中，不参与目标检查。
JUMP_RE = re.compile(r"跳[转过到至]*[^Q\n]*?Q(\d+)")

# 删减清单关键词：命中说明可能是没删干净的内部内容（warning，由人裁决）。
# 大小写敏感——"Zoom" 是会议软件，"zoom in" 是缩放操作，不能混。
# 「问卷逻辑」列不扫：那一列照 md 原样写，内部性质是合法的。
INTERNAL_KEYWORDS = (
    "UserTesting", "Unmoderated", "Think-Out-Loud", "Zoom",
    "共享屏幕", "无人主持", "留档版",
    "研究落点", "内部备注", "本研究的核心品类", "由平台定向筛选", "硬性门槛题",
)

# 同一层级称呼的近义词组：同一张表命中两种以上就报（warning）。
# 「部分」只认「第 X 部分」，避免命中「大部分玩家」这类普通用语。
# 表头不参与——outline 的「模块」列是固定结构词，不算内容选择。
# 近义词表按踩坑记录逐次扩充，不做通用词频统计（误报会耗掉检查的信誉）。
SECTION_WORD_PATTERNS = (
    ("模块", re.compile(r"模块")),
    ("板块", re.compile(r"板块")),
    ("第X部分", re.compile(r"第[一二三四五六七八九十百\d]+\s*部分")),
    ("章节", re.compile(r"章节")),
    ("品类节", re.compile(r"品类节")),
)


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


def _header_column(sheet, name):
    """表头（第 1 行）里名为 name 的列号；没有返回 None。"""
    for cell in sheet[1]:
        if isinstance(cell.value, str) and cell.value.strip() == name:
            return cell.column
    return None


def _q_cells(sheet):
    """整格恰好是 Qn 的格子：[(行号, n)]。

    纵向合并后题号只出现在题目首行，所以这里每道题只贡献一个格子。
    """
    result = []
    for row in sheet.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                m = re.fullmatch(r"Q(\d+)", c.value.strip())
                if m:
                    result.append((c.row, int(m.group(1))))
    return result


def _iter_text_cells(sheet, skip_cols=()):
    """第 2 行起、跳过指定列的非空文本格。"""
    for row in sheet.iter_rows(min_row=2):
        for c in row:
            if c.column in skip_cols:
                continue
            if isinstance(c.value, str) and c.value.strip():
                yield c


def _section_rows(sheet):
    """survey 表的分节横条行：整行只有 A 列有值。

    2026-09-11 之前这份表最左列是「模块」，横条行的判定是「A 列有值、题号列为空」；
    现在最左列是「目的」，prose 行的目的格里有「板块引导语」这类标签，
    只看题号列会把引导语行误判成横条，所以改成「其余各列全空」才算横条。
    """
    rows = []
    for row in sheet.iter_rows(min_row=2):
        cells = {c.column: c.value for c in row}
        a = cells.get(1)
        if not (isinstance(a, str) and a.strip()):
            continue
        others = [v for col, v in cells.items()
                  if col != 1 and isinstance(v, str) and v.strip()]
        if others:
            continue
        rows.append(row[0].row)
    return rows


def _check_question_numbers(workbook, problems):
    """题号完整性：区间内连续无缺失、无重复（都是 error）。
    首题不是 Q1 只报 warning——两版大纲本来就允许起点不同。

    2026-09 起题号统一用 `Q1` 式（问卷表与大纲表一致），
    不再用「题 1」——兼容检查仍保留，命中才报。

    重复检查的依据：纵向合并后题号只在题目首行的格子里，
    同一题号出现两次，意味着真的有两道题同号
    （历史事故：中文问卷里两道题都曾是 Q4）。
    """
    for sheet in workbook.worksheets:
        counts, coords = Counter(), {}
        for row in sheet.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    m = re.fullmatch(r"Q(\d+)", c.value.strip())
                    if m:
                        v = "Q" + m.group(1)
                        counts[v] += 1
                        coords.setdefault(v, []).append(c.coordinate)
        if counts:
            qnos = sorted(int(v[1:]) for v in counts)
            # 缺号一律报 error；但「首题不是 Q1」单独降成 warning——
            # 有主持人版大纲从 Q2 起是既定的骨架约定（无人版的 Q1 是下载导航，
            # 有人版由试玩说明承载），报 error 会把这条真实设计判成故障。
            # 其余缺口仍按 error，真正的断号照旧挡得住。
            missing = [n for n in range(min(qnos), qnos[-1] + 1) if n not in qnos]
            if missing:
                problems.append(("error", sheet.title, "-", "题号缺失",
                                 "Q" + "、Q".join(map(str, missing))))
            if min(qnos) != 1:
                problems.append(("warning", sheet.title, "-", "首题不是 Q1",
                                 f"本表从 Q{min(qnos)} 起；若是有意为之（如"
                                 f"有主持人版从 Q2 与无人版对齐）可放行"))
            for v in sorted(counts):
                if counts[v] > 1:
                    problems.append(("error", sheet.title,
                                     "、".join(coords[v]), "题号重复",
                                     f"{v} 出现了 {counts[v]} 次"))

        # 旧式「题 N」：仅当文件里真的用了这种写法时才检查
        tnos = set()
        for row in sheet.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    m = re.fullmatch(r"题\s*(\d+)", c.value.strip())
                    if m:
                        tnos.add(int(m.group(1)))
        if tnos:
            missing = [n for n in range(1, max(tnos) + 1) if n not in tnos]
            if missing:
                problems.append(("error", sheet.title, "-", "题号缺失",
                                 "题 " + "、题 ".join(map(str, missing))))


def _check_logic(sheet, problems):
    """「问卷逻辑」列：跳转目标存在性（error）与末板块写法（warning）。

    只做这两项（2026-09-10 与用户议定）。「同一题的多个跳转造成
    不可达题」的图遍历检查刻意不做——某题只能经特定路径到达，
    往往是设计意图，宁漏勿误。
    """
    logic_col = _header_column(sheet, "问卷逻辑")
    if logic_col is None:
        return
    qno_col = _header_column(sheet, "题号")
    qnos = {n for _, n in _q_cells(sheet)}

    logic_cells = []
    for row in sheet.iter_rows(min_row=2):
        for c in row:
            if c.column == logic_col and isinstance(c.value, str) and c.value.strip():
                logic_cells.append((c.row, c.value))

    # 1) 跳转目标存在性：写了「跳转到 Qn」而本表没有 Qn
    for r, text in logic_cells:
        for target in JUMP_RE.findall(text):
            if int(target) not in qnos:
                coord = f"{get_column_letter(logic_col)}{r}"
                problems.append(("error", sheet.title, coord, "跳转目标不存在",
                                 f"写了跳转到 Q{target}，但本表没有 Q{target}"))

    # 2) 末板块写法：最后一个含题目的 section 区段里，
    #    跳过项应写「本模块结束」，而不是指向别处的跳转
    if qno_col is None:
        return
    sections = _section_rows(sheet)
    qno_rows = [r for r, _ in _q_cells(sheet)]
    if not sections or not qno_rows:
        return
    bounds = sections + [sheet.max_row + 1]
    last = None
    for i in range(len(sections)):
        if any(bounds[i] <= r < bounds[i + 1] for r in qno_rows):
            last = i
    if last is None:
        return
    seg_start, seg_end = bounds[last], bounds[last + 1]
    for r, text in logic_cells:
        if seg_start <= r < seg_end and "跳" in text and "本模块结束" not in text:
            coord = f"{get_column_letter(logic_col)}{r}"
            problems.append(("warning", sheet.title, coord, "末板块写法",
                             "末板块的跳过项建议写「本模块结束」"))


def _check_internal_keywords(sheet, problems):
    """删减清单的关键词扫描（warning）。

    只提示，不自动删——命中平台名、内部标注等词，说明可能有
    该删而未删的内部内容，最终由人判断。
    """
    logic_col = _header_column(sheet, "问卷逻辑")
    for c in _iter_text_cells(sheet, skip_cols=(logic_col,)):
        hits = [kw for kw in INTERNAL_KEYWORDS if kw in c.value]
        if hits:
            problems.append(("warning", sheet.title, c.coordinate, "疑似内部内容",
                             "、".join(hits) + "（按删减清单判断是否该删）"))


def _check_naming(sheet, problems):
    """同一层级称呼混用（warning）：模块 / 板块 / 第X部分 / 章节 / 品类节。

    K3P 踩坑：「模块」「板块」「品类节」「第 X 部分」混用，
    读者要自己判断是不是同一层。
    """
    logic_col = _header_column(sheet, "问卷逻辑")
    found = {}
    for c in _iter_text_cells(sheet, skip_cols=(logic_col,)):
        for label, pattern in SECTION_WORD_PATTERNS:
            if label not in found and pattern.search(c.value):
                found[label] = c.coordinate
    if len(found) >= 2:
        detail = "、".join(f"{label}（{coord}）" for label, coord in found.items())
        problems.append(("warning", sheet.title, "-", "层级称呼混用", detail))


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
                    problems.append(("error", sheet.title, str(a), "合并重叠", str(b)))

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
                    problems.append(("error", sheet.title, cell.coordinate, "行高不足",
                                     f"需要{needed:.0f} 实际{height:.0f} 共{lines}行"))
                if "**" in cell.value:
                    problems.append(("error", sheet.title, cell.coordinate,
                                     "Markdown残留", "**"))

        if not sheet.freeze_panes:
            problems.append(("error", sheet.title, "-", "未冻结首行", ""))
        if sheet.sheet_view.showGridLines:
            problems.append(("error", sheet.title, "-", "网格线未关闭", ""))

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
                    problems.append(("error", sheet.title, cell.coordinate, "中文字体异常",
                                     f"应为微软雅黑，实际 {name}"))

        # 打印设置检查
        if sheet.page_setup.orientation != "landscape":
            problems.append(("error", sheet.title, "-", "未设横向打印", ""))
        if not sheet.print_title_rows:
            problems.append(("error", sheet.title, "-", "未设重复表头", ""))

        # 内容层面的检查（2026-09-10 新增）
        _check_logic(sheet, problems)
        _check_internal_keywords(sheet, problems)
        _check_naming(sheet, problems)

    _check_question_numbers(workbook, problems)
    return workbook, problems


def _print_problems(problems, title):
    print(title)
    for level, sheet_name, coord, kind, detail in problems[:60]:
        print(f"  [{sheet_name}] {coord} {kind} {detail}")
    if len(problems) > 60:
        print(f"  ... 另有 {len(problems) - 60} 项")


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

    errors = [p for p in problems if p[0] == "error"]
    warnings = [p for p in problems if p[0] == "warning"]
    if errors:
        _print_problems(errors, f"错误 {len(errors)} 项（必须修复）：")
    if warnings:
        _print_problems(warnings, f"提醒 {len(warnings)} 项（疑似问题，由人裁决）：")
    # 只有 error 才影响退出码；warning 属于「报出来、人来修」
    sys.exit(2 if errors else 0)


if __name__ == "__main__":
    main()
