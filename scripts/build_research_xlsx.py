#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""用研文档 → Excel 构建器（模板脚本 · 三类表）

表分三类，每类可放任意多个实例（含 0 个）：

    needs      需求梳理      纯中文
    survey     问卷（中英）  中英逐行并列，含「问卷逻辑」列
    survey_cn  问卷（纯中文）纯中文，含「问卷逻辑」列
    outline    访谈大纲      中英逐行并列，可选加「追问方向 / 观察记录点」两列

问卷有两种列规格：中英并列的 `survey` 与纯中文的 `survey_cn`。
中文定稿在先、英文尚未制作的项目，用 `survey_cn` 先出纯中文版；
本地化完成后再换成 `survey`。两者共用同一个构建函数，行格式只差英文那一格。

常见的表清单：

    需求梳理 · 玩家筛选问卷 · 无主持人访谈大纲              （三表）
    需求梳理 · 玩家筛选问卷 · 无主持人访谈大纲 · 有主持人访谈大纲 （四表）
    玩家筛选问卷 A · 玩家筛选问卷 B                        （只出问卷，不出大纲）
    需求梳理 · 题材偏好问卷（纯中文）                       （满意度问卷的模块，暂不出英文）

用法
----
1. 修改下方「内容区」：填 NEEDS / SURVEY / SURVEY_CN / OUTLINE_*，再在 META["sheets"] 里排列表序。
2. **把同目录的 validate_xlsx.py 一并复制过来**（否则自动校验会静默跳过）。
   若确实忘了复制，脚本会自动到技能安装目录去找；两处都没有才会跳过，并明确提示。
3. 运行：python build_research_xlsx.py <输出文件.xlsx>
   - 若省略文件名，默认输出 `META["filename"]`。

样式区（Sheet 类、行高估算）是固化规范，一般不需要改。
详细格式约定见 references/format-spec.md。
完整实战样例见 references/example-4sheets.py。

依赖：openpyxl
"""
import math
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================================
# 一、样式与构建工具（固化规范 · 一般不改）
# ============================================================================

FONT_CN = "微软雅黑"      # 中文列字体
FONT_EN = "Aptos"         # 英文列字体（Office 2024+ 默认；未安装时会回退）
LINE_COLOR = "DCE3E8"
_side = Side(style="thin", color=LINE_COLOR)
_side_strong = Side(style="medium", color="B8C6D0")
BORDER = Border(left=_side, right=_side, top=_side, bottom=_side)

LINE_HEIGHT = 15.2      # 每行文字占用的磅值
ROW_PADDING = 5         # 单元格上下留白
ROW_MIN = 24            # 最小行高
ROW_MAX = 300           # 最大行高（防止极端长文撑爆）
BASE_SIZE = 10          # 正文字号

# 打印设置（留档、导出 PDF 时生效）
PRINT_LAYOUT = {
    "orientation": "landscape",   # 横向
    "fit_to_width": 1,            # 缩放到一页宽
    "repeat_header": "1:1",       # 每页重复表头行（表头在第 1 行）
    "margin": 0.4,                # 页边距（英寸）
    "show_page_number": True,     # 页脚显示页码
}

# 界面设置
VIEW_ZOOM = 100               # 默认缩放
FREEZE_ROWS = 1               # 冻结首行（表头）


def _font(size=BASE_SIZE, bold=False, color="1F2A33", italic=False, english=False):
    return Font(name=FONT_EN if english else FONT_CN,
                size=size, bold=bold, color=color, italic=italic)


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _align(h="left", v="top", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# 单元格样式表：key → (字号, 加粗, 颜色, 斜体, 填充, 水平对齐, 垂直对齐)
# 字体名由所在列决定（中文列用 FONT_CN，英文列用 FONT_EN），不在此处写死。
# `judge` / `stop` 两个 key 服务于问卷的「问卷逻辑」列：
#   judge → 一般逻辑（跳转、互斥、需填文本）
#   stop  → 终止类逻辑，红色加粗
_STYLE_SPEC = {
    "title":   (12, True,  "FFFFFF", False, "334E5C", "left",   "center"),
    "note":    (9,  False, "6B7C88", True,  "F7F9FB", "left",   "center"),
    "header":  (10, True,  "FFFFFF", False, "5C7A8A", "center", "center"),
    "section": (10, True,  "1F3A4D", False, "E3EDF3", "left",   "center"),
    "module":  (10, True,  "1F3A4D", False, "EEF4F8", "left",   "center"),
    "key":     (10, True,  "1F3A4D", False, "F4F7F9", "left",   "center"),
    "label":   (10, False, "334E5C", False, None,     "center", "center"),
    "judge":   (10, False, "8A4A2B", False, None,     "left",   "top"),
    "stop":    (10, True,  "A32D2D", False, None,     "left",   "top"),
    "scale":   (10, False, "334E5C", False, "F7F9FB", "left",   "top"),
    "muted":   (10, False, "6B7C88", False, None,     "left",   "top"),
    "body":    (10, False, "1F2A33", False, None,     "left",   "top"),
}


def make_style(style_key, english=False):
    """按样式 key + 是否英文列，生成 (font, fill, alignment)。"""
    size, bold, color, italic, fill_color, h, v = _STYLE_SPEC[style_key]
    font = _font(size, bold, color, italic, english=english)
    fill = _fill(fill_color) if fill_color else None
    return font, fill, _align(h, v)


def clean(text):
    """去掉 Markdown 粗体/斜体标记。"""
    if text is None:
        return ""
    return str(text).replace("**", "").replace("*", "")


def display_len(text):
    """显示宽度：中日韩字符按 2 计，其余按 1 计。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def estimate_lines(text, width):
    """估算文本在给定列宽下需要几行。"""
    if not text:
        return 1
    usable = max(width - 2.5, 4)
    total = 0
    for segment in str(text).split("\n"):
        total += max(1, math.ceil(display_len(segment) / usable))
    return total


class Sheet:
    """按行追加内容，自动计算行高、处理合并、按列切换中英字体。"""

    def __init__(self, workbook, name, widths, tab_color=None, en_columns=()):
        self.ws = workbook.create_sheet(name)
        self.widths = widths
        self.ncols = len(widths)
        self.en_columns = set(en_columns)   # 英文列的列号（1 起）
        for idx, width in enumerate(widths, 1):
            self.ws.column_dimensions[get_column_letter(idx)].width = width
        self.ws.sheet_view.showGridLines = False
        if tab_color:
            self.ws.sheet_properties.tabColor = tab_color
        self.row = 0

    def add(self, cells, height=None):
        """cells: [(文本, 样式key, 跨列数), ...]"""
        self.row += 1
        column = 1
        needed = 0
        for text, style_key, span in cells:
            text = clean(text)
            cell = self.ws.cell(row=self.row, column=column, value=text)
            # 字体判定：整格都在英文列上才算英文格；
            # 跨列（标题、分节等）默认用中文字体。
            english = (span == 1 and column in self.en_columns)
            font, fill, align = make_style(style_key, english)
            cell.font = font
            cell.alignment = align
            if fill is not None:
                cell.fill = fill
            effective_width = sum(self.widths[column - 1: column - 1 + span])
            if style_key != "title":
                lines = estimate_lines(text, effective_width)
                needed = max(needed, lines * LINE_HEIGHT + ROW_PADDING)
            if span > 1:
                self.ws.merge_cells(start_row=self.row, start_column=column,
                                    end_row=self.row, end_column=column + span - 1)
            column += span
        self.ws.row_dimensions[self.row].height = (
            height if height else max(ROW_MIN, min(needed, ROW_MAX)))
        return self.row

    def merge_down(self, column, first, last):
        """纵向合并某列。"""
        if last > first:
            self.ws.merge_cells(start_row=first, start_column=column,
                                end_row=last, end_column=column)

    def mark_block_top(self, row):
        """给区块首行加粗上边框，视觉上分隔题目。"""
        for column in range(1, self.ncols + 1):
            cell = self.ws.cell(row=row, column=column)
            cell.border = Border(left=_side, right=_side,
                                 top=_side_strong, bottom=_side)

    def finish(self, freeze):
        self.ws.freeze_panes = freeze
        for row in self.ws.iter_rows(min_row=1, max_row=self.ws.max_row,
                                     max_col=self.ncols):
            for cell in row:
                if cell.border.top is None or cell.border.top.style is None:
                    cell.border = BORDER
        self._apply_print_settings()
        self._apply_view_settings()

    def _apply_print_settings(self):
        """打印设置：横向、缩放到一页宽、重复表头、页脚页码。"""
        ws = self.ws
        ws.page_setup.orientation = PRINT_LAYOUT["orientation"]
        ws.page_setup.fitToWidth = PRINT_LAYOUT["fit_to_width"]
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = PRINT_LAYOUT["repeat_header"]
        margin = PRINT_LAYOUT["margin"]
        ws.page_margins.left = ws.page_margins.right = margin
        ws.page_margins.top = ws.page_margins.bottom = margin
        if PRINT_LAYOUT["show_page_number"]:
            ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
            ws.oddFooter.center.size = 8
            ws.oddFooter.center.color = "808080"

    def _apply_view_settings(self):
        """界面设置：缩放、选中首个内容格。"""
        self.ws.sheet_view.zoomScale = VIEW_ZOOM
        self.ws.sheet_view.zoomScaleNormal = VIEW_ZOOM


def blank(count):
    """生成 count 个空单元格。"""
    return [("", "body", 1) for _ in range(count)]

# ============================================================================
# 二、表结构定义（三类表的列与列宽 · 一般不改）
# ============================================================================

# 大纲的列集。默认六列；要放追问与观察记录点就换 COLUMNS_8。
COLUMNS_6 = ("模块", "目的", "题号", "题型", "中文", "English")
COLUMNS_8 = COLUMNS_6 + ("追问方向", "观察记录点")

# 问卷的两种列规格。末列统一叫「问卷逻辑」，记录题目之间的跳转与因果关系：
# 甄别问卷里写的是终止与通过，满意度问卷里写的是跳转、互斥与需填文本。
# 列名统一，语义由每份表自己的内容决定。
COLUMNS_SURVEY = ("题型", "题号", "中文", "English", "问卷逻辑")
COLUMNS_SURVEY_CN = ("题型", "题号", "中文", "问卷逻辑")

# 各类型的列宽。大纲按列数取不同预设（八列时收窄中文列，控制横向总宽）。
WIDTHS = {
    "needs": (17, 9, 112),
    "survey": (11, 8, 54, 54, 27),
    "survey_cn": (11, 8, 80, 36),
    "outline": {
        6: (19, 21, 8, 14, 50, 50),
        8: (19, 21, 8, 12, 44, 44, 36, 36),
    },
}

# 标签页颜色。outline 多实例时按顺序取色，便于区分。
TAB_COLORS = {
    "needs": "334E5C",
    "survey": "3F6E8C",
    "survey_cn": "3F6E8C",
    "outline": ("6B8E7F", "7A6A9B", "8E7F6B"),
}

# 大纲的行类型 → 目标列名。列名必须出现在该表的 columns 里，否则报错。
OUTLINE_ROW_COLUMNS = {
    "fu": "追问方向",
    "obs": "观察记录点",
}

# ============================================================================
# 三、内容区（每次改这里）
# ============================================================================

# ---------------------------------------------------------------- 需求梳理
# 每项为 ("类型", 标签, 内容)：
#   kv      → 标签 + 单段内容
#   kvlist  → 标签 + 多个 (子标签, 内容) 行
#   numlist → 标签 + 多个 (序号, 内容) 行
#   muted   → kv 的灰色版，用于「待定」
NEEDS = [
    ("kv", "项目", "……"),
    ("kv", "调研背景", "……"),
    ("kv", "核心验证对象", "……"),
    ("numlist", "研究目标", [
        ("……", "研究假设：……"),
    ]),
    ("kv", "研究方法", "……"),
    ("kv", "单场时长", "……"),
    ("kvlist", "样本条件", [
        ("游戏", "……"),
        ("性别", "……"),
        ("年龄", "……"),
    ]),
    ("kv", "目标国家", "……"),
    ("muted", "样本量", "待定"),
    ("muted", "素材", "……：待定"),
]

# ---------------------------------------------------------------- 问卷（中英并列）
# 每项为 ("类型", ...)：
#   section → ("section", "分节标题")
#   prose   → ("prose", [(中文, 英文), ...])
#             开场或结束语。每一对被拆成一行；
#             若要把整段文字放在同一个单元格里，写成单个元组、用 \n 在字符串内换行。
#   q       → ("q", 题型, 题号, 中文题干, 英文题干, 问卷逻辑, [(中文选项, 英文选项, 逻辑), ...])
#             题目行的「问卷逻辑」是整题的默认值；选项行的逻辑写在每个选项上。
#             逻辑留空 = 无特殊行为；甄别问卷里写 "终止" 的项，脚本按红色加粗渲染。
#             逻辑列照 md 写，不自行改成"机筛"或别的措辞。
SURVEY = [
    ("section", "开场说明"),
    ("prose", [
        ("……", "……"),
    ]),
    ("section", "第一部分：……"),
    ("q", "单选", "Q1",
     "……", "……",
     "",
     [("……", "……", ""),
      ("……", "……", "终止")]),
    ("section", "结束语"),
    ("prose", [
        ("……\n……", "……\n……"),
    ]),
]

# ---------------------------------------------------------------- 问卷（纯中文）
# 用于中文定稿在先、英文版尚未制作的场景。行格式比 SURVEY 少一处英文：
#   section → ("section", "分节标题")
#   prose   → ("prose", ["段落", ...])          每个元素占一行
#   q       → ("q", 题型, 题号, 题干, [(选项, 逻辑), ...])
#             题型、题号两列在该题所有行上纵向合并
#   field   → ("field", "题干附属的输入行", "问卷逻辑")
#             追加在当前题目之后，与上一题共用题型与题号。
#             用于「用一句话说说为什么」这类并列在选择题下的补充输入；
#             它的存在前提是平台支持"选择题 + 附加文本"，不支持时要另立题号。
#
# 「问卷逻辑」记录玩家选了这一项之后会发生什么，写在触发它的那一行上：
#   跳转     → "跳转到 Q6"
#   终止     → "本模块结束"（红色加粗，用 "终止" 字样以外的一般描述也行）
#   互斥     → "与其余色调选项互斥，不可同时选中"
#   需填文本 → "需填写文本（必填，≤20 字符）"
# 吸引点类题目（依赖前一道题才有意义的那种）不写条件，其出现与否由上一条跳转隐含。
SURVEY_CN = [
    ("section", "开场说明"),
    ("prose", [
        "……",
    ]),
    ("section", "第一部分 · ……"),
    ("q", "单选", "Q1",
     "……",
     [("……", ""),
      ("其他", "需填写文本（必填，≤20 字符）")]),
    ("q", "填空", "Q2",
     "……",
     [("暂时想不起来，跳过这个模块", "跳转到 Q4")]),
    ("q", "多选", "Q3",
     "……（最多选 3 项）",
     [("……", ""),
      ("……", "")]),
    ("field", "用一句话说说为什么（选填）", ""),
]

# ---------------------------------------------------------------- 访谈大纲（无主持人）
# 每项为 ("类型", ...)：
#   page  → ("page", "Task B · Page 5", "本页目的")
#   kv    → ("kv", 行标签, 中文, 英文)
#   sub   → ("sub", 中文引导语, 英文引导语)
#   q     → ("q", 目的, 题号, 题型, 中文题干, 英文题干)
#   scale → ("scale", 中文量表说明, 英文量表说明)
#   opt   → ("opt", 中文选项, 英文选项)
#   题号 / 题型留空字符串表示该行不需要。
OUTLINE_UNMODERATED = [
    ("page", "Task A · Instruction", "开场说明：……"),
    ("kv", "页面标题", "……", "……"),
    ("q", "轻背景：……", "Q1", "单选", "……", "……"),
    ("opt", "……", "……"),
]

# ---------------------------------------------------------------- 访谈大纲（有主持人）
# 同上，另加两种行类型（需在 columns 里声明对应列）：
#   fu  → ("fu", 中文追问，每行一条)    落到「追问方向」列
#   obs → ("obs", 中文记录点)           落到「观察记录点」列
OUTLINE_MODERATED = [
    ("page", "环节 1 · ……", "……"),
    ("kv", "提问话术", "……", "……"),
    ("obs", "……"),
    ("q", "……", "Q1", "口头回答", "……", "……"),
    ("fu", "……"),
    ("obs", "……"),
]

# ---------------------------------------------------------------- 表清单
# sheets 的顺序就是工作表顺序。三类都可放任意多个实例，也可以不放。
# 表名跟随用途：「玩家筛选问卷」（甄别）、「题材偏好问卷」「满意度问卷」（满意度类）……
META = {
    "filename": "【YYMM】项目名.xlsx",
    "sheets": [
        {"type": "needs", "name": "需求梳理", "content": NEEDS},
        {"type": "survey", "name": "玩家筛选问卷", "content": SURVEY},
        {"type": "outline", "name": "无主持人访谈大纲",
         "content": OUTLINE_UNMODERATED},
        {"type": "outline", "name": "有主持人访谈大纲",
         "content": OUTLINE_MODERATED, "columns": COLUMNS_8},
        # 纯中文的问卷（不列英文列）换成这一条：
        # {"type": "survey_cn", "name": "题材偏好问卷", "content": SURVEY_CN},
    ],
}

# ============================================================================
# 四、构建入口
# ============================================================================


def resolve_widths(spec):
    """按类型与列数取列宽，允许 spec["widths"] 覆盖。"""
    kind = spec["type"]
    if kind == "outline":
        ncols = len(spec.get("columns", COLUMNS_6))
        table = WIDTHS["outline"]
        if ncols not in table:
            raise ValueError(
                f'大纲「{spec["name"]}」列数为 {ncols}，没有对应的列宽预设；'
                f'请在本文件的 WIDTHS["outline"] 里补一条。')
        widths = list(table[ncols])
    else:
        widths = list(WIDTHS[kind])
    if "widths" in spec:
        widths = [spec["widths"][i] if i in spec["widths"] else w
                  for i, w in enumerate(widths)]
    return widths


def resolve_tab_color(spec, outline_index):
    """标签页颜色：spec 里可显式指定，否则按类型取；大纲多实例顺延取色。"""
    if "tab_color" in spec:
        return spec["tab_color"]
    kind = spec["type"]
    if kind == "outline":
        palette = TAB_COLORS["outline"]
        return palette[outline_index % len(palette)]
    return TAB_COLORS[kind]


def build_needs(workbook, spec):
    sheet = Sheet(workbook, spec["name"], resolve_widths(spec),
                  resolve_tab_color(spec, 0))
    sheet.add([("项目", "header", 1), ("序号", "header", 1), ("内容", "header", 1)])
    for item in spec["content"]:
        kind = item[0]
        if kind in ("kv", "muted"):
            style = "muted" if kind == "muted" else "body"
            sheet.add([(item[1], "key", 1), ("", "body", 1), (item[2], style, 1)])
        elif kind == "kvlist":
            first = sheet.row + 1
            for sub_label, content in item[2]:
                sheet.add([(item[1] if sheet.row + 1 == first else "", "key", 1),
                           (sub_label, "label", 1), (content, "body", 1)])
            sheet.merge_down(1, first, sheet.row)
        elif kind == "numlist":
            first = sheet.row + 1
            for index, content in enumerate(item[2], 1):
                if isinstance(content, (tuple, list)):
                    content = "\n".join(content)
                sheet.add([(item[1] if index == 1 else "", "key", 1),
                           (str(index), "label", 1), (content, "body", 1)])
            sheet.merge_down(1, first, sheet.row)
        else:
            raise ValueError(f"需求梳理不认识的类型：{kind}")
    sheet.finish("A2")


def build_survey(workbook, spec):
    """问卷表（中英并列）：题型 | 题号 | 中文 | English | 问卷逻辑。"""
    columns = spec.get("columns", COLUMNS_SURVEY)
    if columns != COLUMNS_SURVEY:
        raise ValueError(
            f'问卷「{spec["name"]}」的 columns 为 {columns}；'
            f'中英并列问卷只能用 COLUMNS_SURVEY，纯中文请把 type 写成 survey_cn。')
    sheet = Sheet(workbook, spec["name"], resolve_widths(spec),
                  resolve_tab_color(spec, 0), en_columns=(4,))
    sheet.add([(name, "header", 1) for name in columns])
    for item in spec["content"]:
        kind = item[0]
        if kind == "section":
            sheet.add([(item[1], "section", 5)])
        elif kind == "prose":
            for cn, en in item[1]:
                sheet.add(blank(2) + [(cn, "body", 1), (en, "body", 1), ("", "body", 1)])
        elif kind == "q":
            _, qtype, qno, stem_cn, stem_en, logic, options = item
            first = sheet.row + 1
            sheet.add([(qtype, "label", 1), (qno, "label", 1),
                       (stem_cn, "body", 1), (stem_en, "body", 1), (logic, "judge", 1)])
            for cn, en, mark in options:
                sheet.add(blank(2) + [(cn, "body", 1), (en, "body", 1),
                                      (mark, "stop" if mark else "body", 1)])
            sheet.merge_down(1, first, sheet.row)
            sheet.merge_down(2, first, sheet.row)
            sheet.mark_block_top(first)
        else:
            raise ValueError(f"问卷不认识的类型：{kind}")
    sheet.finish("A2")


def build_survey_cn(workbook, spec):
    """问卷表（纯中文）：题型 | 题号 | 中文 | 问卷逻辑。

    与中英并列版的差别只有一处：没有 English 列，行里也不写英文。
    """
    columns = spec.get("columns", COLUMNS_SURVEY_CN)
    if columns != COLUMNS_SURVEY_CN:
        raise ValueError(
            f'纯中文问卷「{spec["name"]}」的 columns 为 {columns}；'
            f'应为 COLUMNS_SURVEY_CN。')
    sheet = Sheet(workbook, spec["name"], resolve_widths(spec),
                  resolve_tab_color(spec, 0))
    sheet.add([(name, "header", 1) for name in columns])
    block_first = None

    def close_block():
        """把当前题目的「题型 / 题号」两列纵向合并。"""
        nonlocal block_first
        if block_first is not None and sheet.row > block_first:
            sheet.merge_down(1, block_first, sheet.row)
            sheet.merge_down(2, block_first, sheet.row)
        block_first = None

    for item in spec["content"]:
        kind = item[0]
        if kind == "section":
            close_block()
            sheet.add([(item[1], "section", len(columns))])
        elif kind == "prose":
            close_block()
            for paragraph in item[1]:
                sheet.add(blank(2) + [(paragraph, "body", 1), ("", "body", 1)])
        elif kind == "q":
            close_block()
            _, qtype, qno, stem, options = item
            sheet.add([(qtype, "label", 1), (qno, "label", 1),
                       (stem, "body", 1), ("", "body", 1)])
            block_first = sheet.row
            sheet.mark_block_top(sheet.row)
            for option, logic in options:
                sheet.add(blank(2) + [(option, "body", 1),
                                      (logic, "stop" if logic else "body", 1)])
        elif kind == "field":
            if block_first is None:
                raise ValueError(
                    f'纯中文问卷「{spec["name"]}」的 field 行没有可依附的题目；'
                    f'它必须紧跟在某道题之后。')
            _, text, logic = item
            sheet.add(blank(2) + [(text, "scale", 1), (logic, "judge", 1)])
        else:
            raise ValueError(f"纯中文问卷不认识的类型：{kind}")
    close_block()
    sheet.finish("A2")


def build_outline(workbook, spec, outline_index):
    columns = spec.get("columns", COLUMNS_6)
    ncols = len(columns)
    widths = resolve_widths(spec)
    en_columns = (columns.index("English") + 1,) if "English" in columns else ()
    sheet = Sheet(workbook, spec["name"], widths,
                  resolve_tab_color(spec, outline_index), en_columns=en_columns)
    sheet.add([(name, "header", 1) for name in columns])

    # 按列名定位：列没声明却用了对应行类型，就当场报错，不静默丢内容。
    def col_of(row_kind, label):
        name = OUTLINE_ROW_COLUMNS.get(row_kind, label)
        if name not in columns:
            raise ValueError(
                f'大纲「{spec["name"]}」用了 "{row_kind}" 行，但 columns 里没有「{name}」列。')
        return columns.index(name) + 1

    page_first = None
    block_first = None

    def close_block():
        """把当前题目的「目的 / 题号 / 题型」三列纵向合并。"""
        nonlocal block_first
        if block_first is not None and sheet.row > block_first:
            for name in ("目的", "题号", "题型"):
                if name in columns:
                    sheet.merge_down(columns.index(name) + 1, block_first, sheet.row)
        block_first = None

    def pad(cells):
        """补足到 ncols 列。"""
        return cells + [("", "body", 1)] * (ncols - len(cells))

    for item in spec["content"]:
        kind = item[0]
        if kind == "page":
            close_block()
            if page_first is not None:
                sheet.merge_down(1, page_first, sheet.row)
            sheet.add([(item[1], "module", 1), (item[2], "section", ncols - 1)],
                      height=36)
            page_first = sheet.row
        elif kind == "kv":
            close_block()
            _, label, cn, en = item
            sheet.add(pad([("", "module", 1), (label, "key", 1), ("", "label", 1),
                           ("", "label", 1), (cn, "body", 1), (en, "body", 1)]))
        elif kind == "sub":
            close_block()
            _, cn, en = item
            sheet.add(pad([("", "module", 1), ("页面引导语", "key", 1), ("", "label", 1),
                           ("", "label", 1), (cn, "body", 1), (en, "body", 1)]))
        elif kind == "q":
            close_block()
            _, purpose, qno, qtype, cn, en = item
            sheet.add(pad([("", "module", 1), (purpose, "key", 1), (qno, "label", 1),
                           (qtype, "label", 1), (cn, "body", 1), (en, "body", 1)]))
            block_first = sheet.row
            sheet.mark_block_top(sheet.row)
        elif kind in ("scale", "opt"):
            _, cn, en = item
            style = "scale" if kind == "scale" else "body"
            sheet.add(pad([("", "module", 1), ("", "body", 1), ("", "body", 1),
                           ("", "body", 1), (cn, style, 1), (en, style, 1)]))
        elif kind in ("fu", "obs"):
            _, cn = item
            cells = [("", "body", 1)] * ncols
            cells[col_of(kind, "") - 1] = (cn, "body", 1)
            sheet.add(cells)
        else:
            raise ValueError(f"访谈大纲不认识的类型：{kind}")
    close_block()
    if page_first is not None:
        sheet.merge_down(1, page_first, sheet.row)
    sheet.finish("A2")


def _backup_if_exists(path):
    """目标文件已存在时，先备份到 _backup/ 再生成。"""
    import os
    import shutil
    import datetime
    if not os.path.exists(path):
        return
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(path)) or ".", "_backup")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = os.path.splitext(os.path.basename(path))[0]
    target = os.path.join(backup_dir, f"{name}_{stamp}.xlsx")
    shutil.copy2(path, target)
    print(f"  已备份原文件 → {target}")


def load_validator():
    """找到 validate_xlsx 模块。同目录优先，其次技能安装目录。

    校验脚本漏复制时不该静默跳过——「已生成」的提示照常打印，
    很容易被当成一切正常。所以两处都找过才算真的没有。
    """
    try:
        import validate_xlsx
        return validate_xlsx
    except ImportError:
        pass
    fallback = os.path.join(os.path.expanduser("~"), ".workbuddy", "skills",
                            "md-to-research-xlsx", "scripts")
    if os.path.isdir(fallback) and fallback not in sys.path:
        sys.path.insert(0, fallback)
        try:
            import validate_xlsx
            print(f"（校验脚本取自技能目录：{fallback}）")
            return validate_xlsx
        except ImportError:
            pass
    return None


def main():
    # 参数里除开关以外的那一个才是输出文件名。开关写成 --no-backup，
    # 顺序随意；不这样过滤的话，把开关写在前面会被当成文件名，
    # 生成一个名叫 "--no-backup" 的文件，且不报错。
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    output = args[0] if args else META["filename"]
    unknown = [a for a in args[1:]]
    if unknown:
        print(f"（忽略多余的参数：{unknown}）")
    if "--no-backup" not in sys.argv:
        _backup_if_exists(output)

    workbook = Workbook()
    workbook.remove(workbook.active)
    outline_index = 0
    for spec in META["sheets"]:
        kind = spec["type"]
        if kind == "needs":
            build_needs(workbook, spec)
        elif kind == "survey":
            build_survey(workbook, spec)
        elif kind == "survey_cn":
            build_survey_cn(workbook, spec)
        elif kind == "outline":
            build_outline(workbook, spec, outline_index)
            outline_index += 1
        else:
            raise ValueError(
                f'不认识的表类型：{kind}'
                f'（应为 needs / survey / survey_cn / outline）')
    workbook.save(output)

    print("已生成:", output)
    for sheet in workbook.worksheets:
        print(f"  {sheet.title}: {sheet.max_row} 行 x {sheet.max_column} 列")

    # 自动校验
    validator = load_validator()
    if validator is None:
        print("\n未找到 validate_xlsx.py，本次跳过自动校验。")
        print("  请把技能里的 scripts/validate_xlsx.py 复制到工作目录后重跑；")
        print("  「已生成」不等于校验通过，行高裁切与合并重叠都查不出来。")
        return
    _, problems = validator.check(output)
    if problems:
        # 问题分两级：error 必须修复；warning 是疑似问题，由人裁决。
        # 两类都只打印，不挡生成——保持「报出来、人来修」的工作流。
        errors = [p for p in problems if p[0] == "error"]
        warnings = [p for p in problems if p[0] == "warning"]
        if errors:
            print(f"\n校验发现 {len(errors)} 项错误（必须修复）：")
            for _, sheet_name, coord, kind, detail in errors[:60]:
                print(f"  [{sheet_name}] {coord} {kind} {detail}")
            if len(errors) > 60:
                print(f"  ... 另有 {len(errors) - 60} 项")
        if warnings:
            print(f"\n校验提醒 {len(warnings)} 项（疑似问题，由人裁决）：")
            for _, sheet_name, coord, kind, detail in warnings[:60]:
                print(f"  [{sheet_name}] {coord} {kind} {detail}")
            if len(warnings) > 60:
                print(f"  ... 另有 {len(warnings) - 60} 项")
    else:
        print("\n校验通过：无裁切、无合并重叠、无 Markdown 残留。")


if __name__ == "__main__":
    main()
