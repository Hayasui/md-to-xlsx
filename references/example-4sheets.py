#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""用研文档 → Excel 构建器（实战样例 · 三类表 / 四张表）

本文件是「3D二合北美首次测试」的真实脚本，含四张表：
    需求梳理 · 玩家筛选问卷 · 无主持人访谈大纲 · 有主持人访谈大纲
可作为填内容的参考——字段怎么填、长文本怎么换行、大纲的 fu/obs 行怎么用，
都能从这里抄。模板见 scripts/build_research_xlsx.py。

（模板脚本 · 三类表）

表分三类，每类可放任意多个实例（含 0 个）：

    needs   需求梳理        纯中文
    survey  玩家筛选问卷    中英逐行并列
    outline 访谈大纲        中英逐行并列，可选加「追问方向 / 观察记录点」两列

常见的表清单：

    需求梳理 · 玩家筛选问卷 · 无主持人访谈大纲              （三表）
    需求梳理 · 玩家筛选问卷 · 无主持人访谈大纲 · 有主持人访谈大纲 （四表）
    玩家筛选问卷 A · 玩家筛选问卷 B                        （只出问卷，不出大纲）

用法
----
1. 修改下方「内容区」：填 NEEDS / SURVEY / OUTLINE_*，再在 META["sheets"] 里排列表序。
2. **把同目录的 validate_xlsx.py 一并复制过来**（否则自动校验会静默跳过）。
3. 运行：python build_research_xlsx.py <输出文件.xlsx>
   - 若省略文件名，默认输出 `META["filename"]`。

样式区（Sheet 类、行高估算）是固化规范，一般不需要改。
详细格式约定见 references/format-spec.md。
完整实战样例见 references/example-4sheets.py。

依赖：openpyxl
"""
import math
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

# 各类型的列宽。大纲按列数取不同预设（八列时收窄中文列，控制横向总宽）。
WIDTHS = {
    "needs": (17, 9, 112),
    "survey": (11, 8, 54, 54, 27),
    "outline": {
        6: (19, 21, 8, 14, 50, 50),
        8: (19, 21, 8, 12, 44, 44, 36, 36),
    },
}

# 标签页颜色。outline 多实例时按顺序取色，便于区分。
TAB_COLORS = {
    "needs": "334E5C",
    "survey": "3F6E8C",
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

NEEDS = [
    ("kv", "项目", "3D二合北美首次测试（Merge Funland）"),
    ("kv", "调研背景", "了解玩家对二合（Merge）新作的看法和体验，评估二合核心玩家对 3D 棋盘、"
                    "游乐场题材与美术风格的体验感受。"),
    ("kv", "核心验证对象", "3D 二合核心玩法（主要）｜美术风格与题材（次要）"),
    ("numlist", "研究目标", [
        ("3D 棋盘是“加分”还是“劝退”？",
         "研究假设：3D 棋盘与常规二合差异较大，玩家初次接触可能不适应、不喜欢 3D 体验，"
         "导致游玩意愿低。"),
        ("卡通粘土美术是否“太幼稚”？",
         "研究假设：美术偏向卡通粘土风、可能偏低龄化，不符合二合玩家画像（35+），评价偏低。"),
        ("游乐场 + 放置是“差异化亮点”还是“与棋盘脱节”？",
         "研究假设：游乐场题材 + 放置与二合竞品差异大，玩家可能喜好度低、或觉得与棋盘关联弱，"
         "导致游玩意愿低。"),
        ("没剧情是“无所谓”还是“少了点什么”？",
         "研究假设：缺乏剧情点缀的二合游戏，在心流（目标感、连续驱动力、代入感）上可能相对劣势，"
         "导致体验不够完整、持续动力不足。"),
    ]),
    ("kv", "研究方法", "线上 1v1 访谈 或 线上无主持访谈"),
    ("kv", "单场时长", "约 90 分钟（含约 30 分钟自由试玩 + 约 60 分钟试玩后问答）"),
    ("kvlist", "样本条件", [
        ("游戏", "玩过《Gossip Harbor: Merge & Story》且当前等级 ≥50 级；"
               "拥有 6 个月以上连续二合游戏经验，且最近三个月持续玩 GH"),
        ("性别", "女性"),
        ("年龄", "35-60 岁"),
        ("设备", "拥有一台能通过 Google Play 链接下载并游玩的 Android 设备"),
        ("参与意愿", "能参加约 90 分钟的活动"),
        ("其他条件", "无主持人访谈：同意后续可能的 GH 等级抽查；"
                 "有主持人访谈：同意在进入主题前检查当前 GH 等级"),
    ]),
    ("kv", "目标国家", "美国"),
    ("kv", "样本量", "至少 12 人，包含 35-44 岁以及 45 岁以上的玩家"),
    ("muted", "素材", "Google Play 商店链接：待定"),
]

SURVEY = [
    ("section", "开场说明"),
    ("prose", [
        ("感谢你抽空填写这份问卷。我们正在为一款全新的手机游戏招募体验玩家，"
         "想先通过几个简单问题，了解一下你平时玩什么样的游戏。\n"
         "问卷大约需要 3-5 分钟。你的回答仅用于研究筛选，不会用于其他用途。"
         "如果你符合条件，我们会邀请你参加一场约 90 分钟的游戏体验与访谈活动。",
         "This short screening survey takes about 3-5 minutes. We're looking for players "
         "who enjoy casual phone games and play Gossip Harbor: Merge & Story on Android. "
         "If you qualify, you may be invited to a ~90-minute session: you'll try a new "
         "merge game on your Android device, and join a one-on-one interview afterwards. "
         "To take part you'll need an Android device that can download the game through a "
         "Google Play link we share. If this sounds like a fit, please continue."),
    ]),

    ("section", "第一部分：确认你的资格（下面几道题用于判断你是否符合本次研究）"),

    ("q", "多选", "Q1",
     "过去的 3 个月里，你玩过下面哪些休闲益智手游？（可多选）",
     "Which of the following casual puzzle games have you played in the past 3 months? "
     "(Select all that apply)",
     "未选择「二合 / 合成类（Merge）」，且未在「其他」中说明玩过二合游戏 → 终止",
     [("传统三消类 —— 比如 Royal Match、Candy Crush、Royal Kingdom",
       "Match-3 — e.g., Royal Match, Candy Crush, Royal Kingdom", ""),
      ("二合 / 合成类（Merge）—— 比如 Gossip Harbor、Merge Mansion、Travel Town",
       "Merge — e.g., Gossip Harbor, Merge Mansion, Travel Town", ""),
      ("点消 / 消除爆炸类 —— 比如 Toon Blast、Toy Blast",
       "Bubble / blast — e.g., Toon Blast, Toy Blast", ""),
      ("Match 3D / 配对类 —— 比如 Match Factory、Triple Match 3D",
       "Match 3D / pairing — e.g., Match Factory, Triple Match 3D", ""),
      ("方块 / 堆叠类 —— 比如 Block Blast、Block!",
       "Block stacking — e.g., Block Blast, Block!", ""),
      ("解压 / 整理类 —— 比如 Screwdom、Magic Sort",
       "Relaxing / sorting — e.g., Screwdom, Magic Sort", ""),
      ("其他 —— 请简单说明", "Other — please describe briefly", "")]),

    ("q", "单选", "Q2",
     "你玩过《Gossip Harbor: Merge & Story》吗？"
     "（这是一款把合并物品、经营餐厅和故事推进结合在一起的二合游戏）",
     "Have you played Gossip Harbor: Merge & Story? (a merge game that combines combining "
     "items, running a restaurant, and progressing through a story)",
     "只有选①通过；选②或③ → 终止",
     [("玩过（且目前仍在玩）", "Yes, I play it (and I still play it now)", ""),
      ("玩过（但现在已经不玩了）", "Yes, I have played it (but I no longer play it)", "终止"),
      ("没玩过 —— 只是听说过 / 完全不知道",
       "No, I haven't — I've only heard of it / I don't know it", "终止")]),

    ("q", "单选", "Q3",
     "你连续玩《Gossip Harbor》大概多久了？从你开始持续玩"
     "《Gossip Harbor: Merge & Story》算起，到现在大概连续多久了？",
     "How long have you been playing Gossip Harbor continuously? From when you first "
     "started playing Gossip Harbor: Merge & Story continuously until now, roughly how "
     "long has it been?",
     "需要连续玩 GH 3 个月以上；不足 3 个月 → 终止",
     [("不到 1 个月", "Less than 1 month", "终止"),
      ("1-3 个月", "1-3 months", "终止"),
      ("3-6 个月", "3-6 months", ""),
      ("6 个月 - 1 年", "6 months - 1 year", ""),
      ("1 年以上", "More than 1 year", "")]),

    ("q", "单选", "Q4",
     "你在《Gossip Harbor》里现在的等级是多少？可以在进入游戏主界面后，"
     "在左上角的玩家头像处查看，头像旁边的数字就是当前等级。",
     "What is your current level in Gossip Harbor? You can check this by opening the game's "
     "main screen: your level appears as a number next to your player avatar in the top-left "
     "corner.",
     "达到 50 级及以上才符合要求；低于 50 级 → 终止",
     [("49 级及以下", "Level 49 or below", "终止"),
      ("50-99 级", "Level 50-99", ""),
      ("100-149 级", "Level 100-149", ""),
      ("150-199 级", "Level 150-199", ""),
      ("200 级及以上", "Level 200 or above", "")]),

    ("q", "单选", "Q5",
     "你连续玩二合游戏大概多久了？从你第一次接触并持续玩二合游戏"
     "（像 Gossip Harbor、Merge Mansion、Travel Town 这类）算起，到现在大概连续多久了？",
     "How long have you been playing merge games continuously? From the time you first "
     "started playing merge games continuously (like Gossip Harbor, Merge Mansion, "
     "Travel Town, and so on) until now, roughly how long have you been playing?",
     "需要 6 个月以上连续二合游戏经验；不足 6 个月 → 终止",
     [("不到 1 个月", "Less than 1 month", "终止"),
      ("1-6 个月（未满 6 个月）", "1-6 months (not yet 6 months)", "终止"),
      ("6 个月 - 1 年", "6 months - 1 year", ""),
      ("1 - 2 年", "1 - 2 years", ""),
      ("2 年以上", "More than 2 years", "")]),

    ("q", "单选", "Q6",
     "想核实一下你是不是真的有在玩。这道题凭印象答就好，不难。\n"
     "在《Gossip Harbor》里，你要怎么把物品组合出新的东西？请凭你自己的游玩印象来答。",
     "A quick question to confirm you're really playing the game.\n"
     "In Gossip Harbor, how do you combine items to make something new? Please answer "
     "based on your own experience playing the game.",
     "选①通过；其他 → 终止（机筛）",
     [("把两个相同的物品拖到一起，合并成更高一级的物品",
       "Drag two identical items together to merge them into a higher-level item", ""),
      ("把三个相同的物品连成一条线，一起消除掉",
       "Line up three identical items in a row and clear them together", "终止"),
      ("点击屏幕上的泡泡，把它们一个一个戳破",
       "Pop bubbles on the screen one by one", "终止"),
      ("不清楚 / 没注意过", "Not sure / I never noticed", "终止")]),

    ("q", "单选", "Q7",
     "再核实一道题，同样凭印象答就好。\n"
     "在《Gossip Harbor》里，游戏中推动剧情、完成任务需要的资源是？",
     "Another quick question to confirm you're really playing the game.\n"
     "In Gossip Harbor, which resource is needed to progress the story and complete tasks?",
     "选②通过；其他 → 终止（机筛）",
     [("体力", "Energy", "终止"),
      ("金币", "Coins", ""),
      ("宝石", "Gems", "终止"),
      ("XP", "XP", "终止")]),

    ("q", "单选", "Q8",
     "你是否有符合要求的 Android 设备？"
     "这次体验活动需要你通过我们分享的 Google Play 商店链接，"
     "在 Android 设备上下载并游玩一款游戏。"
     "你是否有一台 Android 设备，能够访问我们提供的 Google Play 商店链接，"
     "并下载、运行这款游戏？",
     "Do you have a compatible Android device? For this session you'll need to download "
     "and play a game on an Android device through a Google Play store link we share with "
     "you. Do you have an Android device that can access the Google Play link we provide "
     "and download and run this game?",
     "选①通过；选②，或为 iOS 用户 → 终止",
     [("有，我能通过 Google Play 链接下载并游玩",
       "Yes, I can download and play via a Google Play link", ""),
      ("没有，我无法做到", "No, I can't", "终止")]),

    ("q", "单选", "Q9",
     "你能配合这场游戏体验与访谈吗？\n"
     "如果你符合条件，我们会邀请你参加一场约 90 分钟的活动，分为两部分：\n"
     "· 第一部分（约 30 分钟）：你在自己的 Android 手机上自由体验一款新的二合游戏，"
     "需要一边玩一边说出你的想法。\n"
     "· 第二部分（约 60 分钟）：与我们进行一对一的线上访谈，聊聊你的体验感受。\n"
     "全部完成后，你将获得一份报酬。你的所有回答仅用于研究，严格保密。\n"
     "你是否愿意且能够参加这场全程约 90 分钟的活动？",
     "Can you take part in this game experience and interview?\n"
     "If you qualify, we'd like to invite you to join a session of about 90 minutes, which "
     "has two parts:\n"
     "· Part 1 (about 30 minutes): You'll freely try a new merge game on your own Android "
     "device, and you'll be asked to say your thoughts out loud as you play.\n"
     "· Part 2 (about 60 minutes): A one-on-one online interview to talk about your "
     "experience.\n"
     "After you complete everything, you'll receive a reward. Your responses are used for "
     "research only and are strictly confidential.\n"
     "Are you willing and able to participate in this approximately 90-minute session?",
     "选①通过；选② → 终止",
     [("愿意，且能参加", "Yes, I'm willing and able to participate", ""),
      ("无法参加", "No, I cannot participate", "终止")]),

    ("q", "单选", "Q10",
     "关于游戏经历真实性的确认。为了保证玩家的游戏经历真实可信，"
     "我们可能会在后续随机抽取部分参与者，要求他们展示 Gossip Harbor 的游戏等级"
     "来确认仍符合条件。不符合条件的人将被取消获得奖励的资格。\n"
     "你是否同意参加这项可能的抽查？",
     "Level authenticity confirmation. To make sure players' gaming history is authentic, "
     "our team may randomly select some participants at a later point and ask them to show "
     "their Gossip Harbor level to confirm they still qualify. Anyone who doesn't meet the "
     "criteria will lose their eligibility to receive the reward.\n"
     "Do you agree to this possible check?",
     "选①通过；选② → 终止",
     [("同意", "Yes, I agree", ""),
      ("不同意", "No, I don't agree", "终止")]),

    ("section", "结束语"),
    ("prose", [
        ("再次感谢你的耐心填写！你的回答对我们非常重要。"
         "上面这些问题已经确认了你的基本情况。\n"
         "如果你符合条件，接下来会进入游戏体验环节。试玩过程中请一边玩一边大声说出你的想法"
         "和感受，这会帮助我们了解你对这款游戏的真实体验。\n"
         "试玩结束后，你还需要完成一份简短的问卷，聊聊你玩过的其他二合游戏、"
         "以及你对这款游戏的感受，完成后活动才算全部结束。\n"
         "你的所有回答都将严格保密，仅用于研究目的。\n"
         "如果这次没能通过筛选，也感谢你抽出时间参与。祝你生活愉快！",
         "Thank you again for your time — your responses are very important to us. "
         "The questions above confirm your basic profile.\n"
         "If you qualify, you'll move on to the game experience. During the playtest, please "
         "say your thoughts and feelings out loud as you play — this will help us understand "
         "your real experience with the game.\n"
         "After you finish playing, you'll also need to complete a short questionnaire about "
         "other merge games you play and your thoughts on this game, which wraps up the full "
         "session.\n"
         "All responses will remain confidential and will be used for research purposes only.\n"
         "If you don't qualify this time, thank you for taking part. We hope you have a "
         "great day!"),
    ]),
]

OUTLINE_UNMODERATED = [
    ("page", "Task A · Instruction", "开场说明：引导玩家边玩边说"),
    ("kv", "页面标题", "欢迎参加本次游戏体验", "Welcome"),
    ("kv", "说明文字",
     "你好，欢迎参加这次游戏体验！整个过程分两步：先自己玩大约 30 分钟，"
     "然后回答几个简短的问题，全程约 90 分钟。\n"
     "玩的时候请记住这几点：\n"
     "· 一边玩，一边把脑子里想的说出来——不管是觉得有意思、有点困惑，还是单纯觉得有趣，"
     "都可以随口说出来。没有固定的说法，想怎么说都行。\n"
     "· 你的屏幕和声音会被自动记录下来，所以放开了边玩边说就好，"
     "不用安装任何东西，也不用手动共享屏幕。\n"
     "· 如果遇到游戏问题（比如卡住、闪退、连不上、操作失灵），请告诉我们；"
     "除此之外没人会打扰你，不用管“该怎么玩才对”，按你喜欢的玩法来就行。\n"
     "祝玩得开心！我们等会儿见。",
     "Hi! Welcome to this game session. It has two parts: first you'll play a game for about "
     "30 minutes, then answer a few short questions. The whole thing takes about 90 minutes.\n"
     "A few things to keep in mind while you play:\n"
     "· Say your thoughts out loud — whether a moment feels fun, confusing, or just "
     "interesting, share it as it comes to you. There's no wrong way to put it.\n"
     "· Your screen and audio are recorded automatically, so speak naturally as you play. "
     "You don't need to install anything or share your screen.\n"
     "· If you hit a game issue (like it freezes, crashes, won't connect, or the controls "
     "stop working), let us know; otherwise no one will interrupt you — don't worry about "
     "“the right way to play,” just play the way you like.\n"
     "Have fun! See you in a bit."),

    ("page", "Task A · Navigation", "引导玩家完成下载、打开与试玩"),
    ("kv", "导航任务名称", "下载并打开游戏", "Download & Open the Game"),
    ("kv", "说明文字",
     "现在请完成下面几步：\n"
     "1. 在你的 Android 手机上下载并安装这款游戏。\n"
     "2. 安装完成后，打开游戏，按自己的方式开始玩。游戏里有新手教程，会一步步教你。\n"
     "3. 玩大约 30 分钟，边玩边说你的想法。时间到了之后，你会回到这里做后续回答。",
     "Now please do the following:\n"
     "1. On your device, download and install this game.\n"
     "2. Once it's installed, open the game and start playing however you like. There's a "
     "tutorial that will walk you through the basics.\n"
     "3. Play for about 30 minutes, saying your thoughts out loud as you go. When the time "
     "is up, you'll come back here for a few follow-up questions."),

    ("page", "Task B · Page 1", "轻背景：了解玩家的常用设备"),
    ("q", "了解玩家的常用设备", "Q1", "单选",
     "除了今天用在游戏上的这台手机，你平时还主要在什么设备上玩休闲游戏？",
     "Other than the phone you're using for the game today, what device do you usually "
     "play casual games on?"),
    ("opt", "手机", "Phone"),
    ("opt", "平板", "Tablet"),    ("opt", "手机和平板都差不多", "Phone and tablet about equally"),
    ("opt", "其他", "Other"),

    ("page", "Task B · Page 2", "背景画像：玩过哪些二合、日均时长、付费情况"),
    ("sub", "下面几道题想了解一下你玩过哪些二合游戏。没有对错，按实际情况填就行。"
            "其中有两道题会请你一边选、一边顺口说说——看到“顺便说说”的时候，就放开讲两句。",
     "A few questions to understand which merge games you play. There's no right or wrong "
     "answer — just answer based on your actual experience. Two of these questions ask you "
     "to talk out loud while you choose — when you see “tell us,” just say what's on your "
     "mind."),
    ("q", "确认玩过哪些二合游戏", "Q2", "多选",
     "过去 3 个月里，下面这些二合（合并 / 合成）游戏，你玩过哪些？（可以多选）",
     "Over the past 3 months, which of these merge games have you played? "
     "(Select all that apply)"),
    ("opt", "Travel Town – Merge Adventure", "Travel Town – Merge Adventure"),
    ("opt", "Merge Mansion", "Merge Mansion"),
    ("opt", "Gossip Harbor: Merge & Story", "Gossip Harbor: Merge & Story"),
    ("opt", "Merge Cooking", "Merge Cooking"),
    ("opt", "Flambé: Merge and Cook", "Flambé: Merge and Cook"),
    ("opt", "Tasty Travels: Merge Game", "Tasty Travels: Merge Game"),
    ("opt", "Merge Dragons!", "Merge Dragons!"),
    ("opt", "Love & Pies", "Love & Pies"),
    ("opt", "Merge Miners", "Merge Miners"),
    ("opt", "上面都没有", "None of the above"),
    ("q", "画像分层：日均游戏时长", "Q3", "单选",
     "过去 7 天里，你平均每天大概花多长时间玩二合游戏？"
     "选完之后，顺便说说你一般什么时候玩、一般是什么情况下玩"
     "（比如睡前、通勤、做家务的间隙）。",
     "Over the past 7 days, about how much time did you spend playing merge games on an "
     "average day? After you choose, tell us when you usually play and what you're usually "
     "doing (for example, before bed, on your commute, in between chores)."),
    ("opt", "几乎不玩 / 每天不到 10 分钟", "Almost never / less than 10 minutes per day"),
    ("opt", "每天 10-30 分钟", "10-30 minutes per day"),
    ("opt", "每天 30 分钟 - 1 小时", "30 minutes - 1 hour per day"),
    ("opt", "每天 1-2 小时", "1-2 hours per day"),
    ("opt", "每天 2 小时以上", "More than 2 hours per day"),
    ("q", "画像分层：月度付费金额", "Q4", "单选",
     "过去 1 个月里，你在二合游戏上大概花了多少钱（比如买金币、道具、通行证之类）？"
     "选完之后，顺便说说你最近一次花钱买的是什么。",
     "Over the past month, roughly how much did you spend on merge games (for example, on "
     "coins, items, or a battle pass)? After you choose, tell us what you last bought."),
    ("opt", "没花过钱", "I didn't spend anything"),
    ("opt", "$1-9", "$1-$9"),
    ("opt", "$10-49", "$10-$49"),
    ("opt", "$50-99", "$50-$99"),
    ("opt", "$100 及以上", "$100 or more"),
    ("q", "付费动机与不付费原因", "Q5", "口头回答",
     "最后想请你多聊两句花钱的事。下面两段，按你自己的情况选一段答就行。\n"
     "如果你刚才那道题选的是“没花过钱”——是什么让你没花钱？觉得没必要，"
     "还是一直没碰到想买的？你在别的二合游戏里花过钱吗？\n"
     "如果你花过钱——你一般什么情况下最容易掏钱（比如卡关了、有限时活动、想快点升级）？"
     "在别的二合游戏里也是这样吗？",
     "One more thing about spending. Pick whichever of the two parts below fits you.\n"
     "If you chose “I didn't spend anything” — what's kept you from spending? Is it that you "
     "don't see the need, or that nothing has tempted you yet? Have you spent money in other "
     "merge games?\n"
     "If you have spent money — what tends to make you pull the trigger (for example, being "
     "stuck on a level, a limited-time event, wanting to level up faster)? Is it the same in "
     "other merge games?"),

    ("page", "Task B · Page 3", "定基准：选出与今天这款最像的游戏"),
    ("sub", "接下来想先请你找到一款跟你今天玩的这款最像的二合游戏。"
            "我们会拿它跟今天这款做对比。",
     "First, we'd like you to think of a merge game that's most similar to the one you "
     "played today. We'll use it as a point of comparison."),
    ("q", "定基准：选出最像今天这款的游戏", "Q6", "单选",
     "你刚才玩了今天这款游戏，也回顾了一下之前玩过的那些二合游戏。"
     "你觉得下面哪一款和今天这款最像？可以从玩法、棋盘、题材、画面这些方面想想。"
     "请你记住自己选的这一款——后面几道题还会用到它。",
     "You just played today's game and looked back over the merge games you've played. "
     "Which of these do you think is most like today's game? Think about things like "
     "gameplay, the board, the theme, and the visuals. Please remember the one you pick — "
     "we'll come back to it a few times later."),
    ("opt", "Travel Town – Merge Adventure", "Travel Town – Merge Adventure"),
    ("opt", "Merge Mansion", "Merge Mansion"),
    ("opt", "Gossip Harbor: Merge & Story", "Gossip Harbor: Merge & Story"),
    ("opt", "Merge Cooking", "Merge Cooking"),
    ("opt", "Flambé: Merge and Cook", "Flambé: Merge and Cook"),
    ("opt", "Tasty Travels: Merge Game", "Tasty Travels: Merge Game"),
    ("opt", "Merge Dragons!", "Merge Dragons!"),
    ("opt", "Love & Pies", "Love & Pies"),
    ("opt", "Merge Miners", "Merge Miners"),
    ("opt", "都不像（其他）", "Other"),
    ("q", "基准理由：说明像在哪里", "Q7", "口头回答",
     "为什么觉得这款跟今天的最像？能说说到底是哪方面像吗（比如玩法、棋盘、题材、画面）？"
     "如果你刚才选了“都不像（其他）”，也请告诉我们你心里想的是哪一款。",
     "Why do you think this game is most like today's? Can you tell us what specifically "
     "feels similar (for example, the gameplay, the board, the theme, or the visuals)? "
     "If you chose “Other,” please tell us which game you have in mind."),

    ("page", "Task B · Page 4", "定量核心：满意度、意愿、NPS 与受众"),
    ("sub", "接下来请你给刚才的体验打打分、谈谈感受。答案没有对错，凭第一印象就好。",
     "Now we'd like you to rate your experience and share your thoughts. There's no right "
     "or wrong answer — go with your first impression."),
    ("q", "试玩整体满意度", "Q8", "打分",
     "想想你刚才试玩的感受，给今天这款游戏的整体满意度打个分，1 分最低，10 分最高。"
     "你会打几分？",
     "Based on your experience just now, how satisfied are you overall with today's game, "
     "on a scale of 1 to 10 (1 = the lowest, 10 = the highest)? What would you score it?"),
    ("scale", "1 = 非常不满意 ｜ 10 = 非常满意", "1 = Very unsatisfied ｜ 10 = Extremely satisfied"),
    ("q", "基准游戏满意度（对照）", "Q9", "打分",
     "再想想过去 3 个月里你选的那款游戏给你的整体感受，也给它打个分，1 分最低，10 分最高。"
     "你会打几分？",
     "Now, thinking about your overall experience over the past 3 months with the game you "
     "selected, how satisfied are you overall with it, on a scale of 1 to 10 (1 = the lowest, "
     "10 = the highest)? What would you score it?"),
    ("scale", "1 = 非常不满意 ｜ 10 = 非常满意", "1 = Very unsatisfied ｜ 10 = Extremely satisfied"),
    ("q", "继续体验意愿", "Q10", "打分",
     "假设今天不是测试、也没有任何奖励，就凭你刚才的体验——明天你还想继续玩这款游戏的"
     "可能性有多大？1 分代表肯定不会，10 分代表肯定会。你会打几分？",
     "Forget that this is a test and there's a reward — just going on your experience, how "
     "likely are you to keep playing this game tomorrow? 1 means definitely not, 10 means "
     "definitely yes."),
    ("scale", "1 = 肯定不会玩 ｜ 10 = 肯定会玩", "1 = Definitely won't play ｜ 10 = Definitely will play"),
    ("q", "下载意愿", "Q11", "单选",
     "这款游戏现在已经能在应用商店下载了。假设你还没下载，"
     "你下载它的意愿是下面哪一种？",
     "This game is already available to download on the app store. Imagine you hadn't "
     "downloaded it yet — how likely would you be to download it? Choose one."),
    ("opt", "肯定会下载", "Definitely would download"),
    ("opt", "可能会下载", "Probably would download"),
    ("opt", "不确定", "Not sure"),
    ("opt", "可能不会下载", "Probably wouldn't download"),
    ("opt", "肯定不会下载", "Definitely wouldn't download"),
    ("q", "推荐意愿（NPS）", "Q12", "NPS",
     "今天这款游戏，你愿意推荐给朋友或家人吗？从 0 到 10 分，"
     "0 分代表完全不愿意推荐，10 分代表非常愿意推荐。你会打几分？",
     "For today's game, how likely are you to recommend it to a friend or family member? "
     "From 0 to 10, 0 means not at all likely to recommend, 10 means extremely likely. "
     "What would you score it?"),
    ("q", "推荐意愿对照（NPS）", "Q13", "NPS",
     "同样地，你常玩的那款［你选的基准游戏］，你愿意推荐给朋友或家人吗？"
     "从 0 到 10 分，你会打几分？",
     "Likewise, for the game you usually play (the baseline game you selected), how likely "
     "are you to recommend it to friends or family? From 0 to 10, what would you score it?"),
    ("q", "受众圈层", "Q14", "口头回答",
     "你觉得今天这款游戏主要是给什么样的人玩的？你会推荐给身边什么样的朋友？"
     "是大多数朋友都合适，还是只有一部分爱玩这类游戏的朋友才合适？"
     "那你常玩的那款（你选的基准游戏）呢？",
     "Who do you think today's game is mainly for? Who would you recommend it to — most of "
     "your friends, or mainly the friends who enjoy this kind of game? And what about the "
     "game you usually play (the baseline game you selected)?"),

    ("page", "Task B · Page 5", "了解 3D 棋盘上手难度、棋子辨识度与合成手感"),
    ("sub", "咱们从你刚才玩的棋盘聊起。", "Let's talk about the board you played on."),
    ("q", "3D 棋盘上手感受（定量）", "Q15", "矩阵",
     "回想一下你刚上手这个棋盘时的感受，对下面每句话，你有多同意？"
     "（1 = 很不同意，5 = 很同意）",
     "Thinking back to when you first picked up this board, how much do you agree with each "
     "statement below? (1 = strongly disagree, 5 = strongly agree)"),
    ("scale", "1 = 很不同意 ｜ 5 = 很同意", "1 = Strongly disagree ｜ 5 = Strongly agree"),
    ("opt", "我一上手就能理解这个棋盘怎么玩", "I understood how to play on this board right away"),
    ("opt", "这个立体的棋盘让我更感兴趣", "The 3D board made me more interested"),
    ("opt", "这个立体的棋盘有点乱、不太好看清", "The 3D board felt a bit messy / hard to see clearly"),
    ("opt", "我不知道东西在哪、该从哪下手", "I wasn't sure where things were or where to start"),
    ("q", "棋子合成链路与辨识度", "Q16", "口头回答",
     "这个游戏里，从最普通的棋子一步步合成出更高级棋子，这条链子你能看明白吗？"
     "（比如先合什么、再合什么、最后能出什么）这些棋子（面包、甜甜圈、草莓、海豚这些）"
     "相互之间差别够大吗？有没有长得太像、容易看混的？",
     "Could you follow the chain of merging up from the most basic piece to a higher-level "
     "one? (Like merge this, then that, and see what you end up with.) Do these pieces — "
     "bread, donut, strawberry, dolphin, and so on — look different enough from each other? "
     "Are there any that look so similar that they're easy to confuse?"),
    ("q", "与基准游戏对比及合成手感", "Q17", "口头回答",
     "跟你常玩的那款（你选的基准游戏）比，今天这个棋盘是更顺手还是更别扭？"
     "拖棋子去合成的时候手感怎么样（会不会拖不动、放不准、不好对齐）？"
     "合成成功时的反馈（特效 / 声音 / 动作）让你觉得爽快还是平淡？",
     "Compared to the game you usually play (the baseline game you selected), do you find "
     "this board easier or more awkward to use? When you drag a piece to merge it, how does "
     "it feel — does it drag awkwardly, not drop where you want, or feel hard to line up? "
     "When a merge succeeds, does the feedback (the effects, sound, or motion) feel "
     "satisfying or flat?"),

    ("page", "Task B · Page 6", "了解美术风格感知、适龄性与基准对比"),
    ("sub", "接着聊聊这个游戏的画面和风格。",
     "Next, let's talk about the game's visuals and style."),
    ("q", "美术风格感知与适龄性", "Q18", "口头回答",
     "刚才游戏里那些棋子（合成用的东西）和角色，你觉得它长什么样、什么风格？"
     "如果让你用 1-2 个词形容，你会用哪两个词？这个风格，"
     "你觉得适合你这个年纪的人玩吗，还是有点偏小、偏幼稚？"
     "如果觉得偏幼稚，具体是哪里让你有这种感觉（角色、棋子、颜色，还是整体）？",
     "You saw the pieces (merge items) and characters in the game. How would you describe "
     "their look and style? If you had to pick 1-2 words, which would you choose? Does this "
     "style feel like it suits someone your age, or does it feel a bit young / childish? "
     "If it feels young, what specifically gives you that feeling — the characters, the "
     "pieces, the colors, or the overall look?"),
    ("q", "与基准游戏美术对比", "Q19", "口头回答",
     "跟你常玩的那款（你选的基准游戏）比，你觉得哪个画面更高级、更精致？"
     "你更喜欢哪一个的风格？为什么？",
     "Compared to the game you usually play (the baseline game you selected), which one's "
     "look do you think is more polished / more premium? Which style do you prefer? Why?"),

    ("page", "Task B · Page 7", "了解题材、建地块与放置玩法的感知与理解"),
    ("sub", "最后聊聊这个游戏的题材和玩法。这个游戏是游乐场主题的，你刚才应该也注意到了。",
     "Finally, let's talk about the game's theme and gameplay. This one has a carnival / "
     "theme-park theme — you may have noticed that."),
    ("q", "游乐场题材感知", "Q20", "口头回答",
     "你第一眼看到“游乐场”这个主题，什么感觉？觉得新鲜、有意思，"
     "还是跟平常玩的差别太大、有点怪？你觉得游乐场这个题材，"
     "适不适合做成一款你平时会玩的二合游戏？为什么？",
     "When you first saw the “carnival” theme, what did you think? Did it feel fresh and "
     "interesting, or very different from what you usually play and a bit odd? Do you think "
     "the carnival theme works well as a merge game you'd normally play? Why?"),
    ("q", "建地块 / 外围玩法理解", "Q21", "口头回答",
     "你刚才有没有注意到，游戏里有个花金币搭游乐场（激活格子）的部分？"
     "你明白这部分是干什么的吗？知不知道激活格子之后会发生什么、怎么升级、怎么赚金币？"
     "有哪里让你觉得不清楚？",
     "Did you notice the part of the game where you spend coins to build up the carnival "
     "(activate tiles)? Did you understand what that part is for? Do you know what happens "
     "after you activate a tile — how it levels up, how you earn coins? Was anything about "
     "it unclear?"),
    ("q", "题材与棋盘关联度", "Q22", "口头回答",
     "你觉得游乐场这个题材，和你刚玩的那个棋盘（合成解谜），联系紧密吗？"
     "还是这两部分有点脱节、各玩各的？如果让你改，你会让它俩怎么配合更好？",
     "Do you feel the carnival theme and the board (merge puzzle) you just played are "
     "closely connected, or do they feel a bit disjointed / like two separate things? If "
     "you could change it, how would you make them work together better?"),
    ("q", "放置玩法感知", "Q23", "口头回答",
     "游戏里好像有挂机 / 放置让你能拿到收入（比如放着自动收钱），你注意到了吗？"
     "你觉得“放着也能赚钱”这个玩法，跟你理解的二合游戏顺不顺？"
     "是更有动力，还是感觉跟二合没关系？",
     "The game seems to have idle / auto-income (like collecting money automatically while "
     "you're away or doing nothing). Did you notice it? Do you think “earning money without "
     "doing anything” fits in with the merge games you play? Does it pull you in more, or "
     "does it feel unrelated to merge?"),

    ("page", "Task B · Page 8", "了解剧情缺失的感知及对投入感的影响"),
    ("sub", "最后一道题稍微有点绕——这款游戏到底有没有剧情故事？我们聊聊这个。",
     "One last, and this one's a bit trickier: does this game have a story or not? Let's dig "
     "into that."),
    ("q", "对剧情缺失的最初感知", "Q24", "口头回答",
     "你玩的时候，有没有意识到“这游戏里没什么故事 / 剧情”？还是压根没在意这回事？"
     "如果没在意，是不是因为核心玩法（合成棋盘）本身就够吸引你，"
     "有没有剧情都不太在乎？",
     "While you were playing, did you notice that this game doesn't really have a story or "
     "much of a narrative? Or did you not really think about it? If you didn't think about "
     "it — is that because the core gameplay (the merge board) was engaging enough on its "
     "own, so a story just didn't matter to you?"),
    ("q", "剧情对投入感与心流的影响", "Q25", "口头回答",
     "你玩过的那些二合游戏（比如你选的基准游戏），有没有剧情？如果有，"
     "那些剧情会不会让你更投入、更想一直玩下去？反过来，那些没有剧情的二合游戏，"
     "你玩的时候会不会觉得“少了点什么”“玩着玩着就没目标了”，"
     "还是觉得无所谓、纯玩机制就挺好？",
     "The merge games you've played before (like the baseline game you selected) — do they "
     "have a story? If they do, does the story make you more invested and give you more "
     "reason to keep playing? On the other hand, for merge games without a story, do you "
     "ever feel “something's missing” or “after a while there's no clear goal”, or do you "
     "feel it's fine — just the gameplay is enough for me?"),
    ("q", "加剧情的影响（回扣题材）", "Q26", "口头回答",
     "有些二合游戏会给你一个故事或一个谜（比如“这家店背后藏着什么秘密”），"
     "让你一边合成一边琢磨后续。如果给这款游戏加一点剧情，比如给游乐场编个小故事、"
     "或者给你一个想达成的目标，你觉得会不会让你更愿意继续玩？为什么？",
     "Some merge games give you a story or a mystery to follow (like “what secret is behind "
     "this shop”), so you merge while wanting to see what happens next. If this game added a "
     "bit of story — like a small backstory for the carnival or a goal for you to work "
     "toward — do you think it would make you more willing to keep playing? Why?"),

    ("page", "Task B · Page 9", "开放建议、优先级排序与收尾致谢"),
    ("sub", "最后几道开放问题，想到什么就说什么，好的坏的都可以。",
     "A few last open questions — say whatever comes to mind, good or bad."),
    ("q", "开放建议", "Q27", "口头回答",
     "这款游戏你还有其他想说的吗？或者有什么建议？随便聊聊。",
     "Is there anything else you'd like to say about this game? Any suggestions? "
     "Feel free to share."),
    ("q", "优先级排序", "Q28", "排序",
     "在你刚才提到的所有问题 / 建议里，你觉得最该先解决哪一个？"
     "请把下面几类按“最该先解决”到“可以后解决”排个序（1 = 最该先解决）。",
     "Of all the issues or suggestions you've mentioned, which do you think should be fixed "
     "first? Please rank the following from “most urgent to fix” to “can be fixed later” "
     "(1 = most urgent)."),
    ("opt", "3D 棋盘不好上手 / 不好看清", "The 3D board is hard to pick up / hard to see clearly"),
    ("opt", "棋子分不清、合成链路看不懂", "The pieces are hard to tell apart / the merge chain is confusing"),
    ("opt", "美术风格偏幼稚 / 不够精致", "The art style feels childish / not polished enough"),
    ("opt", "游乐场题材和棋盘脱节", "The carnival theme feels disconnected from the board"),
    ("opt", "没有剧情、缺少目标感", "No story / lacks a sense of goal"),
    ("opt", "其他", "Other"),
    ("q", "收尾总结", "Q29", "口头回答",
     "今天非常感谢你的参与！最后想请你总结一句：这款游戏让你印象最深的一件事是什么？"
     "好坏都可以。",
     "Thank you so much for your time today! As a final thought: what's the one thing about "
     "this game that stood out to you most — good or bad?"),
]

OUTLINE_MODERATED = [
    ("page", "环节 1 · 开场与资格核验", "开场介绍，确认设备与 Gossip Harbor 等级"),
    ("kv", "提问话术",
     "你好，谢谢你今天抽时间参加。我是今天的主持人［姓名］，接下来大概 90 分钟，"
     "前半段请你玩一款游戏，后半段我们聊聊你的感受。\n"
     "先说几件事：你觉得不舒服的地方、觉得奇怪的地方，都可以直接说，没有标准答案，"
     "我们就是想听真实的想法。\n"
     "开始之前，我想先确认两件事。第一，你今天用的是什么手机？是安卓的对吗？"
     "第二，麻烦你打开《Gossip Harbor》，让我看一下你现在的主界面——主要是确认一下等级。",
     "Hi, thanks for making time today. I'm [name], your moderator. We've got about 90 "
     "minutes: the first part you'll spend playing a game, then we'll talk about your "
     "experience.\n"
     "A few things first: if anything feels off or strange, just say so. There are no right "
     "answers; we just want your honest take.\n"
     "Before we start, two quick checks. First, what phone are you using today — Android, "
     "right? Second, could you open Gossip Harbor and show me your main screen? I just need "
     "to confirm your level."),
    ("obs", "· 设备型号与系统版本，确认设备符合要求\n"
            "· Gossip Harbor 当前等级（需 ≥50，硬门槛，不达标则终止并说明原因）\n"
            "· 玩家上机前的状态：紧张还是放松，有没有提问"),

    ("page", "环节 2 · 试玩说明", "讲清边玩边说、主持人不会打扰"),
    ("kv", "提问话术",
     "接下来请你玩一款游戏，大约 30 分钟。这款游戏刚上线不久，你可能还没接触过。"
     "我们想听你最真实的感受——觉得好、觉得不好，都可以说。\n"
     "玩的时候，如果你脑子里冒出什么想法——觉得有意思、觉得困惑、或者单纯想说一句——"
     "都可以随时说出来，我会在旁边听着记下来。但我要提醒一句："
     "这 30 分钟里我不会打断你，也不会回答关于游戏怎么玩的问题，你就当自己一个人玩。"
     "如果全程不想说话，也完全没问题。\n"
     "游戏问题（卡住、闪退之类）可以直接跟我说，其他情况我就不打扰你了。"
     "准备好了吗？那我们开始。",
     "Next you'll play a game for about 30 minutes. It launched recently, so you may not "
     "have tried it yet. We want your honest reaction — what you like and what you don't.\n"
     "While you play, if anything pops into your head — something fun, something confusing, "
     "or just a thought you want to say — feel free to say it out loud. I'll be listening "
     "and taking notes. But one thing: I won't interrupt you for these 30 minutes, and I "
     "won't answer questions about how to play — just treat it like you're playing on your "
     "own. If you'd rather stay quiet the whole time, that's completely fine too.\n"
     "If you hit a game problem (freezes, crashes), just tell me. Otherwise I'll leave you "
     "to it. Ready? Let's go."),
    ("obs", "· 玩家是否理解“可以说话但不被追问”这个设定\n"
            "· 玩家此前是否玩过或见过这款游戏（若玩过，记下他玩到什么程度，"
            "分析时要区分“初次接触”与“有先前经验”两类样本）"),

    ("page", "环节 3 · 静默试玩观察", "30 分钟静默观察，记录未经追问的第一反应"),
    ("kv", "主持人动作",
     "不主动开麦，不主动追问。看着玩家的录屏，按时间点记笔记。",
     "Don't unmute, don't probe, don't draw attention. Watch their screen and take "
     "timestamped notes."),
    ("obs", "· 新手教程阶段：有没有卡住、有没有跳过、跳过后是否迷茫\n"
            "· 第一次看到 3D 棋盘时的即时反应（原话、表情、停顿）\n"
            "· 拖拽合成的动作是否犹豫、误操作多不多\n"
            "· 有没有注意到游乐场、建地块、挂机收入这些外围部分，什么时候注意到的\n"
            "· 有没有主动提到“没有剧情”\n"
            "· 情绪起伏点：什么时候显得投入、什么时候显得无聊或烦躁\n"
            "· 玩家主动说出的所有形容词与比喻，逐字记"),

    ("page", "环节 4 · 背景信息", "了解设备、玩过哪些二合、日均时长与付费情况"),
    ("q", "了解玩家的常用设备", "Q1", "口头回答",
     "我们先聊点背景。除了今天这台手机，你平时还主要用什么设备玩休闲游戏？",
     "Let's start with some background. Besides the phone you're using today, what device "
     "do you mainly play casual games on?"),
    ("fu", "· 平板和手机分别玩什么？为什么分开玩？\n· 有没有在电脑上玩过？"),
    ("obs", "设备偏好，以及“手机”这个答案背后的使用场景（躺着玩、通勤玩）"),

    ("q", "确认玩过哪些二合游戏", "Q2", "多选",
     "接下来我念一串二合游戏的名字，你听到玩过的就告诉我。过去 3 个月里，"
     "下面这些你玩过哪些？\n"
     "Travel Town、Merge Mansion、Gossip Harbor、Merge Cooking、Flambé、Tasty Travels、"
     "Merge Dragons、Love & Pies、Merge Miners。",
     "I'll read a list of merge games. Tell me if you've played any of them. Over the past "
     "3 months, which of these have you played?\n"
     "Travel Town, Merge Mansion, Gossip Harbor, Merge Cooking, Flambé, Tasty Travels, "
     "Merge Dragons, Love & Pies, Merge Miners."),
    ("fu", "· 清单之外还玩过别的二合吗？\n· 这几款里，有没有玩过就卸载的？"),
    ("obs", "玩家玩过的二合广度，为下一环节的基准游戏提供候选"),

    ("q", "画像分层：日均游戏时长", "Q3", "单选",
     "过去 7 天里，你平均每天大概花多长时间玩二合游戏？",
     "Over the past 7 days, about how much time did you spend playing merge games on an "
     "average day?"),
    ("fu", "· 一般什么时候玩？在什么情况下玩？（睡前、通勤、做家务的间隙）\n"
           "· 最近一周的时长和平时比，是多了还是少了？为什么？\n"
           "· 一般一次玩多久？一天玩几次？"),
    ("obs", "时长档位（几乎不玩 / 10-30 分钟 / 30 分钟-1 小时 / 1-2 小时 / 2 小时以上） + 使用场景，两者都要记"),

    ("q", "画像分层：月度付费金额", "Q4", "单选",
     "过去 1 个月里，你在二合游戏上大概花了多少钱？比如买金币、道具、通行证这类。",
     "Over the past month, roughly how much did you spend on merge games? Things like "
     "coins, items, or a battle pass."),
    ("fu", "· 最近一次花钱，买的是什么？\n· 那次是什么触发的？（卡关了、有限时活动、想快点升级）"),
    ("obs", "金额档位（没花过钱 / $1-9 / $10-49 / $50-99 / $100 及以上） + 最近一次付费的具体内容"),

    ("q", "付费动机与不付费原因", "Q5", "口头回答",
     "（主持人按玩家上一题的实际情况选问一支。）\n"
     "若玩家没花过钱：是什么让你没有花钱？觉得没必要，还是一直没碰到想买的？"
     "你在别的二合游戏里花过钱吗？\n"
     "若玩家花过钱：你一般在什么情况下最容易掏钱？比如卡关了、有限时活动、想快点升级。",
     "(The moderator asks the branch that fits the participant's answer to the previous "
     "question.)\n"
     "If they haven't spent: What's kept you from spending? Is it that you don't see the "
     "need, or that nothing has tempted you yet? Have you spent money in other merge games?\n"
     "If they have spent: What tends to make you pull the trigger? Like being stuck on a "
     "level, a limited-time event, or wanting to level up faster."),
    ("fu", "· 什么样的东西会让你愿意花钱？\n"
           "· 在别的二合游戏里花过钱的话，那是什么情况？\n"
           "· 有没有哪次花完觉得后悔？"),
    ("obs", "分辨“不付费是主动选择，还是没被触发”；付费行为在不同游戏之间是否一致"),

    ("page", "环节 5 · 基准游戏", "选出与今天这款最像的游戏"),
    ("q", "定基准：选出最像今天这款的游戏", "Q6", "单选",
     "你刚才玩了今天这款，也回顾了之前玩过的那些二合游戏。"
     "你觉得下面哪一款和今天这款最像？可以从玩法、棋盘、题材、画面这些方面想想。\n"
     "Travel Town、Merge Mansion、Gossip Harbor、Merge Cooking、Flambé、Tasty Travels、"
     "Merge Dragons、Love & Pies、Merge Miners，或者都不像。",
     "You just played today's game and looked back over the merge games you've played. "
     "Which of these do you think is most like today's game? Think about gameplay, the "
     "board, the theme, the visuals.\n"
     "Travel Town, Merge Mansion, Gossip Harbor, Merge Cooking, Flambé, Tasty Travels, "
     "Merge Dragons, Love & Pies, Merge Miners — or none of them."),
    ("obs", "玩家选的基准游戏（务必记在纸上，后面 Q9、Q13、Q17、Q19、Q25 会反复引用）"),

    ("q", "基准理由：说明像在哪里", "Q7", "口头回答",
     "为什么觉得这款跟今天的最像？能说说到底是哪方面像吗？",
     "Why do you think that one is most like today's game? What's similar about it?"),
    ("fu", "· 玩法上像在哪里？棋盘上像在哪里？\n"
           "· 画面或题材上像不像？\n"
           "· 如果刚才选了“都不像”，那你心里想的是哪一款？"),
    ("obs", "像的层面（玩法 / 棋盘 / 题材 / 画面），这是判断基准是否有效的关键"),

    ("page", "环节 6 · 整体感受与打分", "满意度、继续意愿、下载意愿、NPS 与受众"),
    ("q", "试玩整体满意度", "Q8", "打分",
     "想想你刚才试玩的感受，给今天这款游戏的整体满意度打个分。"
     "1 分最低，10 分最高。你会打几分？",
     "Thinking about your experience just now, how satisfied are you overall with today's "
     "game? 1 is the lowest, 10 is the highest. What would you give it?"),
    ("fu", "为什么给这个分？哪里让你满意、哪里拉低了分？"),
    ("obs", "分数 + 理由里的具体维度"),

    ("q", "基准游戏满意度（对照）", "Q9", "打分",
     "再想想过去 3 个月里，［主持人复述玩家选的基准游戏］给你的整体感受，也给它打个分。"
     "1 分最低，10 分最高。你会打几分？",
     "Now thinking about the past 3 months with [repeat the baseline game they picked], "
     "how satisfied are you overall with it? 1 is the lowest, 10 is the highest. What would "
     "you give it?"),
    ("fu", "这个分和刚才那款差在哪里？"),
    ("obs", "分数。两个分相减就是尖叫度得分。"),

    ("q", "继续体验意愿", "Q10", "打分",
     "假设今天不是测试、也没有任何奖励，就凭你刚才的体验——"
     "明天你还想继续玩这款游戏的可能性有多大？1 分代表肯定不会，10 分代表肯定会。",
     "Forget that this is a test and there's a reward — just going on your experience, how "
     "likely are you to keep playing this game tomorrow? 1 means definitely not, 10 means "
     "definitely yes."),
    ("fu", "什么会让你明天打开它？什么会让你不想打开？"),
    ("obs", "分数，注意这是“可能性”量表，不是 NPS"),

    ("q", "下载意愿", "Q11", "单选",
     "这款游戏现在已经能在应用商店下载了。假设你还没下载，"
     "你下载它的意愿是下面哪一种——肯定会下载、可能会下载、不确定、"
     "可能不会下载、肯定不会下载？",
     "This game is already available to download on the app store. Imagine you hadn't "
     "downloaded it yet — how likely would you be to download it: definitely would, "
     "probably would, not sure, probably wouldn't, or definitely wouldn't?"),
    ("fu", "为什么这么选？题材或 3D 有没有影响你的决定？"),
    ("obs", "档位（肯定会下载 / 可能会下载 / 不确定 / 可能不会下载 / 肯定不会下载）+ 理由"),

    ("q", "推荐意愿（NPS）", "Q12", "NPS",
     "今天这款游戏，你愿意推荐给朋友或家人吗？从 0 到 10 分，"
     "0 分代表完全不愿意推荐，10 分代表非常愿意推荐。你会打几分？",
     "How likely are you to recommend today's game to a friend or family member? From 0 to "
     "10, where 0 means you definitely wouldn't recommend it and 10 means you definitely "
     "would. What would you give it?"),
    ("obs", "分数"),

    ("q", "推荐意愿对照（NPS）", "Q13", "NPS",
     "同样地，你常玩的那款［主持人复述基准游戏］，你愿意推荐给朋友或家人吗？"
     "从 0 到 10 分，你会打几分？",
     "Same question for the game you usually play, [repeat the baseline game]. How likely "
     "are you to recommend it to a friend or family member? From 0 to 10, what would you "
     "give it?"),
    ("obs", "分数。两款分开算推荐者（9-10）、被动者（7-8）、贬损者（0-6）"),

    ("q", "受众圈层", "Q14", "口头回答",
     "你觉得今天这款游戏主要是给什么样的人玩的？你会推荐给身边什么样的朋友？"
     "是大多数朋友都合适，还是只有一部分爱玩这类游戏的朋友才合适？那你常玩的那款呢？",
     "Who do you think today's game is mainly for? Who would you recommend it to? Would "
     "most of your friends be a good fit, or only some friends who like this kind of game? "
     "And how about the game you usually play?"),
    ("fu", "· 为什么是这类人，而不是别人？\n"
           "· 有没有哪类人你会特别想推荐，哪类人你觉得不适合？"),
    ("obs", "玩家用什么词描述目标人群（年龄、性别、性格、生活状态），逐字记"),

    ("page", "环节 7 · 3D 棋盘深挖", "3D 棋盘上手难度、棋子辨识度与合成手感"),
    ("q", "3D 棋盘上手感受（定量）", "Q15", "矩阵",
     "回想一下你刚上手这个棋盘时的感受。我念几句话，你告诉我有多同意，"
     "1 分很不同意，5 分很同意。\n"
     "· 我一上手就能理解这个棋盘怎么玩\n"
     "· 这个立体的棋盘让我更感兴趣\n"
     "· 这个立体的棋盘有点乱、不太好看清\n"
     "· 我不知道东西在哪、该从哪下手",
     "Think back to when you first picked up this board. I'll read a few statements and you "
     "tell me how much you agree, 1 being strongly disagree and 5 strongly agree.\n"
     "· I understood how to play on this board right away\n"
     "· The 3D board made me more interested\n"
     "· The 3D board felt a bit messy / hard to see clearly\n"
     "· I wasn't sure where things were or where to start"),
    ("fu", "· 这四句话里面，哪一条最符合你的想法？为什么？\n"
           "· 刚上手那几分钟，你实际是怎么摸索的？\n"
           "· 现在比一开始清楚一些了吗，还是依然有点乱？"),
    ("obs", "四个子维度的分数（上手理解度 / 3D 吸引力 / 3D 可读性 / 上手迷茫）"),

    ("q", "棋子合成链路与辨识度", "Q16", "口头回答",
     "这个游戏里，从最普通的棋子一步步合成出更高级棋子，这条链子你能看明白吗？"
     "比如先合什么、再合什么、最后能出什么。这些棋子相互之间差别够大吗？"
     "有没有长得太像、容易看混的？",
     "In this game, you merge up from the most basic piece to more advanced ones — can you "
     "follow that chain? Like what merges into what, and what you end up with. Are the "
     "pieces different enough from each other? Any that look too similar and are easy to "
     "mix up?"),
    ("fu", "· 你是靠什么分辨棋子的？颜色、形状，还是名字？\n"
           "· 有没有哪两个棋子你反复看混？具体是哪两个？\n"
           "· 合成链你玩了多久才摸清？是教程教的，还是自己试出来的？\n"
           "· 如果看不明白，是什么让你卡住？"),
    ("obs", "能不能说清合成链；具体点名了哪几个棋子看混"),

    ("q", "与基准游戏对比及合成手感", "Q17", "口头回答",
     "跟你常玩的那款［复述基准游戏］比，今天这个棋盘是更顺手还是更别扭？"
     "拖棋子去合成的时候手感怎么样？会不会拖不动、放不准、不好对齐？"
     "合成成功时的反馈，比如特效、声音、动作，让你觉得爽快还是平淡？",
     "Compared with the game you usually play, [repeat the baseline game], does today's "
     "board feel smoother or more awkward? How does dragging pieces to merge feel? Does it "
     "drag properly, land accurately, line up well? When a merge succeeds, the feedback — "
     "effects, sound, motion — does it feel satisfying or flat?"),
    ("fu", "· 更顺手的是哪部分？更别扭的是哪部分？\n"
           "· 那种别扭是“还不习惯、玩久了就好”，还是“就是不喜欢”？\n"
           "· 如果让你把 3D 改成 2D 平面，你会愿意吗？为什么？"),
    ("obs", "分辨 3D 是“新奇的吸引力”还是“陌生的劝退感”；上手不适是短期的还是根本性的"),

    ("page", "环节 8 · 美术风格深挖", "美术风格感知、适龄性与基准对比"),
    ("q", "美术风格感知与适龄性", "Q18", "口头回答",
     "刚才游戏里那些棋子和角色，你觉得它长什么样、什么风格？"
     "如果让你用 1-2 个词形容，你会用哪两个词？这个风格，"
     "你觉得适合你这个年纪的人玩吗，还是有点偏小、偏幼稚？",
     "The pieces and characters in the game — how would you describe how they look, and "
     "their style? If you had to pick 1-2 words, which would you use? Do you think this "
     "style suits someone your age, or does it feel a bit young or childish?"),
    ("fu", "· 如果觉得偏幼稚，具体是哪里让你有这种感觉？角色、棋子、颜色，还是整体？\n"
           "· 如果觉得合适，那你平时玩的那款在风格上有什么不一样？\n"
           "· 这个风格会不会影响你愿不愿意继续玩？\n"
           "· 你会怎么形容它的质感？塑料感、粘土感，还是精致？"),
    ("obs", "玩家用的形容词汇；分辨反感来自“太幼稚”还是“不够精致、不够高级”"),

    ("q", "与基准游戏美术对比", "Q19", "口头回答",
     "跟你常玩的那款［复述基准游戏］比，你觉得哪个画面更高级、更精致？"
     "你更喜欢哪一个的风格？为什么？",
     "Compared with the game you usually play, [repeat the baseline game], which one looks "
     "more premium or more polished? Which style do you prefer? Why?"),
    ("fu", "· “更精致”具体体现在哪里？细节、光影，还是整体的统一感？\n"
           "· 如果今天这款的美术要改，你最想先改什么？"),
    ("obs", "玩家原话里对美术“高级 vs 幼稚”的具体判断词"),

    ("page", "环节 9 · 题材与玩法深挖", "题材、建地块与放置玩法的感知与理解"),
    ("q", "游乐场题材感知", "Q20", "口头回答",
     "你第一眼看到“游乐场”这个主题，什么感觉？觉得新鲜、有意思，"
     "还是跟平常玩的差别太大、有点怪？你觉得游乐场这个题材，"
     "适不适合做成一款你平时会玩的二合游戏？为什么？",
     "When you first saw the “carnival” theme, what did you think? Did it feel fresh and "
     "interesting, or too different from what you usually play — a bit odd? Do you think "
     "the carnival theme works as a merge game you'd play regularly? Why?"),
    ("fu", "· 你平时玩的那些是什么题材？（餐厅、小镇、花园）跟游乐场比怎么样？\n"
           "· 如果换成别的题材，你会更想玩吗？换成什么？\n"
           "· 游乐场这个题材对你来说是加分，还是无所谓？"),
    ("obs", "对题材的第一反应与理由"),

    ("q", "建地块 / 外围玩法理解", "Q21", "口头回答",
     "你刚才有没有注意到，游戏里有个花金币搭游乐场、激活格子的部分？"
     "你明白这部分是干什么的吗？知不知道激活格子之后会发生什么、怎么升级、怎么赚金币？"
     "有哪里让你觉得不清楚？",
     "Did you notice the part of the game where you spend coins to build up the carnival and "
     "activate tiles? Do you understand what that part is for? Do you know what happens "
     "after you activate a tile, how it upgrades, how you earn coins? Anything that felt "
     "unclear?"),
    ("fu", "· 你是什么时候注意到这个部分的？是教程带你的，还是自己发现的？\n"
           "· 你觉得它跟合成棋盘是一回事，还是两回事？\n"
           "· 如果一直搞不懂，你会不会干脆不管它？"),
    ("obs", "对“花钱激活格子 → 解锁设施 → 增加收入”这条循环的理解度"),

    ("q", "题材与棋盘关联度", "Q22", "口头回答",
     "你觉得游乐场这个题材，和你刚玩的那个棋盘（合成解谜），联系紧密吗？"
     "还是这两部分有点脱节、各玩各的？如果让你改，你会让它俩怎么配合更好？",
     "Do you feel the carnival theme and the board you just played (the merge puzzle) are "
     "closely connected? Or do they feel a bit disconnected, each doing its own thing? If "
     "you could change it, how would you make them work together better?"),
    ("fu", "· 你觉得合成出来的东西，跟游乐场有什么关系吗？\n"
           "· 如果让你设计，你希望这两块怎么连起来？"),
    ("obs", "玩家是否感知到题材与核心玩法的断裂"),

    ("q", "放置玩法感知", "Q23", "口头回答",
     "游戏里好像有挂机、放置让你能拿到收入，比如放着自动收钱。你注意到了吗？"
     "你觉得“放着也能赚钱”这个玩法，跟你理解的二合游戏顺不顺？"
     "是更有动力，还是感觉跟二合没关系？",
     "The game seems to have idle or auto-income, like collecting money automatically while "
     "you're away. Did you notice that? Does “earning while you're not playing” fit with "
     "your idea of a merge game? Does it make you more motivated, or does it feel unrelated "
     "to merging?"),
    ("fu", "· 你平时玩的二合有类似的设计吗？\n"
           "· 这个设计会让你更想回来看看，还是觉得可有可无？"),
    ("obs", "分辨“游乐场 + 放置”是差异化亮点，还是与核心棋盘脱节的劝退点"),

    ("page", "环节 10 · 剧情深挖", "剧情缺失的感知及对投入感的影响"),
    ("q", "对剧情缺失的最初感知", "Q24", "口头回答",
     "你玩的时候，有没有意识到“这游戏里没什么故事、剧情”？还是压根没在意这回事？",
     "While you were playing, did you notice that this game doesn't really have a story or "
     "narrative? Or did that not cross your mind at all?"),
    ("fu", "· 如果没在意，是因为合成棋盘本身就够吸引你吗？\n"
           "· 你平时玩的那款有剧情吗？你会在意吗？"),
    ("obs", "对剧情缺失的感知强度"),

    ("q", "剧情对投入感与心流的影响", "Q25", "口头回答",
     "你玩过的那些二合游戏，比如你选的基准游戏，有没有剧情？如果有，"
     "那些剧情会不会让你更投入、更想一直玩下去？反过来，那些没有剧情的二合游戏，"
     "你玩的时候会不会觉得“少了点什么”“玩着玩着就没目标了”，"
     "还是觉得无所谓、纯玩机制就挺好？",
     "The merge games you've played — like the baseline game you picked — do they have a "
     "story? If so, does the story make you more engaged, more wanting to keep playing? On "
     "the other hand, for merge games without a story, do you feel like “something's "
     "missing” or “I run out of goals,” or do you feel it doesn't matter and the mechanics "
     "alone are enough?"),
    ("fu", "· 有没有哪款游戏的剧情让你印象特别深？是什么？\n"
           "· 你会为了看剧情而多玩一会儿吗？\n"
           "· “没目标”具体是指什么？不知道下一步做什么，还是不知道为什么要做？"),
    ("obs", "玩家的核心动机是纯机制驱动，还是依赖剧情 / 目标牵引"),

    ("q", "加剧情的影响（回扣题材）", "Q26", "口头回答",
     "有些二合游戏会给你一个故事或一个谜，比如“这家店背后藏着什么秘密”，"
     "让你一边合成一边琢磨后续。如果给这款游戏加一点剧情，比如给游乐场编个小故事，"
     "或者给你一个想达成的目标，你觉得会不会让你更愿意继续玩？为什么？",
     "Some merge games give you a story or a mystery — like “what secret is this shop "
     "hiding?” — so you merge while wondering what happens next. If this game added a bit "
     "of story, like a small backstory for the carnival, or a goal to work toward, do you "
     "think you'd want to keep playing more? Why?"),
    ("fu", "· 什么样的故事你会想看？经营类、悬疑，还是人物关系？\n"
           "· 如果加了剧情但玩法不变，你会多玩多久？\n"
           "· 还是说你更希望他们把精力放在棋盘本身？"),
    ("obs", "剧情包装对持续动力的作用，以及玩家更看重机制还是包装"),

    ("page", "环节 11 · 收尾", "开放建议、优先级排序与致谢"),
    ("q", "开放建议", "Q27", "口头回答",
     "这款游戏你还有其他想说的吗？或者有什么建议？随便聊聊。",
     "Is there anything else you'd like to say about this game? Any suggestions? Say "
     "whatever comes to mind."),
    ("obs", "玩家主动提出的、我们没问到的点，往往是最有价值的"),

    ("q", "优先级排序", "Q28", "排序",
     "在你刚才提到的所有问题、建议里，你觉得最该先解决哪一个？"
     "我念几类，你按“最该先解决”到“可以后解决”排个序。\n"
     "3D 棋盘不好上手、不好看清；棋子分不清、合成链路看不懂；"
     "美术风格偏幼稚、不够精致；游乐场题材和棋盘脱节；没有剧情、缺少目标感；其他。",
     "Of all the issues or suggestions you've mentioned, which do you think should be fixed "
     "first? I'll read a few categories and you rank them from “most urgent to fix” to “can "
     "be fixed later.”\n"
     "The 3D board is hard to pick up or hard to see clearly; the pieces are hard to tell "
     "apart or the merge chain is confusing; the art style feels childish or not polished "
     "enough; the carnival theme feels disconnected from the board; no story or lacks a "
     "sense of goal; other."),
    ("fu", "为什么把它排在第一位？"),
    ("obs", "排序结果与理由"),

    ("q", "收尾总结", "Q29", "口头回答",
     "今天非常感谢你的参与！最后想请你总结一句："
     "这款游戏让你印象最深的一件事是什么？好坏都可以。",
     "Thank you so much for your time today! One last thing: what's the one thing about "
     "this game that stood out most to you? Good or bad, either is fine."),
    ("obs", "玩家最终留下的记忆点"),
]

# ---------------------------------------------------------------- 表清单
# sheets 的顺序就是工作表顺序。三类都可放任意多个实例，也可以不放。
META = {
    "filename": "【2609】3D二合测试访谈V3.xlsx",
    "sheets": [
        {"type": "needs", "name": "需求梳理", "content": NEEDS},
        {"type": "survey", "name": "玩家筛选问卷", "content": SURVEY},
        {"type": "outline", "name": "无主持人访谈大纲",
         "content": OUTLINE_UNMODERATED},
        {"type": "outline", "name": "有主持人访谈大纲",
         "content": OUTLINE_MODERATED, "columns": COLUMNS_8},
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
    sheet = Sheet(workbook, spec["name"], resolve_widths(spec),
                  resolve_tab_color(spec, 0), en_columns=(4,))
    sheet.add([("题型", "header", 1), ("题号", "header", 1),
               ("中文", "header", 1), ("English", "header", 1), ("判定", "header", 1)])
    for item in spec["content"]:
        kind = item[0]
        if kind == "section":
            sheet.add([(item[1], "section", 5)])
        elif kind == "prose":
            for cn, en in item[1]:
                sheet.add(blank(2) + [(cn, "body", 1), (en, "body", 1), ("", "body", 1)])
        elif kind == "q":
            _, qtype, qno, stem_cn, stem_en, judge, options = item
            first = sheet.row + 1
            sheet.add([(qtype, "label", 1), (qno, "label", 1),
                       (stem_cn, "body", 1), (stem_en, "body", 1), (judge, "judge", 1)])
            for cn, en, mark in options:
                sheet.add(blank(2) + [(cn, "body", 1), (en, "body", 1),
                                      (mark, "stop" if mark else "body", 1)])
            sheet.merge_down(1, first, sheet.row)
            sheet.merge_down(2, first, sheet.row)
            sheet.mark_block_top(first)
        else:
            raise ValueError(f"玩家筛选问卷不认识的类型：{kind}")
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


def main():
    output = sys.argv[1] if len(sys.argv) > 1 else META["filename"]
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
        elif kind == "outline":
            build_outline(workbook, spec, outline_index)
            outline_index += 1
        else:
            raise ValueError(f'不认识的表类型：{kind}（应为 needs / survey / outline）')
    workbook.save(output)

    print("已生成:", output)
    for sheet in workbook.worksheets:
        print(f"  {sheet.title}: {sheet.max_row} 行 x {sheet.max_column} 列")

    # 自动校验（需同目录有 validate_xlsx.py）
    try:
        import validate_xlsx
        _, problems = validate_xlsx.check(output)
        if problems:
            print(f"\n校验发现 {len(problems)} 个问题：")
            for sheet_name, coord, kind, detail in problems[:60]:
                print(f"  [{sheet_name}] {coord} {kind} {detail}")
        else:
            print("\n校验通过：无裁切、无合并重叠、无 Markdown 残留。")
    except ImportError:
        print("\n（未找到 validate_xlsx.py，跳过自动校验；"
              "请把技能里的 scripts/validate_xlsx.py 复制到工作目录）")


if __name__ == "__main__":
    main()
