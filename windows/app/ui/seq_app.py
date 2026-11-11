#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF转换拼接工具 v6.2 —— 界面层（tkinter）

⭐ 6.0 起是**三个模式**（顶部三个大按钮切换）：
    导出 PDF   ／  PDF 首页替换   ／  合并打印
   第一个是横版量产（一批 → 各自一份 PDF）；
   第二个是把签好字的扫描页按顺序贴回首页；
   第三个是把一批文件拼成一份、直接拿去打印（中间产物只在临时目录里活一下）。
   「一个模式有什么脾气」全在 MODE_CAPS 表里 —— 加模式 / 改脾气只动那张表。

下面是第二个模式（PDF 首页替换）的流程，v3.0 起就是**两段式**的：
    ① 在「原件来源」里把要处理的 PDF 挑进来（一次可多选；散在几个文件夹里
       就分几次点，从一个夹里挑一批、再进下一个夹挑一批）
       → 点「预览」→ 表格按**文件名**排好队给你看（这一步**不需要**扫描件）；
    ② 照着表格的顺序把纸质件摊开去扫 → 回来选「扫描件 PDF」→ 再点一次「预览」
       → 叠上扫描页对应关系，调页序 / 页码 / 方向 → 点「开始替换」。

和 v1 的界面保持一致（悬浮说明、右键菜单、记住上次选择），
只是把「名字匹配」换成「按顺序硬对应」，并多了「上移 / 下移」调序。

工具栏的 ①②③ 三个模块各占一行 —— 编号只留给这三个模块，上面几块不再编号，
免得界面上出现两个「①」。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk
from typing import NamedTuple

_APP = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP not in sys.path:
    sys.path.insert(0, _APP)
import paths as _paths  # noqa: E402
_paths.install()
import seq_core as C  # noqa: E402
import merge_core as M  # noqa: E402

# ── 「把文件拖进来」（可选能力，见 enable_dnd / App._setup_dnd）──────────────
# ⭐ 为什么这里只 import、只置标志位，绝不在这里报错退出：
#    拖放是 tkinter **自己没有**的能力，靠第三方库 tkinterdnd2 提供。而它
#    **不是必需品** —— 依赖清单是这一版才把 tkinterdnd2 加进去的，所以在
#    用户重装环境之前，老环境里根本没有这个包。那种情况下工具必须照样
#    能开、能点按钮挑文件，只是少一个方便。（用户 2026-10-09 要的机制。）
try:
    import tkinterdnd2 as _dnd

    # ⚠️ `require` 在**子模块**里，包顶层没有 —— 写成 `_dnd.require(...)`
    #    会 AttributeError（踩过）。它能给**已经建好的**窗口补上拖放能力，
    #    正是我们要的：不必换成 `TkinterDnD.Tk()`。
    from tkinterdnd2.TkinterDnD import require as _dnd_require

    DND_FILES = _dnd.DND_FILES
    _DND_ERR = ""
except Exception as _exc:  # noqa: BLE001 —— 没装 / 装了但平台的 tkdnd 库缺失
    _dnd = None
    _dnd_require = None
    DND_FILES = "DND_Files"
    _DND_ERR = str(_exc)

# `tkinterdnd2.require(root)` 真的成功过没有。⚠️ 光 import 成功不算 ——
# 包在、但 tkdnd 的二进制跟这台机器的 Tcl/Tk 对不上时，require 会抛异常。
_DND_READY = False

APP_TITLE = f"{C.TOOL_NAME}  v{C.TOOL_VERSION}"
APP_SUBTITLE = "批量转 PDF　·　把签字页贴回首页　·　多份拼成一份"

# ══════════════════════════════════════════════════════════════════════════
# 统一 UI 色板（4.0 起）
# ══════════════════════════════════════════════════════════════════════════
# 出处：`vibecoding成品/项目i5统一UI美学/浅色系UI风格指南.md`，
#       基准作品 = PDF 差异对比工具 5.0（用户：「那个 UI 很完美」）。
# ⚠️ 不许在这里自创第三套色板。要加颜色先看指南第 2 节有没有现成的。
# ⚠️ 主色 PRIMARY 只许出现在**一颗实心按钮**上（面积 ≤5%）——
#    2026-10-22 把饱和蓝铺成整条区块带/表头，当场被否：「蓝色饱和度有点高」。
FONT = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"
BG = "#f5f6f8"          # 窗口底
CARD = "#ffffff"        # 卡片 / 输入框 / 表格
DZ_BG = "#f7f9fc"       # 块中块（唯一输入块常态底）
DZ_BG_ON = "#eaf1ff"    # 拖入悬停 / 列表选中
STATUS_BG = "#eceef1"   # 底部状态条（比窗口底略深，形成收口）
BORDER = "#d9dee5"      # 全局唯一描边色，1px
TEXT = "#2a2f36"        # 正文（不用纯黑）
TEXT2 = "#7a828c"       # 次要文字
MUTED = "#9aa1a9"       # 弱提示（9 号）
DISABLED_FG = "#c3c8ce"  # 禁用文字
PRIMARY = "#2f6fed"     # 主色（只给一颗按钮）
PRIMARY_D = "#2257c4"   # 主色深：强调字、hover
HEAD_BG = "#eaf1ff"     # 表头 tint（与总表 / Excel 同一套）
HEAD_FG = "#2257c4"     # 表头文字（tint 底上不能用 PRIMARY，对比不够）
SPINE = "#c6d9f8"       # 浅蓝细脊线（选中块的描边）
BAR = "#9CC0F5"         # 进度条
ZEBRA = "#f7f9fc"       # 斑马纹
HOVER = "#f0f2f5"       # 次按钮 hover
WARN = "#b26a00"        # 提醒（不是错误）
DANGER = "#c0392b"      # 止损 / 失败
OK_GREEN = "#1b6e45"    # 成功
# 和 v1 用不同的文件名，免得两个版本的记录互相踩
CONFIG_PATH = Path.home() / ".pdf_batch_cover_swap.json"
# 5 = 6.2 加了 page_orient（合并打印的页面方向）。老配置照读，缺的键走默认值。
CONFIG_VERSION = 5
# ── 配置里到底记什么：**白名单**，没列进来的一律不落盘 ────────────────────
# ⭐ 2026-10-08 用户要求：路径类的东西一个都不记。原话 ——
#    「以前是选文件夹嘛，所以说要记住那个位置。但是现在不是了，现在这个
#    文件夹又要变、那个扫描件要变，所以不用记住这些东西」。
#    而且记住反而更糟：上一轮的 PDF 早被删掉 / 改名了，下次打开还指着它，
#    「它找不到，它就会卡在那里」。所以「原件来源」清单和「扫描件 PDF」
#    都不在表里 —— 每次打开都是空的，重新挑一次即可。
#    表里留下的是**跟「这一批处理」无关**的东西，记住了真的省事。
#    ⚠️ "suffix" 与 "conv_suffix" 是**两个模式各记各的**「输出文件名后缀」
#       （见 SUFFIX_VAR_BY_MODE）—— 2026-10-08 用户明确要求它们不共享记忆。
CONFIG_PERSIST_VARS = ("suffix", "conv_suffix", "scale", "reverse",
                       "print_mode", "page_orient")
# 窗口太小就不恢复上次的大小（会把状态栏挤出去）；下面 default_geometry 也拿它兜底
GEOM_MIN = (980, 700)
# 初始窗口大小：要装得下「来源清单 + 操作行 + ①②③ 三行工具栏 + 表格」，不靠最大化
GEOM_DEFAULT = (1140, 860)
SAVE_DEBOUNCE_MS = 900
# 「扫描件 PDF」那一格是**手打也认**的输入框，绑着 StringVar —— 敲一个字符就触发一次。
# 选完扫描件要自动补配对（见 _on_scan_change），逐字符跑一遍读取显然不行，所以防抖。
SCAN_DEBOUNCE_MS = 400
# 表格上方那排小地址的占位文案（两个模式共用 —— 所以别写「PDF」）
SRC_DIR_EMPTY = "原件来源：（还没挑任何文件）"
# 空清单时贴在清单框**里面**的占位话（i5 第 8 节：空态别留一片白）。
# 2026-10-23 起它也是唯一的「拖放反馈」文字位 —— 标题行删了，见 SRC_DROP_HINT。
SRC_EMPTY_HINT = "（空 · 拖入文件或点右侧「＋ 添加文件…」）"
# 拖着文件悬在窗口上、且清单还是空的时候，占位话换成这句。
# ⚠️ 清单非空时**不换**（那一片是用户正看着的文件名，别在上面盖字）。
SRC_DROP_HINT = "松手就加进来 ↓"

# ⭐ 「这一版能转哪些类型」**从 ⑩ 表 + 当前平台派生**（core.convert_kind_words）——
#    别再手写「Word / Excel」：加一种格式时，手写的那七八处文案会静默漏掉一处，
#    从此对用户说假话（在转不了它的平台上照样写「支持」）。
#    现在两个平台都一样（Word / Excel 都有），所以派生的意义在「以后加格式只动一行」。
CONVERT_KINDS = C.convert_kind_words()
# 合并模式认的类型 = 转换那 18 种 **再加 PDF 本身**（PDF 不转、直接进拼接队列）。
# 同样从表 + 平台派生，别手写 —— 见 merge_core.MERGE_EXTS。
MERGE_KINDS = M.merge_kind_words()

# ── 三个模式「还没挑文件时」状态栏那句就绪提示 ────────────────────────────
# ⚠️ 这三句要放在 MODE_CAPS **之前**：能力表里按名字引用它们，
#    写在后面会 NameError（模块导入期就炸，界面根本起不来）。
STATUS_READY = "就绪：把要处理的文件挑进「原件来源」，点「预览」看排队顺序"
STATUS_READY_CONVERT = ("就绪：把要转的 " + CONVERT_KINDS +
                        " 挑进「原件来源」—— 加进来就会列在下面")
STATUS_READY_MERGE = ("就绪：把要拼的 " + MERGE_KINDS +
                      " 挑进「原件来源」，选好打印方式和输出位置")
# 同理，「一批只拉起一次应用」那句在两边要说不同的话：
# macOS 要解释清楚「授权按应用记、不是按文件」（用户实测时最困惑的就是这点）；
# Windows 上根本没有这套授权，照抄会让人去找一个不存在的东西。
CONVERT_BATCH_NOTE = (
    "**一批文件只拉起一次 " + CONVERT_KINDS + "**（"
    + ("macOS 的授权按应用记，不是按文件）；" if sys.platform == "darwin"
       else "同格式合成一批，连着转完）；")
)

# ── 两种模式 ──────────────────────────────────────────────────────────────
# ⭐ 2026-10-08 用户提出：在「首页替换」之前其实还有一步 —— 把几十份 Word / Excel
#    批量转成 PDF（拿打印、签字）。它跟首页替换是两件事，所以做成**两个模式**，
#    顶部两个大按钮切换，各自一套界面。改模式/加模式只动这张表。
class AppMode(NamedTuple):
    key: str
    label: str          # 大按钮上的字
    tip: str            # 悬浮长说明
    blurb: str          # 按钮右边那句短提示


MODE_SPECS: tuple[AppMode, ...] = (
    AppMode(
        "convert", "导出 PDF",
        "把一批 " + CONVERT_KINDS + " 批量转成 PDF —— 打印、签字之前的那一步。\n\n"
        "· 原件一个字节都不动，转出的 PDF 写在每份原件旁边；\n"
        "· 用本机装的 " + CONVERT_KINDS + " 转 —— 就是你平时打印用的那个程序。",
        "先把 " + CONVERT_KINDS + " 批量转成 PDF（打印用）",
    ),
    AppMode(
        "swap", "PDF 首页替换",
        "把签好字的那份多页扫描件，按顺序贴回每份原件的首页。\n\n"
        "需要两样：原件（PDF）＋ 一份扫好的扫描件 PDF。\n"
        "没选扫描件时，「预览」只把原件排好队给你看 —— 照着那个顺序去扫。",
        "把签好字的扫描件贴回原件首页",
    ),
    AppMode(
        "merge", "合并打印",
        "把一批文件拼成**一份** PDF，拿去打印。\n\n"
        "· 认的文件跟「导出 PDF」一样（" + MERGE_KINDS + "），PDF 直接拼、不用转；\n"
        "· 打印机不会自动翻页时选「双面」：奇数页的那份末尾补一张空白，\n"
        "  每份都从纸的正面开始，不会串到上一份的背面；\n"
        "· 首页要单独签字 / 盖章就选「首页单面」：第 1 页之后插一张空白，\n"
        "  正文从下一张纸的正面接着印；\n"
        "· 横竖混着的一批，可以用「页面方向」统一转成竖向或横向；\n"
        "· 中间产物全部丢在临时目录，拼完即删 —— 你的原件文件夹不会被弄脏。",
        "多份文件拼成一份，双面打印不串页",
    ),
)
DEFAULT_MODE = "convert"
ALL_MODES: tuple[str, ...] = tuple(m.key for m in MODE_SPECS)
SWAP_ONLY: tuple[str, ...] = ("swap",)
CONVERT_ONLY: tuple[str, ...] = ("convert",)
MERGE_ONLY: tuple[str, ...] = ("merge",)
# 「有队列要列」的两个模式（合并与导出走同一条转换链路，只是产物去向不同）
QUEUE_MODES: tuple[str, ...] = ("convert", "merge")
# 有「输出文件名后缀」的两个模式 —— 合并是一份产物一个路径，后缀对它没意义
SUFFIX_MODES: tuple[str, ...] = ("swap", "convert")


class Pane(NamedTuple):
    """界面上一整块区域属于哪个模式 —— 切模式时按这张表 pack / pack_forget。"""

    widget: tk.Widget
    modes: tuple[str, ...]
    opts: dict


# 操作行里每个按钮/标签属于哪个模式 + 它自己的 pack 参数（顺序即显示顺序）。
# ⚠️ 别把「哪个按钮在哪个模式」写进代码里 —— 写进这张表，切模式的地方只看表。
ACTION_LAYOUT: tuple[tuple[str, dict, tuple[str, ...]], ...] = (
    ("btn_scan", {"side": "left"}, SWAP_ONLY),
    ("btn_run", {"side": "left", "padx": 6}, SWAP_ONLY),
    ("btn_convert", {"side": "left"}, CONVERT_ONLY),
    ("btn_merge", {"side": "left"}, MERGE_ONLY),
    ("btn_merge_stop", {"side": "left", "padx": 6}, MERGE_ONLY),
    ("btn_recheck", {"side": "left", "padx": 6}, QUEUE_MODES),
    ("lbl_engine", {"side": "left", "padx": 8}, QUEUE_MODES),
)

# 「选项」框里两行各属于哪个模式（缩放勾选框只有替换模式用得上）
# ⚠️ 顺序即界面顺序，`_apply_mode` 会**先全部收起、再按这张表的顺序重摆** ——
#    直接 pack 是「追加到末尾」，切几次模式顺序就会跟着切换历史乱掉（4.0 之前就是这样）。
OPTION_LAYOUT: tuple[tuple[str, dict, tuple[str, ...]], ...] = (
    ("scale_row", {"fill": "x", "pady": (0, 4)}, SWAP_ONLY),
    ("suffix_row", {"fill": "x"}, SUFFIX_MODES),
    ("print_mode_row", {"fill": "x", "pady": (0, 4)}, MERGE_ONLY),
    ("orient_row", {"fill": "x", "pady": (0, 4)}, MERGE_ONLY),
    ("merge_out_row", {"fill": "x"}, MERGE_ONLY),
)

COLUMNS: tuple[tuple[str, str, int, str], ...] = (
    ("no", "#", 34, "center"),
    ("src", "原件", 220, "w"),
    ("dir", "所在文件夹", 130, "w"),
    ("pages", "页数", 46, "center"),
    ("size", "首页尺寸", 82, "center"),
    ("scan", "扫描页", 58, "center"),
    ("scansize", "扫描页尺寸", 84, "center"),
    ("rot", "方向", 70, "center"),
    ("status", "状态", 74, "center"),
    ("note", "说明", 168, "w"),
)

# 转换模式那张表：没有扫描页/方向，多了「类型」和「产物名」。
# ⚠️ 「产物页数」只有真转完了才有（转换前那还是个 Excel，无从谈起）。
CONVERT_COLUMNS: tuple[tuple[str, str, int, str], ...] = (
    ("no", "#", 34, "center"),
    ("src", "原件", 230, "w"),
    ("dir", "所在文件夹", 130, "w"),
    ("kind", "类型", 62, "center"),
    ("out", "产物名", 210, "w"),
    ("pages", "产物页数", 68, "center"),
    ("status", "状态", 80, "center"),
    ("note", "说明", 200, "w"),
)

# 合并模式那张表：跟转换像（有类型），但「产物名」换成「补白」——
# 合并只有一份产物，每行关心的是「这一份补没补白、最后占几页」。
MERGE_COLUMNS: tuple[tuple[str, str, int, str], ...] = (
    ("no", "#", 34, "center"),
    ("src", "原件", 230, "w"),
    ("dir", "所在文件夹", 130, "w"),
    ("kind", "类型", 62, "center"),
    ("pages", "页数", 58, "center"),
    ("pad", "补白", 66, "center"),
    ("status", "状态", 80, "center"),
    ("note", "说明", 200, "w"),
)

# 一套 Treeview 两种列定义 —— 切模式时换这一份，别建两棵树
COLUMNS_BY_MODE: dict[str, tuple[tuple[str, str, int, str], ...]] = {
    "swap": COLUMNS,
    "convert": CONVERT_COLUMNS,
    "merge": MERGE_COLUMNS,
}

# 「哪个状态上什么色」只此一张表 —— 三个模式的状态词互不重叠，查它一张就够。
#  ⚠️ 以前是两张（替换一张、转换一张），同一套颜色写了两遍，加状态必漏一张。
STATUS_COLORS: dict[str, str] = {
    "替换失败": DANGER, "替换成功": OK_GREEN, "已跳过": TEXT2,
    "转换失败": DANGER, "转换成功": OK_GREEN,
    # 合并模式：本来就是 PDF、不用转的那几行 —— 说「转换成功」是对用户说假话
    M.STATUS_READY: OK_GREEN,
}


# ══════════════════════════════════════════════════════════════════════════
# 模式能力表（6.0 新增）—— 把「这个模式认什么、用哪张表、有没有后缀」收成一张
# ══════════════════════════════════════════════════════════════════════════
# ⚠️ 为什么要这张表：5.0 之前只有两个模式，界面上到处是 `mode == "convert"`
#    这样的二元判断（共 26 处）。加第三个模式时，每一处 False 分支都会**静默**
#    走进「替换」那条路 —— 表现是右键菜单不对、表格填错、按钮亮错，而且不报错。
#    收成表之后：加模式只加一行，判断全改成查表，漏不掉。
class ModeCaps(NamedTuple):
    """一个模式在界面上是什么脾气。改模式只动 MODE_CAPS。"""

    key: str
    exts: tuple[str, ...]      # 认哪些扩展名（拖文件夹展开 / 选择框都用它）
    queue: str                 # 主要数据挂在哪个属性上：plan / conv_plan / merge_plan
    suffix: bool               # 有没有「输出文件名后缀」这一行
    scan: bool                 # 要不要扫描件（只有替换要）
    adjust: bool               # 有没有 ①②③ 那三行调整工具栏
    run_btn: str               # 主按钮属性名（决定按钮文字与可用性判定）
    menu: str                  # 右键菜单属性名
    kind_words: str            # 「认哪些文件」的说法（提示 / 状态栏用）
    dir_tail: str              # 表格上方那排小地址的尾巴：产物会写在哪
    ready: str                 # 还没挑文件时状态栏那句「就绪…」


MODE_CAPS: dict[str, ModeCaps] = {
    "swap": ModeCaps(
        key="swap", exts=tuple(sorted(C.PDF_EXTS)), queue="plan",
        suffix=True, scan=True, adjust=True, run_btn="btn_run", menu="menu",
        kind_words="PDF", dir_tail="产物写在各自原件旁边",
        ready=STATUS_READY,
    ),
    "convert": ModeCaps(
        key="convert", exts=tuple(C.CONVERT_EXTS), queue="conv_plan",
        suffix=True, scan=False, adjust=False, run_btn="btn_convert",
        menu="menu_conv", kind_words=CONVERT_KINDS,
        dir_tail="转出的 PDF 写在各自原件旁边",
        ready=STATUS_READY_CONVERT,
    ),
    "merge": ModeCaps(
        key="merge", exts=tuple(sorted(M.MERGE_EXTS)), queue="merge_plan",
        suffix=False, scan=False, adjust=False, run_btn="btn_merge",
        menu="menu_merge", kind_words=MERGE_KINDS,
        dir_tail="合成一份 PDF，写到你选的位置",
        ready=STATUS_READY_MERGE,
    ),
}


def caps_of(mode: str) -> ModeCaps:
    """取当前模式的能力 —— 键写错时立刻炸，别让它静默走错分支。"""
    return MODE_CAPS[mode]

# 工具栏上三个模块的按钮清单见下面的 MODULE_ROWS（那里是唯一真相源）


class FolderRow(NamedTuple):
    """文件夹区每一行的定义 —— 改文案只动这张表。

    hint：右边那句灰字（要短，一眼扫完）
    detail：鼠标停在灰字上才弹出来的长说明（要全）
    """

    key: str
    label: str
    hint: str
    detail: str = ""
    kind: str = "dir"          # dir = 选文件夹；file = 选文件


class RowWidgets(NamedTuple):
    frame: ttk.Frame
    label: ttk.Label
    entry: ttk.Entry
    button: ttk.Button
    hint: ttk.Label


# 「原件来源」不是一行输入框，而是一张**可增删的清单**（见 _build_sources）：
# 每一行就是一个要处理的 PDF 文件（原件不用改名、也不用先归拢到同一个夹里）。
# ⚠️ 2026-10-08 用户要求**删掉「添加文件夹」这个入口** —— 一个文件夹里往往既躺着
#    要处理的 PDF、又有一堆共表 / DWG / 图片，整夹收进来是误收；「一次多选」更准。
#    core 那边仍然认「文件夹」这种来源（命令行与回归脚本还在用），只是界面不再给入口。
# ⭐ 2026-10-08 用户提出：两个模式的**来源清单也别共享**（一边挑 PDF、一边挑 Word / Excel，
#    根本不是同一批东西）—— 当时是靠「框标题各写各的」来提醒这件事。
# ⭐ 2026-10-23 用户把那一行标题**整条删掉了**。原话：「导出 PDF 底下那一行字，什么
#    『待转文件 Word Excel（每次现挑，不记住）』，这就不需要」—— 用的人已经熟了，
#    那行只占高度，还把「不记住」当卖点挂在脸上。
#    所以 SRC_BOX_TITLE_BY_MODE / SCAN_BOX_TITLE 和 App._set_src_title 一起撤掉。
#    现在「这一格收的是哪种文件」靠**清单里的文件名 + 下面的灰字 + 顶部模式条**表达。
#    ⚠️ 别再把标题加回来 —— 要加先问用户。
# 框下那行灰字也按模式换：「按文件名排队」只对替换模式有意义（转换那边顺序无关）
# ⭐ 2026-11-06 用户要求删掉灰字长句（「一长串的那种可以删掉一些」）—— 每个模式
#    只留一句最要紧的；长说明本来就收在按钮的悬浮提示里，不占界面。
#    回归盯着两处字眼别丢：替换/合并要有「按文件名」，转换要有「落在各自原件旁边」。
_SRC_HINT_LOCATE = "双击某一行可定位。"
SRC_HINT_BY_MODE: dict[str, str] = {
    "swap": "排队顺序 = 按文件名排；" + _SRC_HINT_LOCATE,
    "convert": "转出的 PDF 落在各自原件旁边；" + _SRC_HINT_LOCATE,
    "merge": "拼接顺序 = 按文件名排；" + _SRC_HINT_LOCATE,
}
SRC_TIP_FILES = (
    "挑要处理的 PDF，**支持一次选多个**（按住 Cmd / Ctrl 点，或拖框选）。\n"
    "散在好几个文件夹里也没关系 —— 分几次加，重复的会自动去重。\n\n"
    "「移除选中 / 清空」只动这张清单，**不会删磁盘上的任何文件**。"
)
SRC_TIP_REMOVE = "把选中的来源从清单里去掉（不会删文件本身，只是这次不处理它）。"
SRC_TIP_CLEAR = "清空整张清单（不会删文件本身）。"
# 转换模式下「＋ 添加文件…」的悬浮说明换一份 —— 两个模式认的文件类型不一样，
# 沿用 PDF 那份会让人以为要挑 PDF。
SRC_TIP_FILES_CONVERT = (
    "挑要转成 PDF 的 " + CONVERT_KINDS + "，**支持一次选多个**（按住 Cmd / Ctrl 点，或拖框选）。\n"
    "认这些扩展名：" + "、".join(C.CONVERT_EXTS) + "\n\n"
    "加进来只会列在下面的表格里，**不会立刻开始转** —— 确认「产物名」再点「开始转换」。\n"
    "「移除选中 / 清空」只动这张清单，**不会删磁盘上的任何文件**。"
)
# 合并模式也换一份：它收的是「可转格式 + PDF」，而且加进来**不会**在原件旁边留产物。
SRC_TIP_FILES_MERGE = (
    "挑要拼进合集的文件（" + MERGE_KINDS + "），**支持一次选多个**"
    "（按住 Cmd / Ctrl 点，或拖框选）。\n"
    "PDF 直接拼、不用转；其余格式先转成 PDF 再拼。\n\n"
    "加进来只会列在下面的表格里，**不会立刻开始合并**；\n"
    "中间 PDF 全部走临时目录，**原件文件夹不会多出任何文件**。\n"
    "「移除选中 / 清空」只动这张清单，**不会删磁盘上的任何文件**。"
)
# 悬浮说明按模式换的表：{方法名: {模式: 说明文字}}
SRC_TIPS_BY_MODE: dict[str, dict[str, str]] = {
    "on_add_files": {"swap": SRC_TIP_FILES, "convert": SRC_TIP_FILES_CONVERT,
                     "merge": SRC_TIP_FILES_MERGE},
}

# ── 「拖文件进来」（2026-10-09 用户要的：一次挑十几二十份太累）──────────────
# 用户的原话是把软件摆屏幕右边、文件夹摆左边，左边选中什么就直接拖进去。
# ⚠️ 这几句**只在真的挂上了拖放**时才显示（见 App._src_hint / _refresh_source_tips）
#    —— 挂不上还说「可以拖进来」，比不说更糟：用户会以为是自己拖的姿势不对。
DND_HINT = "；也可以把文件（或文件夹）直接拖进窗口。"
# 原来这里还有 DND_BOX_TITLE（拖着文件时把框标题换成「把文件松在这里 ↓」）——
# 标题行 2026-10-23 已整条删掉，改由整块换色 + 空清单里的 SRC_DROP_HINT 招呼。
DND_TIP_LINE = ("也可以直接把文件（或整个文件夹）从访达 / 资源管理器**拖进这个窗口** —— "
                "拖到窗口哪一处都算数，松手就加进清单。")


class SrcButton(NamedTuple):
    """来源清单右边那几个按钮 —— 改按钮只动这张表（文字 / 方法 / 悬浮说明）。"""

    text: str
    method: str          # App 上的方法名（表驱动，避免在这里写一长串 lambda）
    tip: str


SRC_BUTTONS: tuple[SrcButton, ...] = (
    SrcButton("＋ 添加文件…", "on_add_files", SRC_TIP_FILES),
    SrcButton("－ 移除选中", "on_remove_sources", SRC_TIP_REMOVE),
    SrcButton("清空", "on_clear_sources", SRC_TIP_CLEAR),
)

# 扫描件那一行（还是单值 —— 一次扫完的就是一份多页 PDF）
FOLDER_SPECS: tuple[FolderRow, ...] = (
    FolderRow("scan", "扫描件 PDF", "一次扫完的那份多页 PDF",
              "一口气扫出来的那一份 PDF，第 N 页给表格第 N 行。\n\n"
              "可以先不选 —— 没选时「预览」只把原件排好队；选上之后配对会**自动接上**。\n\n"
              "份数必须和原件数量一致，不一致会整批不执行。",
              "file"),
)

# 选项表：顺序即界面顺序（「页序倒序」已搬到工具栏做成按钮，不在这里）
# ⚠️ 「覆盖同名文件」**不再是勾选框** —— 改成运行时弹窗确认（见 _ask_overwrite）。
#    用户原话：「打勾的话又太麻烦了……只有在这个界面点击确定了之后才会直接覆盖，
#    否则就不覆盖」。弹窗**只报重名的份数**，不列文件名（一批可能十几二十份，列不完）。
# ⚠️ 「连同子文件夹一起处理」2026-10-08 已删：它原本是 v1 防「输出件落进子文件夹后
#    第二轮按名字匹配抓重」的；v3.0 输出在原位、匹配靠顺序，用途作废。
#    现在要收子文件夹里的原件，就把它**作为来源加进清单**（显式、可控，不会误收）。
CHECK_SPECS: tuple[tuple[str, str, str], ...] = (
    (
        "scale", "扫描页缩放到原件首页尺寸",
        "默认开。扫描页等比缩放、居中放进原件首页。\n\n"
        "例外：两边版面比例差太多时不硬塞 —— 原样保留，并在「说明」列写明原因。",
    ),
)

# ⭐ 2026-10-08 用户要求：两个模式的「输出文件名后缀」**各记各的，不共享**。
#    原话：「转换那边的后缀和 PDF 拼接那边的后缀不要共享记忆 —— 转换那边大概率
#    不用加后缀，PDF 拼接这边又必须加，这两个记忆不能共享。」
#    表里一行 = 一个模式用哪个变量；界面上那一行输入框是**共用的**，
#    切模式时把它的 textvariable 换成对应那个（见 _apply_mode）。
SUFFIX_VAR_BY_MODE: dict[str, str] = {"swap": "suffix", "convert": "conv_suffix"}
SUFFIX_TIPS_BY_MODE: dict[str, str] = {
    "swap": (
        "产物名 = 原件名 + 后缀。**必填** —— 产物也是 .pdf，不填会跟原件同名、\n"
        "把原件盖掉，所以工具直接拦住不让跑。\n\n"
        "WI-1032.pdf + _已签字首页  →  WI-1032_已签字首页.pdf\n\n"
        "只改名字不改内容；建议以 _ 或 - 开头，最长 40 字。\n"
        "填过会被记住 —— 这是「首页替换」自己记的，跟「Office 转 PDF」互不影响。"
    ),
    "convert": (
        "产物名 = 原件名 + 后缀 + .pdf。**可以不填** —— 原件是 .xlsx / .docx，\n"
        "扩展名不同，不填也不会撞名。\n\n"
        "WI-1032.xlsx + _已转  →  WI-1032_已转.pdf\n"
        "WI-1032.xlsx + 不填   →  WI-1032.pdf\n\n"
        "只改名字不改内容；建议以 _ 或 - 开头，最长 40 字。\n"
        "填过会被记住 —— 这是「Office 转 PDF」自己记的，跟「首页替换」互不影响。"
    ),
}
SUFFIX_NEED_TIPS_BY_MODE: dict[str, str] = {
    "swap": "必填：产物也是 .pdf，后缀为空会跟原件同名、把原件盖掉。",
    "convert": "可不填：产物是 .pdf、原件是 .xlsx / .docx，扩展名不同，不会撞。",
}
SUFFIX_SAMPLE_FALLBACK = "WI-1032"

# ── 转换模式那一套文案 ────────────────────────────────────────────────────
CONVERT_BTN_TEXT = "开始转换"
CONVERT_BTN_TIP = (
    "把「原件来源」里的 " + CONVERT_KINDS + " 转成 PDF。\n\n"
    "· 产物写在**每份原件旁边**，原件一个字节都不动；\n"
    "· " + CONVERT_BATCH_NOTE + "\n"
    "· " + C.CONVERT_RULE_NOTE + "　" + C.page_rule_label() + "\n"
    "· 转完看表格「产物页数」—— 读得出来就说明真生成了。"
)
ENGINE_RECHECK_TIP = ("重新检测本机有没有可用的 " + CONVERT_KINDS + " —— "
                      "刚装完 Microsoft Office 时点一下（程序开着的时候装的，它不会自己发现）。")

# ── 合并打印那一套文案 ────────────────────────────────────────────────────
MERGE_BTN_TEXT = "开始合并"
MERGE_BTN_TIP = (
    "把「原件来源」里的 " + MERGE_KINDS + " 合成**一份** PDF。\n\n"
    "· 要转的先转（跟「导出 PDF」同一条链路），PDF 直接拼；\n"
    "· 双面：奇数页的那一份末尾补一张空白 —— 每份都从纸的正面开始；\n"
    "· 首页单面：第 1 页之后插一张空白，首页背面留空，正文接着双面印；\n"
    "· 页面方向：横竖混批时统一转成竖向 / 横向，或保持原样；\n"
    "· 中间 PDF 全在临时目录里，拼完就删，**你的原件文件夹不会多出任何文件**；\n"
    "· 任何一份转不出来就整批不生成 —— 少一份的合集，打印时才发现最糟。"
)
# 输出默认落在桌面 —— 跟 i4 一个习惯，用户拿到就能直接双击打印。
MERGE_DEFAULT_OUT = str(Path.home() / "Desktop" / M.DEFAULT_OUT_NAME)
MERGE_OUT_TIP = (
    "合成的那一份写到这里。\n\n"
    "· 同名文件会被**覆盖**（跑之前会问你一次）；\n"
    "· 这个路径**不记忆** —— 每次打开都回到桌面上的《" + M.DEFAULT_OUT_NAME + "》，\n"
    "  免得上一轮的旧路径早被删了、还指着它。"
)

PREVIEW_TIP = (
    "一个按钮管两段：\n\n"
    "① **还没选扫描件** → 只把原件按文件名排好队。照着这个顺序去扫，第 1 行对第 1 页。\n\n"
    "② **选好扫描件之后** → 自动配上「第 N 行用第 N 页」；份数对不上会整批不执行。\n\n"
    "换了来源要重新点一次「预览」；只换扫描件会自动重配，不用点。"
)

# ── 三个功能模块的说明文案（都挂在按钮上，鼠标停一下就出来）──────────────
ORDER_TIP = (
    "扫描仪倒着出纸时点一下 —— 整批在 正序 ⇄ 倒序 之间切。\n\n"
    "⚠️ 这是整批重算，会把手动调过的页码清回默认 —— 先定页序，再调个别行。"
)
PAGE_TIP_UP = "把这一行用的扫描页，和**上一行**对调 —— 行的顺序不动，只换页码。"
PAGE_TIP_DOWN = "把这一行用的扫描页，和**下一行**对调 —— 行的顺序不动，只换页码。"

ROT_TIP_AUTO = (
    "扫描件默认按**竖着扫**算：原件首页是横向的那一页，自动左转 90°。\n\n"
    "全部左 / 右转 ＝ 整批一起转；横向左 / 右转 ＝ 只动横向的那些页；\n"
    "方向复位 ＝ 回到默认规则。\n\n"
    "⚠️ 手动定的方向记在**扫描页**上 —— 换一份扫描件会清空。"
)


# ── 工具栏「①②③」三个模块的表 ────────────────────────────────────────
# 一行 = 一个模块，界面上的顺序就是这张表的顺序。
# 建按钮、挂悬浮提示、控「什么时候能点」全都从这张表派生 ——
# 以后加/减按钮只改这里，不会漏掉某一处（这是 v1 踩过的坑）。
class BtnSpec(NamedTuple):
    """一个按钮的定义。method + args 由 App._mod_command 拼成回调。"""

    attr: str                                  # 挂到 App 上的属性名（回归脚本按它取按钮）
    text: str                                  # 按钮文字；留空 = 建完由代码动态填
    width: int
    method: str                                # App 上的方法名
    args: tuple = ()
    tip: str = ""


class ModuleRow(NamedTuple):
    tag: str                                   # 行首编号（① / ② / ③），也是这一行的名字
    buttons: tuple[BtnSpec, ...]
    hint: str = ""                             # 行尾灰字（要短，长说明收进按钮的悬浮提示）


MODULE_ROWS: tuple[ModuleRow, ...] = (
    ModuleRow("① 扫描件页序", (
        BtnSpec("btn_order", "", 12, "on_flip_order", tip=ORDER_TIP),
    ), "扫描仪倒着出纸时点一下"),
    ModuleRow("② 页码对应", (
        BtnSpec("btn_up", "↑ 与上行对调", 12, "_swap_page", (-1,), PAGE_TIP_UP),
        BtnSpec("btn_down", "↓ 与下行对调", 12, "_swap_page", (1,), PAGE_TIP_DOWN),
    ), "双击某一行也能直接指定页号"),
    ModuleRow("③ 方向（整批转）", (
        BtnSpec("btn_rot_all_l", "全部左转 90°", 13, "on_turn_all", (C.ROT_LEFT,), ROT_TIP_AUTO),
        BtnSpec("btn_rot_all_r", "全部右转 90°", 13, "on_turn_all", (C.ROT_RIGHT,), ROT_TIP_AUTO),
        BtnSpec("btn_rot_land_l", "横向左转 90°", 13, "on_turn_land", (C.ROT_LEFT,), ROT_TIP_AUTO),
        BtnSpec("btn_rot_land_r", "横向右转 90°", 13, "on_turn_land", (C.ROT_RIGHT,), ROT_TIP_AUTO),
        BtnSpec("btn_rot_reset", "方向复位", 10, "on_reset_rot", tip=ROT_TIP_AUTO),
    )),
)

# 三个模块的全部按钮（由上面的表派生）—— _set_busy / _sync_adjust_buttons 只看这一份
ADJUST_BUTTONS = tuple(b.attr for row in MODULE_ROWS for b in row.buttons)


def default_geometry(screen_w: int, screen_h: int) -> str:
    """初始窗口大小 —— 要放得下四条工具栏 + 表格，但绝不能超出屏幕。

    交付到 Windows 笔记本上时屏幕可能只有 1366×768，写死一个高尺寸会被任务栏切掉，
    所以这里按屏幕实际大小收一收。
    """
    w = min(GEOM_DEFAULT[0], max(screen_w - 40, GEOM_MIN[0]))
    h = min(GEOM_DEFAULT[1], max(screen_h - 90, GEOM_MIN[1]))
    return f"{w}x{h}"


def fit_geometry(geo: object, screen_w: int, screen_h: int) -> str | None:
    """窗口大小只在合理范围内才恢复，返回 None 表示「这次不恢复」。

    抽成独立函数是因为它不好从界面上测（withdraw 状态下 winfo_width() 恒为 1）。
    恢复出一个看不见的窗口比不恢复更糟，所以宁可保守。
    """
    if not isinstance(geo, (list, tuple)) or len(geo) != 2:
        return None
    try:
        w, h = int(geo[0]), int(geo[1])
    except Exception:  # noqa: BLE001
        return None
    if GEOM_MIN[0] <= w <= screen_w and GEOM_MIN[1] <= h <= screen_h:
        return f"{w}x{h}"
    return None


class Tip:
    """极简悬浮提示：鼠标停留片刻弹出，移开或点击即消失。零第三方依赖。"""

    def __init__(self, widget: tk.Widget, text: str, delay: int = 350, width: int = 520) -> None:
        self.widget = widget
        self.text = text
        self.delay = delay
        self.width = width
        self._after: str | None = None
        self._win: tk.Toplevel | None = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e: tk.Event | None = None) -> None:
        self._cancel()
        self._after = self.widget.after(self.delay, self._show)

    def _cancel(self) -> None:
        if self._after:
            try:
                self.widget.after_cancel(self._after)
            except Exception:  # noqa: BLE001
                pass
            self._after = None

    def _show(self) -> None:
        if self._win or not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._win = tk.Toplevel(self.widget)
        self._win.wm_overrideredirect(True)
        self._win.wm_geometry(f"+{x}+{y}")
        # 统一浅色系：白底 + 1px 描边（原来那身「系统黄」跟界面不是一套）
        tk.Label(
            self._win, text=_plain(self.text), justify="left", anchor="w",
            background=CARD, foreground=TEXT, font=(FONT, 10),
            highlightbackground=BORDER, highlightthickness=1,
            wraplength=self.width, padx=10, pady=8,
        ).pack()

    def _hide(self, _e: tk.Event | None = None) -> None:
        self._cancel()
        if self._win is not None:
            try:
                self._win.destroy()
            except Exception:  # noqa: BLE001
                pass
            self._win = None


def _fmt_size(sz: tuple[float, float]) -> str:
    if not sz or min(sz) <= 0:
        return "—"
    return f"{sz[0]:.0f}×{sz[1]:.0f}"


# ── 「拖文件进来」的两块零件 ────────────────────────────────────────────
#
# ⭐ 为什么拆成两个函数、而不是塞进 App 里：这两段都不需要窗口已经建好 ——
#    `enable_dnd` 只要一个主窗口，`drop_paths` 是纯字符串处理。摆在模块级，
#    回归脚本不建界面也能直接测（沙箱里 Tcl 起不来，建不了窗口）。
def enable_dnd(root: tk.Tk) -> bool:
    """给**已经建好的**主窗口挂上「接拖入」的能力。挂不上返回 False。

    ⭐ 用 `tkinterdnd2.require(root)` 而不是 `TkinterDnD.Tk()`：
       后者的 `__init__` 一旦 tkdnd 加载失败就**把窗口初始化整个炸掉**，
       工具直接打不开；`require()` 是「给已有窗口补一个能力」，
       失败只是没这个能力，窗口和按钮照常。
    ⚠️ require 只该调一次；重复调会重复 lappend auto_path（不致命，但没必要）。
    """
    global _DND_READY
    if _dnd_require is None:
        return False
    try:
        _dnd_require(root)
        _DND_READY = True
    except Exception:  # noqa: BLE001 —— tkdnd 二进制不被支持 / 加载失败
        _DND_READY = False
    return _DND_READY


def drop_paths(data: str, splitlist=None) -> list[str]:
    """`<<Drop>>` 事件的数据 → 一串路径。

    ⚠️ 不能拿 `.split()` 硬切：tkdnd 丢过来的是**一个 Tcl list**，
       路径里有空格 / 花括号时那一项会被花括号整个包起来
       （`{/Users/user/我的 文档/a.pdf}`）。硬切的结果是把一条路径劈成
       「{/Users/user/我的」+「文档/a.pdf}」两条根本不存在的路径 ——
       而这种事**只在带空格的文件夹上出现**，最容易漏测。

    `splitlist` 传窗口的 `tk.splitlist`（Tcl 官方解析，最准）。
    ⚠️ 做成参数是为了可测：沙箱里 Tcl 起不来、建不了窗口，但传一个假的进来
       照样能验证「滤掉空项、剥掉两端空白、解析器失效时的兜底」。
    """
    text = data or ""
    if not text.strip():
        return []
    try:
        items = list(splitlist(text)) if callable(splitlist) else text.split()
    except Exception:  # noqa: BLE001 —— 解析器不认就退回按空白切
        items = text.split()
    return [s for s in (str(p).strip() for p in items) if s]


def _plain(text: str) -> str:
    """去掉说明文字里的 `**` —— Tk 的 Label 不认 Markdown，会原样画出来。

    这些说明本来就是给悬浮提示写的（那一层也不渲染 Markdown），
    统一在显示前洗掉，省得界面上到处是星号。
    """
    return (text or "").replace("**", "")


def _mk_check(parent, text, var, command=None, bg=None):
    """统一的勾选框：**画的是对勾，不是叉**。

    `ttk.Checkbutton` 在 clam 主题下选中态画出来是一个「✗」，和「关闭 / 不通过」
    撞语义（2026-10-21 用户反馈）。改用经典 `tk.Checkbutton`：macOS 由系统画
    原生蓝底白勾，跟 5.0 那两处勾选外观一致。
    """
    return tk.Checkbutton(parent, text=text, variable=var, command=command,
                          bg=bg or BG, fg=TEXT, activebackground=bg or BG,
                          activeforeground=TEXT, selectcolor=CARD,
                          font=(FONT, 10), bd=0, highlightthickness=0,
                          anchor="w")


def _pick_font() -> str:
    """本机能用的界面字体：Mac＝PingFang SC，Windows＝微软雅黑。"""
    want = (("PingFang SC", "Heiti SC")
            if sys.platform == "darwin"
            else ("Microsoft YaHei UI", "Microsoft YaHei", "SimHei"))
    try:
        have = set(tkfont.families())
    except Exception:                        # noqa: BLE001 —— 问不出来就用首选
        return want[0]
    for name in want:
        if name in have:
            return name
    return want[-1]


def apply_theme(root: tk.Misc) -> None:
    """全应用只此一处定义样式（4.0 起统一成 5.0 那套浅色系）。

    ⚠️ 顺序有讲究：先 `theme_use("clam")` 再 configure —— 反过来会被换主题清掉。
    ⚠️ 按钮只有三档（主 / 强调字 / 次），别再加第四种形状（i5 第 9 节）。
    """
    global FONT
    # 字体：Mac 用 PingFang SC，Windows 用微软雅黑（缺了就退回系统默认）。
    # ⚠️ 必须在有 Tk 根窗口之后才能问「本机有哪些字体」—— 所以放在这里，
    #    不放在模块顶层（顶层那时还没有 Tk，问了会抛异常）。
    FONT = _pick_font()

    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:                     # 换不了就用默认主题，别让窗口起不来
        pass
    st.configure(".", background=BG, foreground=TEXT, font=(FONT, 11))
    st.configure("TFrame", background=BG)
    st.configure("TLabel", background=BG, foreground=TEXT, font=(FONT, 11))
    st.configure("TEntry", fieldbackground=CARD, background=CARD, foreground=TEXT,
                 bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                 insertcolor=TEXT, padding=(6, 3))
    st.map("TEntry", bordercolor=[("focus", PRIMARY)],
           lightcolor=[("focus", PRIMARY)])
    st.configure("TButton", font=(FONT, 11), padding=(8, 4),
                 background=CARD, foreground=TEXT, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD)
    st.map("TButton", background=[("active", HOVER), ("pressed", "#e6e9ee")],
           foreground=[("disabled", DISABLED_FG)])
    st.configure("Small.TButton", font=(FONT, 10), padding=(6, 2),
                 background=CARD, foreground=TEXT, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD)
    st.map("Small.TButton", background=[("active", HOVER), ("pressed", "#e6e9ee")],
           foreground=[("disabled", DISABLED_FG)])
    # 主按钮：一屏只有一颗（预览 / 开始转换）
    st.configure("Primary.TButton", font=(FONT, 11, "bold"), padding=(14, 5),
                 background=PRIMARY, foreground="#ffffff",
                 bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY)
    st.map("Primary.TButton",
           background=[("active", PRIMARY_D), ("pressed", PRIMARY_D),
                       ("disabled", "#c9d5ee")],
           foreground=[("disabled", "#f2f6ff")])
    # 强调字按钮：白底 + 深蓝粗字（「开始替换」—— 下一步该点它，但不跟主按钮抢）
    st.configure("Accent.TButton", font=(FONT, 11, "bold"), padding=(12, 4),
                 background=CARD, foreground=PRIMARY_D, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD)
    st.map("Accent.TButton",
           background=[("active", "#eef4ff"), ("pressed", "#dfe9ff"),
                       ("disabled", CARD)],
           foreground=[("disabled", DISABLED_FG)])
    # 模式切换：选中＝白底深蓝粗字，未选中＝窗口底灰字（都不铺主色）
    st.configure("ModeOn.TButton", font=(FONT, 11, "bold"), padding=(12, 5),
                 background=HEAD_BG, foreground=PRIMARY_D, bordercolor=SPINE,
                 lightcolor=HEAD_BG, darkcolor=HEAD_BG)
    st.map("ModeOn.TButton", background=[("active", HEAD_BG)])
    st.configure("ModeOff.TButton", font=(FONT, 11), padding=(12, 5),
                 background=BG, foreground=TEXT2, bordercolor=BORDER,
                 lightcolor=BG, darkcolor=BG)
    st.map("ModeOff.TButton", background=[("active", HOVER)])
    st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                 bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                 rowheight=26, font=(FONT, 10))
    st.configure("Treeview.Heading", background=HEAD_BG, foreground=HEAD_FG,
                 font=(FONT, 10, "bold"), relief="flat", padding=(6, 5),
                 bordercolor=BORDER, lightcolor=HEAD_BG, darkcolor=HEAD_BG)
    st.map("Treeview.Heading", background=[("active", "#e2ecff")])
    st.map("Treeview", background=[("selected", DZ_BG_ON)],
           foreground=[("selected", TEXT)])
    st.configure("Cover.Horizontal.TProgressbar", troughcolor="#dde1e7",
                 background=BAR, bordercolor="#dde1e7",
                 lightcolor=BAR, darkcolor=BAR, thickness=8)
    st.configure("Vertical.TScrollbar", background=CARD, troughcolor=BG,
                 bordercolor=BORDER, arrowcolor=TEXT2, lightcolor=CARD,
                 darkcolor=CARD)
    st.configure("Horizontal.TScrollbar", background=CARD, troughcolor=BG,
                 bordercolor=BORDER, arrowcolor=TEXT2, lightcolor=CARD,
                 darkcolor=CARD)


class App(ttk.Frame):
    # ── 来源清单（两个模式各记各的，见 __init__ 里的 _sources_by_mode）────────
    #
    # 做成属性而不是「到处手动写 self._sources_by_mode[self.mode]」：
    # 清单被读了十几处（列队、画地址行、判空、回归脚本…），散着写迟早漏一处 ——
    # 漏掉的表现就是「在这个模式看见了另一个模式的文件」。属性收口之后，
    # `self.sources.append(...)` / `del self.sources[i]` / `self.sources = [...]`
    # 这些老写法一个字都不用改，落点自动是当前模式那一份。
    @property
    def sources(self) -> list[str]:
        """当前模式的来源清单 —— 「首页替换」和「Office 转 PDF」各存各的。"""
        return self._sources_by_mode[self.mode]

    @sources.setter
    def sources(self, value: list[str]) -> None:
        self._sources_by_mode[self.mode] = list(value)

    def __init__(self, master: tk.Tk) -> None:
        # ⚠️ 外层 padding 必须是 0：底部状态条要通栏收口，
        #    外层一留白它就跟着缩进去，看着又变成一张卡片（见 _build）。
        super().__init__(master, padding=0)
        self.pack(fill="both", expand=True)
        # ⭐ 当前模式（顶部那两个大按钮）。**不写进配置** —— 每次打开都落在
        #    「PDF 首页替换」上，那是主流程；转换只是它前面的一步。
        self.mode: str = DEFAULT_MODE
        self.plan: C.Plan | None = None
        self.conv_plan: C.ConvertPlan | None = None
        # ⭐ 6.0 合并打印：自己的队列（不复用 conv_plan —— 那边算的是「每份产物叫什么」，
        #    这边算的是「拼成一份、每份占几页、补没补白」，两份数据长得很不一样）。
        self.merge_plan: M.MergePlan | None = None
        self.merge_result: M.MergeResult | None = None
        self.merge_stop = threading.Event()          # 「停止」按钮 → 工作线程
        self.merge_progress: tuple[int, int, str] = (0, 0, "")
        self._merge_pump: str | None = None          # 进度泵的 after 句柄
        self._poll_armed = False                     # 轮询链是否已经排上（见 _poll）
        # ⭐ 2026-10-09：不再是「一个引擎」—— 一族一条（见 core 的 ⑩a）。
        #    None ＝ 还没探过（探针要摸文件系统 / 试 import，探一次就缓存着）；
        #    「重新检测引擎」按钮会把它清掉重探。
        self.engines: list[C.ConvEngine] | None = None
        # 转完之后：原件 → 那份结果（产物名 / 页数 / 说明）。表格「产物名 / 产物页数」
        # 两列就是从这里取的 —— 转换前那还是个 Excel，页数无从谈起，只能显示预测名。
        self.conv_results: dict[Path, C.ConvertResult] = {}
        self.busy = False
        # ⭐ 「拖文件进来」此刻还没挂上（要等界面建出来才知道成不成），先按
        #    「没这功能」起手 —— 建界面的路上（_apply_mode）就要读它决定
        #    那句灰字怎么说，晚一步赋值会 AttributeError。
        self.dnd_ok = False
        self.rows: dict[str, RowWidgets] = {}
        self._save_after: str | None = None
        self._scan_after: str | None = None         # 扫描件那一格的防抖句柄
        self._preview_auto = False                  # 这次预览是「选完扫描件自动补的」吗
        self._jobs: list[tuple[str, object]] = []   # 后台线程 → 主线程的交接处
        self._config_broken = False
        # ⑥ 人工判的方向：{扫描页序号: 角度}（只记手动，自动规则由 core 按原件横竖算）
        self.rotations: dict[int, int] = {}
        self._rot_scan = ""
        # ⭐ 来源清单**按模式各记各的**（2026-10-08 用户要求）。
        #    「首页替换」挑的是 PDF、「Office 转 PDF」挑的是 Word / Excel ——
        #    根本不是同一批东西；共用一个清单的话，切过去会看见一堆这个模式用不上的文件
        #    （转换那边还会把它们全列进「认不出格式、已跳过」）。
        #    表里一行 = 一个模式存哪个键；`sources` 是属性，读写自动落到当前模式那份
        #    （见下面的 sources property）—— 所以 `self.sources.append(...)` 这类
        #    老写法一个字都不用改。
        #    ⚠️ 跟后缀那套（SUFFIX_VAR_BY_MODE）是同一个思路，但**清单不落盘**
        #       （见 CONFIG_PERSIST_VARS）：分开只为「切模式不串台」，不为「记住」。
        self._sources_by_mode: dict[str, list[str]] = {m: [] for m in ALL_MODES}
        self.vars = {
            # ⚠️ 「src」没了：原件来源现在是一张清单，不是一行输入框（见 _build_sources）。
            # ⚠️ 「out」也没了：产物写在每份原件旁边，不再有输出文件夹这一项。
            # ⚠️ 「recursive」删了：见 CHECK_SPECS 上面的注释。
            #    overwrite 也不再是勾选框里的变量 —— 它现在只是运行期的内部开关
            #    （覆盖确认弹窗点了「是」才置 True，见 on_run）。
            # ⚠️ 「scan」不再写进配置：它不在 CONFIG_PERSIST_VARS 白名单里。
            "scan": tk.StringVar(),
            "reverse": tk.BooleanVar(value=False),
            "scale": tk.BooleanVar(value=True),
            # ⚠️ 后缀是**两个**变量 —— 两个模式各记各的，别合并回一个（见 SUFFIX_VAR_BY_MODE）
            "suffix": tk.StringVar(value=""),
            "conv_suffix": tk.StringVar(value=""),
            # 合并打印：打印方式**记住**（同一台打印机每次都是同一种），
            # 输出路径**不记**（见 MERGE_OUT_TIP：上一轮的旧路径多半已经被删了）。
            "print_mode": tk.StringVar(value=M.DEFAULT_PRINT_MODE),
            # ⭐ 6.2：页面方向。默认「保持原样」—— 一开就会改变老口径的产物。
            "page_orient": tk.StringVar(value=M.DEFAULT_ORIENT),
            "merge_out": tk.StringVar(value=MERGE_DEFAULT_OUT),
        }
        self._build()
        # ⭐ 「拖文件进来」：窗口建出来之后才挂得上（见 _setup_dnd）。
        #    挂完立刻把两处文案刷一遍 —— 挂得上才提「也可以拖进来」，
        #    挂不上就一个字都不提（说了却拖不动，比不说更糟）。
        self.dnd_ok = self._setup_dnd()
        self._refresh_source_tips()
        self.lbl_src_hint.config(text=self._src_hint())
        self._load_config()
        for key in SUFFIX_VAR_BY_MODE.values():
            self.vars[key].trace_add("write", lambda *_a: self._on_suffix_change())
        # ⭐ 换了扫描件**不作废队列** —— 只把配对补上（见 _on_scan_change）
        self.vars["scan"].trace_add("write", lambda *_a: self._on_scan_change())
        # 打印方式一变，右边那句说明和「开始合并」的可用性跟着变
        self.vars["print_mode"].trace_add("write", lambda *_a: self._refresh_merge_hint())
        # 6.2：页面方向一变，右边那句也要换（同一处刷新）
        self.vars["page_orient"].trace_add("write", lambda *_a: self._refresh_orient_hint())
        # ⚠️ reverse 不挂自动重排 —— 它是按钮显式控制的（on_flip_order），
        #    挂 trace 会在读配置/复位时偷偷重排一次，用户没点却变了
        for var in self.vars.values():
            var.trace_add("write", lambda *_a: self._schedule_save())
        self.winfo_toplevel().protocol("WM_DELETE_WINDOW", self.on_close)
        self._announce_restore()
        # ⚠️ 上面这些 trace 是**在 `_load_config()` 之后**才挂的（不能反过来：
        #    挂早了会在读配置时就把配置又存一遍）。所以配置恢复出来的 suffix
        #    不会触发 _on_suffix_change —— 派生显示必须在这里手工补刷一次。
        #    漏掉这一刀的表现：上次填过的后缀明明恢复了，右边那格却还是空的，
        #    「开始替换」的可用性也没跟着算过。
        self._refresh_suffix_hint()
        self._refresh_srcdir_label()
        self._sync_run_button()

    # ── 界面搭建 ─────────────────────────────────────────────────────────
    # ── 界面搭建（统一浅色系，规范见 项目i5「浅色系UI风格指南」）────────────
    #
    # 层次只有三层：窗口底 BG → 白卡片 CARD → 块中块 DZ_BG。
    # 分界一律 1px 的 BORDER，不加阴影、不加粗框；主色只出现在一颗实心按钮上。
    def _pane(self, widget: tk.Widget, modes: tuple[str, ...], **opts) -> tk.Widget:
        """登记一整块区域属于哪个模式 —— 切模式时按 `_panes` 重新 pack。"""
        self._panes.append(Pane(widget, modes, opts))
        return widget

    def _build_header(self) -> tk.Frame:
        """标题只占一行：应用名 12 加粗 + 副标题 10 灰（不堆第二个标题）。"""
        bar = tk.Frame(self, bg=BG)
        tk.Label(bar, text=APP_TITLE, bg=BG, fg=TEXT,
                 font=(FONT, 12, "bold")).pack(side="left")
        tk.Label(bar, text=APP_SUBTITLE, bg=BG, fg=TEXT2,
                 font=(FONT, 10)).pack(side="left", padx=(8, 0))
        return bar

    def _build_mode_bar(self) -> tk.Frame:
        """顶部模式切换 —— 浅色分段，**不铺主色**。

        选中 = 白底 + 深蓝字 + 加粗；未选中 = 窗口底 + 灰字。
        ⚠️ 别把选中的那个做成实心蓝：一屏只能有一颗实心按钮
           （i5 第 9 节第 1 条，用户 2026-10-22 亲口否过「饱和蓝铺大块」）。
        """
        bar = tk.Frame(self, bg=BG)
        seg = tk.Frame(bar, bg=BG)
        seg.pack(side="left")
        self._mode_buttons: dict[str, ttk.Button] = {}
        for spec in MODE_SPECS:
            b = ttk.Button(seg, text=spec.label, style="ModeOff.TButton",
                           command=lambda k=spec.key: self.on_mode(k))
            b.pack(side="left", padx=(0, 6))
            Tip(b, spec.tip)
            self._mode_buttons[spec.key] = b
        self.lbl_mode = tk.Label(bar, text="", bg=BG, fg=TEXT2, font=(FONT, 10))
        self.lbl_mode.pack(side="left", padx=(8, 0))
        return bar

    # ── 唯一输入块（原件来源）─────────────────────────────────────────────
    def _build_sources(self) -> tk.Frame:
        """「原件来源」—— 整块极浅底 + 1px 框，**整块都是投放目标**。

        为什么是清单：原件经常散在好几个文件夹里（一个 WI 一个文件夹），
        或者只想处理某一堆里的其中几份。单值输入框逼着你先把它们归到一个夹里，
        清单则「加单个文件」和「一次多选」两种都认。

        ⚠️ 2026-10-23 起这块**没有标题行了**（用户要求，见文件上方那段说明）。
           空清单时那句占位话用 place 贴在清单框**里面**，出没都不占行高、
           也不会把下面的灰字推来推去。
        """
        dz = tk.Frame(self, bg=DZ_BG, highlightbackground=BORDER,
                      highlightthickness=1)
        self.box_sources = dz
        dz.columnconfigure(0, weight=1)

        holder = tk.Frame(dz, bg=CARD, highlightbackground=BORDER,
                          highlightthickness=1)
        holder.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(8, 0))
        self.lst_sources = tk.Listbox(
            holder, height=3, selectmode="extended", activestyle="none",
            exportselection=False, bg=CARD, fg=TEXT, font=(FONT, 10),
            selectbackground=DZ_BG_ON, selectforeground=TEXT,
            highlightthickness=0, bd=0, relief="flat")
        sb = ttk.Scrollbar(holder, orient="vertical", command=self.lst_sources.yview)
        self.lst_sources.configure(yscrollcommand=sb.set)
        self.lst_sources.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=3)
        sb.pack(side="left", fill="y")
        self.lst_sources.bind("<Double-1>", self._on_source_double)
        # 空态提示（清单为空时才显示）——见 _refresh_source_list
        # ⚠️ 它贴在清单框**里面**（place in_=清单框）：不占一行，清单有内容就收起来。
        #    父件挂 holder（别真挂到 Listbox 上）—— Listbox 不是容器，子件行为不稳。
        self.lbl_src_empty = tk.Label(holder, text=SRC_EMPTY_HINT,
                                      bg=CARD, fg=MUTED, font=(FONT, 10))
        self.lbl_src_empty.place(in_=self.lst_sources, x=4, y=3)

        btns = tk.Frame(dz, bg=DZ_BG)
        btns.grid(row=0, column=1, sticky="n", padx=(8, 10), pady=(8, 0))
        # 悬浮说明留个引用：它按模式换文字（见 SRC_TIPS_BY_MODE），
        # 建完就丢的话切模式时改不了。
        self._src_tips: dict[str, Tip] = {}
        for i, spec in enumerate(SRC_BUTTONS):
            b = ttk.Button(btns, text=spec.text, style="Small.TButton",
                           command=getattr(self, spec.method))
            b.pack(fill="x", pady=(0 if i == 0 else 5, 0))
            self._src_tips[spec.method] = Tip(b, spec.tip)

        # 这一行按模式换（见 _src_hint）；「也可以拖进来」只在真挂上拖放时才出现
        self.lbl_src_hint = tk.Label(dz, text=_plain(self._src_hint()), bg=DZ_BG,
                                     fg=TEXT2, font=(FONT, 10), anchor="w",
                                     justify="left", wraplength=880)
        self.lbl_src_hint.grid(row=1, column=0, columnspan=2, sticky="w",
                               padx=(10, 10), pady=(4, 6))

        # 拖东西进窗口时整块换色（见 _drag_hint）—— 只亮提示行会让人以为
        # 下半截不能拖，所以这几个件的底都跟着一起换。
        self._dz_tintable = (dz, self.lbl_src_hint, btns)
        return dz

    # ── 扫描件（只有「PDF 首页替换」用得上）───────────────────────────────
    def _build_scan_box(self) -> tk.Frame:
        """扫描件那一行 —— 一张白卡片，**只在替换模式出现**（转换用不到扫描件）。"""
        card = tk.Frame(self, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1)
        tk.Label(card, text="扫描件", bg=CARD, fg=TEXT2,
                 font=(FONT, 10, "bold")).pack(anchor="w", padx=10, pady=(6, 2))
        for spec in FOLDER_SPECS:
            row = tk.Frame(card, bg=CARD)
            lbl = tk.Label(row, text=spec.label, width=10, anchor="w", bg=CARD,
                           fg=TEXT, font=(FONT, 10))
            lbl.pack(side="left")
            ent = ttk.Entry(row, textvariable=self.vars[spec.key])
            ent.pack(side="left", fill="x", expand=True, padx=(0, 6))
            btn = ttk.Button(row, text="浏览…", style="Small.TButton",
                             command=lambda k=spec.key: self._pick(k))
            btn.pack(side="left")
            hint = tk.Label(row, text=spec.hint, bg=CARD, fg=TEXT2, font=(FONT, 10))
            hint.pack(side="left", padx=(10, 0))
            if spec.detail:
                Tip(hint, spec.detail)      # 长说明收进悬浮，别堆在界面上
            self.rows[spec.key] = RowWidgets(row, lbl, ent, btn, hint)
            row.pack(fill="x", padx=10, pady=(0, 6))
        return card

    # ── 选项（两个模式共用）───────────────────────────────────────────────
    def _build_options(self) -> tk.Frame:
        """「输出文件名后缀」那一行所在的白卡片。

        ⚠️ 2026-10-23 用户要求删掉卡片小标题「选项」：转换模式下这张卡里**只有后缀
           一行**，顶着「选项」两个字纯属多余（用的人也已经知道那是干嘛的）。
           现在卡片里就是一行内容，靠白底 + 1px 框跟别区分开。

        ⚠️ 后缀那一行**只有一格输入框**，但背后是两个变量（见 SUFFIX_VAR_BY_MODE）：
           切模式时 `_rebind_suffix` 会把它的 textvariable 换掉，看到 / 填的就是
           那个模式自己那一份，两边互不影响。

        ⚠️ 两行**不在这里 pack** —— 交给 `_apply_mode` 按 OPTION_LAYOUT 的顺序摆，
           否则「先 pack 一次再重 pack」会留下跟切换历史有关的乱序。
        """
        card = tk.Frame(self, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1)
        body = tk.Frame(card, bg=CARD)
        body.pack(fill="x", padx=10, pady=7)

        l1 = tk.Frame(body, bg=CARD)            # 缩放勾选框：只有替换模式用得上
        for key, label, tip in CHECK_SPECS:
            cb = _mk_check(l1, label, self.vars[key], None, bg=CARD)
            cb.pack(side="left", padx=(0, 14))
            Tip(cb, tip)

        l2 = tk.Frame(body, bg=CARD)
        tk.Label(l2, text="输出文件名后缀", bg=CARD, fg=TEXT,
                 font=(FONT, 10)).pack(side="left", padx=(0, 8))
        self.ent_suffix = ttk.Entry(l2, width=16,
                                    textvariable=self.vars[SUFFIX_VAR_BY_MODE[DEFAULT_MODE]])
        self.ent_suffix.pack(side="left", padx=(0, 8))
        self._tip_suffix = Tip(self.ent_suffix, SUFFIX_TIPS_BY_MODE[DEFAULT_MODE])
        # 「必填」不是装饰 —— 后缀为空时「开始替换」直接是禁用的（见 _sync_run_button）
        self.lbl_suffix_need = tk.Label(l2, text="必填", bg=CARD, fg=DANGER,
                                        font=(FONT, 10, "bold"))
        self.lbl_suffix_need.pack(side="left", padx=(0, 10))
        self._tip_suffix_need = Tip(self.lbl_suffix_need,
                                    SUFFIX_NEED_TIPS_BY_MODE[DEFAULT_MODE])
        self.lbl_suffix = tk.Label(l2, text="", bg=CARD, fg=TEXT2, font=(FONT, 10))
        self.lbl_suffix.pack(side="left")

        # ── 合并打印：打印方式（双面 / 单面 / 首页单面其余双面）────────────
        # ⚠️ 用原生 tk.Radiobutton，不用 ttk：ttk 的单选钮在 clam 主题下会带一圈
        #    灰底，跟这张白卡片不搭（i5 第 2.3 节，勾选框同理）。
        l3 = tk.Frame(body, bg=CARD)
        tk.Label(l3, text="打印方式", bg=CARD, fg=TEXT,
                 font=(FONT, 10)).pack(side="left", padx=(0, 8))
        for val in list(M.PrintMode):
            rb = tk.Radiobutton(
                l3, text=M.PRINT_MODE_LABELS[val.value], variable=self.vars["print_mode"],
                value=val.value, font=(FONT, 10), bg=CARD, fg=TEXT,
                activebackground=CARD, activeforeground=TEXT, selectcolor=CARD,
                bd=0, highlightthickness=0, anchor="w")
            rb.pack(side="left", padx=(0, 14))
            Tip(rb, M.print_mode_words(val.value))
        self.lbl_print_tip = tk.Label(l3, text="", bg=CARD, fg=TEXT2, font=(FONT, 10))
        self.lbl_print_tip.pack(side="left")

        # ── 合并打印：页面方向（保持原样 / 全部竖向 / 全部横向）─────────────
        # ⭐ 6.2 新增。横竖混批打印出来纸张不一致，这一行让用户自己定规矩。
        l3b = tk.Frame(body, bg=CARD)
        tk.Label(l3b, text="页面方向", bg=CARD, fg=TEXT,
                 font=(FONT, 10)).pack(side="left", padx=(0, 8))
        for val in list(M.PageOrient):
            rb = tk.Radiobutton(
                l3b, text=M.ORIENT_LABELS[val.value], variable=self.vars["page_orient"],
                value=val.value, font=(FONT, 10), bg=CARD, fg=TEXT,
                activebackground=CARD, activeforeground=TEXT, selectcolor=CARD,
                bd=0, highlightthickness=0, anchor="w")
            rb.pack(side="left", padx=(0, 14))
            Tip(rb, M.orient_words(val.value))
        self.lbl_orient_tip = tk.Label(l3b, text="", bg=CARD, fg=TEXT2, font=(FONT, 10))
        self.lbl_orient_tip.pack(side="left")

        # ── 合并打印：输出到哪一份文件 ───────────────────────────────────
        l4 = tk.Frame(body, bg=CARD)
        tk.Label(l4, text="合并输出到", bg=CARD, fg=TEXT,
                 font=(FONT, 10)).pack(side="left", padx=(0, 8))
        self.ent_merge_out = ttk.Entry(l4, textvariable=self.vars["merge_out"])
        self.ent_merge_out.pack(side="left", fill="x", expand=True, padx=(0, 6))
        Tip(self.ent_merge_out, MERGE_OUT_TIP)
        ttk.Button(l4, text="选择…", style="Small.TButton",
                   command=self.on_pick_merge_out).pack(side="left")

        self._opt_rows = {"scale_row": l1, "suffix_row": l2,
                          "print_mode_row": l3, "orient_row": l3b,
                          "merge_out_row": l4}
        return card

    def _suffix_text(self) -> str:
        """当前模式的输出后缀 —— **两个模式各记各的**（见 SUFFIX_VAR_BY_MODE）。

        ⚠️ 别再写 `self.vars["suffix"]`：那会永远读替换模式那一份。
        ⚠️ 合并模式没有后缀这一说（一份产物一个路径），返回空串而不是 KeyError ——
           `_collect_options` 每个模式都会调它一次。
        """
        if not caps_of(self.mode).suffix:
            return ""
        return self.vars[SUFFIX_VAR_BY_MODE[self.mode]].get()

    def _rebind_suffix(self) -> None:
        """把后缀那一格切到**当前模式自己的**变量，并换掉两份悬浮说明。

        ⚠️ 只能从 `_apply_mode` 里调 —— 启动那一次（`_build` 末尾）也要走同一刀，
           挂在 `on_mode` 里会漏掉启动路径。
        """
        if not hasattr(self, "ent_suffix"):
            return
        if not caps_of(self.mode).suffix:      # 合并模式：那一行根本没显示
            return
        self.ent_suffix.configure(textvariable=self.vars[SUFFIX_VAR_BY_MODE[self.mode]])
        self._tip_suffix.text = SUFFIX_TIPS_BY_MODE[self.mode]
        self._tip_suffix_need.text = SUFFIX_NEED_TIPS_BY_MODE[self.mode]

    # ── 操作行 ───────────────────────────────────────────────────────────
    def _build_actions(self) -> tk.Frame:
        """操作行 —— 两套按钮都建好，靠 ACTION_LAYOUT 表决定哪个模式显示哪几个。

        层级（i5 三档）：预览 = 主按钮（实心，第一步就该点它）；
        开始替换 = 强调字按钮（白底 + 深蓝粗字）；其余 = 次按钮。
        一屏只有一颗实心，主色面积远低于 5%。
        """
        bar = tk.Frame(self, bg=BG)
        self.btn_scan = ttk.Button(bar, text="预览（先列队，再配对）",
                                   style="Primary.TButton", command=self.on_preview)
        Tip(self.btn_scan, PREVIEW_TIP)
        self.btn_run = ttk.Button(bar, text="开始替换", style="Accent.TButton",
                                  command=self.on_run, state="disabled")
        self.btn_convert = ttk.Button(bar, text=CONVERT_BTN_TEXT,
                                      style="Primary.TButton",
                                      command=self.on_convert, state="disabled")
        Tip(self.btn_convert, CONVERT_BTN_TIP)
        # 合并打印的主按钮：跟「开始转换」同为实心，但两个模式不同屏出现，
        # 所以一屏仍然只有一颗（i5 第 9 节的硬约束）。
        self.btn_merge = ttk.Button(bar, text=MERGE_BTN_TEXT,
                                    style="Primary.TButton",
                                    command=self.on_merge, state="disabled")
        Tip(self.btn_merge, MERGE_BTN_TIP)
        # 停止：次按钮（白底），只在合并跑着的时候能点。
        # ⚠️ 引擎那一批转换打断不了 —— 按下去是「这批转完就停」，说明写在提示里。
        self.btn_merge_stop = ttk.Button(bar, text="停止", style="Small.TButton",
                                         command=self.on_merge_stop, state="disabled")
        Tip(self.btn_merge_stop,
            "停下来，不再继续拼。\n\n"
            "· 正在跑的那一批转换**不会被打断**（它交给 Office 一个会话在处理），\n"
            "  这一批跑完立刻停；\n"
            "· 停下之后不会生成合并文件，临时目录也会清掉。")
        self.btn_recheck = ttk.Button(bar, text="重新检测引擎",
                                      style="Small.TButton",
                                      command=self.on_recheck_engine)
        Tip(self.btn_recheck, ENGINE_RECHECK_TIP)
        self.lbl_engine = tk.Label(bar, text="", bg=BG, fg=TEXT2, font=(FONT, 10))
        return bar

    # ── ①②③ 三个调整模块（只有替换模式）─────────────────────────────────
    def _build_modules(self) -> tk.Frame:
        """①②③ 三个模块 —— **只在「PDF 首页替换」里出现**（转换模式没有页可调）。

        ⭐ 三个模块**各占一行**：全挤在一行会逼着用户把窗口拉到最大才看得全。
        用 grid 让第 0 列自动按最宽的标签定宽，三行按钮自然左对齐（不写死像素）。
        """
        card = tk.Frame(self, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1)
        tk.Label(card, text="调整", bg=CARD, fg=TEXT2,
                 font=(FONT, 10, "bold")).pack(anchor="w", padx=10, pady=(6, 2))
        mods = tk.Frame(card, bg=CARD)
        for r, mod in enumerate(MODULE_ROWS):
            tk.Label(mods, text=mod.tag, bg=CARD, fg=TEXT2,
                     font=(FONT, 10)).grid(row=r, column=0, sticky="w",
                                           padx=(0, 10), pady=2)
            cell = tk.Frame(mods, bg=CARD)
            cell.grid(row=r, column=1, sticky="w")
            for spec in mod.buttons:
                b = ttk.Button(cell, text=spec.text, width=spec.width,
                               style="Small.TButton", state="disabled",
                               command=self._mod_command(spec))
                b.pack(side="left", padx=(0, 6))
                setattr(self, spec.attr, b)
                if spec.tip:
                    Tip(b, spec.tip)
            if mod.hint:
                tk.Label(cell, text=mod.hint, bg=CARD, fg=MUTED,
                         font=(FONT, 9)).pack(side="left", padx=(8, 0))
        self.btn_order.config(text=self._order_text())   # 文字随正序/倒序变，建完再填
        self.mod_hint_row = mods        # 那 2×3 的网格本身（回归脚本量它）
        self.box_modules = card         # 整块（切模式时被打包/收起的是它）
        mods.pack(fill="x", padx=10, pady=(0, 8))
        return card

    # ── 底部状态条 ───────────────────────────────────────────────────────
    def _build_bottom(self) -> tk.Frame:
        """底部状态条**先 pack**（且 side="bottom"），再 pack 表格：
        窗口被压矮时牺牲的是表格行数，而不是把「就绪 / 进度」挤出可视区。

        ⚠️ 这一条**通栏**（左右不留边、底色 #eceef1）—— 它给窗口收口，
           留了边就变成又一张卡片，反而更碎。
        """
        strip = tk.Frame(self, bg=STATUS_BG)
        inner = tk.Frame(strip, bg=STATUS_BG)
        inner.pack(fill="x", padx=12, pady=5)
        self.pb = ttk.Progressbar(inner, style="Cover.Horizontal.TProgressbar",
                                  mode="determinate", length=200)
        self.pb.pack(side="left")
        self.status = tk.Label(inner, text=STATUS_READY, bg=STATUS_BG, fg=TEXT,
                               font=(FONT, 10), anchor="w", justify="left")
        self.status.pack(side="left", padx=(10, 0), fill="x", expand=True)
        return strip

    def _build_dirline(self) -> tk.Label:
        """⭐ 「原件来源有哪些」—— 就贴在表格正上方的一排小灰字。

        用户要求：「原件的名字的上方加一排小的地址，说明这是哪一个文件夹里面的」。
        来源可能有多个，每个来源里的文件混在一起按文件名排（见排序规则表），
        所以每行**另外还有一列「所在文件夹」**说明这一份是哪家的。
        """
        self.lbl_srcdir = tk.Label(self, text=SRC_DIR_EMPTY, bg=BG, fg=TEXT2,
                                   font=(FONT, 10), anchor="w", justify="left")
        return self.lbl_srcdir

    def _build_table(self) -> tk.Frame:
        """主体表格：白底 + 浅蓝表头 + 1px 框（配色与总表 / 5.0 同一套 tint）。"""
        mid = tk.Frame(self, bg=CARD, highlightbackground=BORDER,
                       highlightthickness=1)
        self.tree = ttk.Treeview(mid, columns=[c[0] for c in COLUMNS],
                                 show="headings", height=5)
        vs = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        for st, color in STATUS_COLORS.items():
            self.tree.tag_configure(st, foreground=color)
        self.tree.tag_configure("odd", background=ZEBRA)
        self.tree.bind("<Double-1>", self.on_row_double)
        self.tree.bind("<<TreeviewSelect>>", lambda *_a: self._sync_adjust_buttons())
        self.tree.bind("<Button-2>", self._popup)
        self.tree.bind("<Button-3>", self._popup)

        # 方向是**整批**操作（工具栏③），右键里不放单页旋转 —— 免得两套入口打架
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="改用扫描件第…页", command=self.on_assign_page)
        self.menu.add_command(label="忽略此行（不处理）", command=self.on_ignore)
        self.menu.add_separator()
        self.menu.add_command(label="在访达/资源管理器中显示原件", command=self.on_reveal)

        # ⭐ 转换模式的右键菜单**单开一份**：那边没有「扫描页」这一说，
        #    留着那一项会是个永远点不出反应的死菜单项（比藏起来更差）。
        self.menu_conv = tk.Menu(self, tearoff=0)
        self.menu_conv.add_command(label="忽略此行（不转这一份）", command=self.on_ignore)
        self.menu_conv.add_separator()
        self.menu_conv.add_command(label="在访达/资源管理器中显示原件", command=self.on_reveal)

        # 合并模式的菜单跟转换那份同构（同样没有「扫描页」），但文案要说「拼」。
        self.menu_merge = tk.Menu(self, tearoff=0)
        self.menu_merge.add_command(label="忽略此行（不拼这一份）", command=self.on_ignore)
        self.menu_merge.add_separator()
        self.menu_merge.add_command(label="在访达/资源管理器中显示原件", command=self.on_reveal)
        return mid

    def _build(self) -> None:
        """页面顺序 = 这张表：标题 → 模式 → 唯一输入块 → 扫描件 → 选项 →
        操作行 → ①②③ → （表格 + 地址行 + 底部状态条）。

        ⚠️ 左右留白写在各块的 pack 参数里，**不写在外层 padding** ——
           底部状态条要通栏，外层一留白它就跟着缩进去了。
        """
        self._panes = []
        self._pane(self._build_header(), ALL_MODES, fill="x",
                   padx=12, pady=(8, 4))
        self._pane(self._build_mode_bar(), ALL_MODES, fill="x",
                   padx=12, pady=(0, 6))
        self._pane(self._build_sources(), ALL_MODES, fill="x", padx=12)
        self._pane(self._build_scan_box(), SWAP_ONLY, fill="x",
                   padx=12, pady=(5, 0))
        self._pane(self._build_options(), ALL_MODES, fill="x",
                   padx=12, pady=(5, 0))
        self._pane(self._build_actions(), ALL_MODES, fill="x",
                   padx=12, pady=(7, 5))
        self._pane(self._build_modules(), SWAP_ONLY, fill="x",
                   padx=12, pady=(0, 5))
        self._pane(self._build_bottom(), ALL_MODES, side="bottom", fill="x")
        self._pane(self._build_dirline(), ALL_MODES, fill="x",
                   padx=14, pady=(0, 2))
        self._pane(self._build_table(), ALL_MODES, fill="both", expand=True,
                   padx=12, pady=(0, 8))
        self._apply_mode()

    # ── 切模式 ───────────────────────────────────────────────────────────
    def _apply_mode(self) -> None:
        """按 `_panes` / ACTION_LAYOUT / OPTION_LAYOUT 三张表把界面摆成当前模式的样子。"""
        for p in self._panes:                       # 先全部收起，再按表的顺序重摆
            p.widget.pack_forget()
        for p in self._panes:
            if self.mode in p.modes:
                p.widget.pack(**p.opts)
        for name, opts, modes in ACTION_LAYOUT:
            w = getattr(self, name)
            if self.mode in modes:
                w.pack(**opts)
            else:
                w.pack_forget()
        # ⚠️ 选项那两行**先全部收起来，再按表的顺序重摆** —— 直接 pack 等于「追加到
        #    末尾」，切几次模式之后顺序会跟着切换历史乱掉（4.0 之前就这么乱过：
        #    缩放勾选框一会儿在后缀上面、一会儿在下面）。
        for name, _opts, _modes in OPTION_LAYOUT:
            self._opt_rows[name].pack_forget()
        for name, opts, modes in OPTION_LAYOUT:
            w = self._opt_rows[name]
            if self.mode in modes:
                w.pack(**opts)
        self._apply_columns()
        for spec in MODE_SPECS:
            b = self._mode_buttons[spec.key]
            b.configure(style=("ModeOn.TButton" if spec.key == self.mode
                               else "ModeOff.TButton"))
        self.lbl_mode.config(text=self._mode_blurb())
        self._refresh_source_tips()
        self._refresh_engine_label()
        self._rebind_suffix()          # 先换变量，再刷那两格派生显示
        self._refresh_suffix_hint()
        # ⭐ 来源清单**两个模式各记各的**（见 _sources_by_mode）：灰字和清单内容
        #    都得跟着换 —— 漏了这一步，表现就是「切过去还看见上一批文件」。
        #    （那块原来的标题行 2026-10-23 已整条删掉，所以这里不再有换标题这一步。）
        self.lbl_src_hint.config(text=_plain(self._src_hint()))
        self._refresh_source_list()    # 内部会顺手刷「原件来自…」那行
        self._sync_run_button()
        self._sync_adjust_buttons()
        # 还没挑过任何文件时，状态栏回到这个模式的「就绪」那句
        if (not self.busy and self.plan is None and self.conv_plan is None
                and self.merge_plan is None):
            self.status.config(text=caps_of(self.mode).ready, foreground=TEXT)
        # 合并模式那两句派生文案（打印方式 / 页面方向说明跟着选项变）
        self._refresh_merge_hint()
        self._refresh_orient_hint()

    # ── 唯一输入块的拖入高亮 ─────────────────────────────────────────────
    def _dz_tint(self, on: bool) -> None:
        """拖入悬停：整块底换浅蓝、框线换主色。

        ⚠️ 整块（提示行、按钮那一列）一起换 —— 只亮提示行会让人以为
           下半截不能拖。标题行删掉之后，这一整块换色就是拖放的主要反馈。
        """
        bg = DZ_BG_ON if on else DZ_BG
        for w in getattr(self, "_dz_tintable", ()):
            try:
                w.config(background=bg)
            except tk.TclError:
                pass
        try:
            self.box_sources.config(
                highlightbackground=(PRIMARY if on else BORDER),
                highlightcolor=(PRIMARY if on else BORDER))
        except tk.TclError:
            pass

    def _mode_blurb(self) -> str:
        for spec in MODE_SPECS:
            if spec.key == self.mode:
                return "现在这个模式：" + spec.blurb
        return ""

    def _refresh_source_tips(self) -> None:
        """来源清单那几个按钮的悬浮说明按模式换一份（表见 SRC_TIPS_BY_MODE）。

        ⚠️ 悬浮文字是**弹出来那一刻**才读 `Tip.text` 的，所以这里改一下就够了，
           不用把 Tip 拆了重建。
        ⭐ 「也可以拖进来」那句只在 `dnd_ok` 时追加 —— 挂不上拖放还这么说，
           等于教用户去做一件做不到的事。
        """
        for method, per_mode in SRC_TIPS_BY_MODE.items():
            tip = getattr(self, "_src_tips", {}).get(method)
            if tip is None:
                continue
            text = per_mode.get(self.mode)
            if text:
                tip.text = text + ("\n\n" + DND_TIP_LINE if self.dnd_ok else "")

    def _src_hint(self) -> str:
        """「原件来源」框下面那行灰字（按模式换，挂上了拖放再多一句）。"""
        base = SRC_HINT_BY_MODE[self.mode]
        return base + DND_HINT if self.dnd_ok else base

    # ── 把文件拖进来（2026-10-09 用户要的）────────────────────────────────
    #
    # 用户的原话：软件摆屏幕右边、文件夹摆左边，「左边选中什么文件，直接给它
    # 拖进去」。一次挑十几二十份要一个个点，确实累。
    #
    # ⭐ 注册在**主窗口 + 本框架**两处就够，不用铺满每个控件：
    #    · Windows 上 OLE2 只能注册在顶层窗口（子控件没有独立 HWND），
    #      注册顶层必然覆盖整个客户区；
    #    · macOS 上 Cocoa 从最内层 view 往下找，子控件自己没注册就会落到
    #      注册过的父级 —— 所以注册「顶层 + 铺满客户区的 self」两处最稳。
    # ⚠️ 万一真有一处漏接（比如某个控件抢了事件），也不会有脏数据：
    #    `_add_sources` 本来就按路径去重（见那里），重复触发只是白跑一遍。
    def _setup_dnd(self) -> bool:
        """把「接拖入」挂到窗口上。挂不上返回 False（界面照常，只是没这功能）。

        ⚠️ **根窗口本身没有这些方法** —— 别以为漏了它。tkinterdnd2 是把
           `drop_target_register` / `dnd_bind` 补丁到 `tkinter.BaseWidget` 上的，
           而 `tkinter.Tk` 继承的是 `(Misc, Wm)`，**不是** `BaseWidget`
           （踩过：给 root 调这一步直接 AttributeError）。所以这里靠 `hasattr`
           逐个筛，只注册真有这套方法的控件。
        ⭐ 那「拖到窗口任意位置」怎么办：注册在**本框架（`self`）**上就够了 ——
           它是铺满整个客户区的那个 Frame，落在任何子控件上的 drop 都会冒泡上来
           （Windows 的 OLE2 按坐标从内往外找注册过的控件；macOS 的 Cocoa 沿
           view 层次上溯）。子控件自己没注册就不会把事件抢走。
        """
        if not _DND_READY:
            return False
        ok = False
        for w in (self, self.winfo_toplevel()):
            if not hasattr(w, "drop_target_register"):
                continue
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
                w.dnd_bind("<<DropEnter>>", self._on_drop_enter)
                w.dnd_bind("<<DropLeave>>", self._on_drop_leave)
                ok = True
            except Exception:  # noqa: BLE001 —— 挂不上就当没这功能，别再往上炸
                continue
        return ok

    def _drag_hint(self, on: bool) -> None:
        """拖着东西悬在窗口上时的提示。

        ⚠️ 只改整块底色 +（清单还空时）那句占位话，**绝不碰清单内容** —— 提示是临时的，
           用户中途把拖拽取消（拖到一半按 Esc）时，数据必须一个字节都没动。
        """
        if not self.dnd_ok:
            return
        self._dz_tint(on)
        # 标题行删掉之后，能写字的地方只剩清单框里的占位话。只在**清单为空**时改：
        # 已经有文件列着的时候，那片是用户正看着的文件名，别往上盖字。
        if not self.sources:
            try:
                self.lbl_src_empty.config(text=(SRC_DROP_HINT if on else SRC_EMPTY_HINT))
            except tk.TclError:
                pass

    def _on_drop_enter(self, _e: tk.Event) -> str:
        """有东西被拖着进了窗口范围 —— 亮个「松手就加进来」的提示。

        ⚠️ 返回值是 tkdnd 要的**拖放动作**（copy / move / link / ask），
           不返回的话光标那边会显示成「不允许放置」。下同。
        """
        self._drag_hint(True)
        return "copy"

    def _on_drop_leave(self, _e: tk.Event) -> str:
        """拖着东西离开了窗口范围（或者中途放手取消）—— 提示收回去。"""
        self._drag_hint(False)
        return "copy"

    def _on_drop(self, e: tk.Event) -> str:
        """松手了 —— 把落下的文件 / 文件夹加进「原件来源」。

        路径解析见 `drop_paths`（带空格的路径就靠它才不被劈成两半）。
        加进去之后走的是**和「＋ 添加文件…」完全同一条路**（`_add_sources`）——
        所以去重、按模式作废旧结果、转换模式顺手重列队，全都自动跟着，不用另写。
        """
        self._drag_hint(False)
        splitlist = getattr(getattr(self, "tk", None), "splitlist", None)
        paths = drop_paths(getattr(e, "data", ""), splitlist)
        if paths:
            self._add_sources(paths)
        return "copy"

    def _apply_columns(self) -> None:
        """表格列随模式换 —— 两套列名用的是同一张 Treeview，只换列定义。

        ⚠️ 换列**必须清表**：Treeview 的行是按**位置**取值，列一变旧行的每格都会错位。
        """
        cols = COLUMNS_BY_MODE[self.mode]
        self._clear_table()
        self.tree.configure(columns=[c[0] for c in cols])
        for key, title, width, anchor in cols:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor,
                             stretch=(key in ("src", "dir", "out", "note")))

    def on_mode(self, key: str) -> None:
        """点顶部的大按钮切模式。两边的队列各留各的，切回来还在。"""
        if self.busy or key == self.mode or key not in {m.key for m in MODE_SPECS}:
            return
        self.mode = key
        self._apply_mode()
        caps = caps_of(self.mode)
        # ⚠️ **已经列过队的就别重列** —— 重列会把上一轮的结果（转换成功 / 页数）
        #    抹回「待转换」，用户切去别的模式看一眼再切回来，结果就凭空没了。
        #    真变了来源时 `_sources_changed` 早把队列清成 None 了，
        #    所以「队列 is None」正好等价于「得重列」。
        if caps.queue == "conv_plan":
            if self.conv_plan is None:
                self._refresh_convert_plan()
            else:
                self._fill_table()
                self._sync_run_button()
        elif caps.queue == "merge_plan":
            if self.merge_plan is None:
                self._refresh_merge_plan()
            else:
                self._fill_table()
                self._sync_run_button()
        else:
            self._fill_table()
        self.lbl_mode.config(text=self._mode_blurb())

    # ── 选择路径 ─────────────────────────────────────────────────────────
    def _spec(self, key: str) -> FolderRow:
        for s in FOLDER_SPECS:
            if s.key == key:
                return s
        raise KeyError(key)

    def _pick(self, key: str) -> None:
        """弹系统选择框。

        ⚠️ `cur` 只可能来自**这一次会话**里刚选过的那份 —— 因为路径不再写进配置
           （见 CONFIG_PERSIST_VARS），下次打开一定是空的，起手就去用户主目录。
        """
        spec = self._spec(key)
        cur = self.vars[key].get().strip()
        start = cur if cur else str(Path.home())
        if spec.kind == "file":
            got = filedialog.askopenfilename(
                title=f"选择{spec.label}",
                initialdir=str(Path(start).parent if Path(start).is_file() else start),
                filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        else:
            # 扩展位：目前 FOLDER_SPECS 里没有 kind="dir" 的行（原来的「原件文件夹」
            # 已经改成来源清单）。以后要加「选文件夹」的行，表里添一行即可。
            got = filedialog.askdirectory(title=f"选择{spec.label}",
                                          initialdir=start if Path(start).is_dir() else str(Path.home()))
        if got:
            self.vars[key].set(got)

    # ── 原件来源清单 ──────────────────────────────────────────────────────
    def _source_label(self, path: str) -> str:
        """清单里一行怎么显示 —— 方括号里报**类型**（两个模式认的文件不一样）。"""
        p = Path(path).expanduser()
        if p.is_dir():
            kind = "文件夹"
        elif p.is_file():
            kind = p.suffix.lstrip(".").upper() or "文件"
        else:
            kind = "找不到"
        return f"[{kind}] {path}"

    def _refresh_source_list(self) -> None:
        """把 self.sources 灌进 Listbox。

        ⚠️ 这里显示的是**添加顺序**，跟队列顺序无关 —— 队列一律按文件名排
           （见 core 的 ① 排序规则表），所以清单的先后不参与排序、也不影响结果。
        """
        self.lst_sources.delete(0, "end")
        for s in self.sources:
            self.lst_sources.insert("end", self._source_label(s))
        # 空态：给一句「该做什么」，别留一片白（i5 第 8 节）。
        # ⚠️ 占位话贴在清单框**里面**（place in_=清单框），所以这里收起来 / 放出去
        #    都用 place 那一对，别退回 pack —— 用 pack 会凭空多出一行。
        lbl_empty = getattr(self, "lbl_src_empty", None)
        if lbl_empty is not None:
            if self.sources:
                lbl_empty.place_forget()
            elif not lbl_empty.winfo_manager():
                lbl_empty.config(text=SRC_EMPTY_HINT)
                lbl_empty.place(in_=self.lst_sources, x=4, y=3)
        self._refresh_srcdir_label()

    def _add_sources(self, paths: list[str]) -> None:
        """往清单里追加来源（同一路径加两次直接跳过）。

        ⭐ 2026-11-02 起：**文件夹**来源会按当前模式「展开」成里面的可处理文件 ——
           转换模式只收认得的 Word / Excel 等（PDF、未知格式一律不收），
           拼接模式只收 PDF。口径与 core 的 `collect_convertibles` /
           `collect_originals` 完全一致（只扫这一层、不钻子目录）。
           展开之后清单里就是一份份文件，跟逐份拖进来没两样；用户能逐个看到、
           逐个删。之前 2026-10-08 删「添加文件夹」是怕误收共表 / DWG ——
           现在按扩展名过滤，正好把不支持的挡在门外。
        """
        # 当前模式认的扩展名（小写、含点）—— 从能力表来，别在这里写 if
        exts = caps_of(self.mode).exts
        # ⭐ 先记下「此刻有没有旧结果」—— 有的话 `_sources_changed` 里的 `_invalidate`
        #    会自己把状态栏换成「预览作废了、重新来一遍」，我们别再去盖它；
        #    没有的话状态栏是空白的，才轮到我们补一句「已加进 N 份」。
        had = (self.plan is not None or self.conv_plan is not None
               or self.merge_plan is not None)
        added = 0
        skipped = 0
        empty_dirs = 0        # 拖进来的文件夹里，一个认得的文件都没有
        for p in paths:
            p = str(p).strip()
            if not p:
                continue
            entry = Path(p).expanduser()
            if entry.is_dir():
                # 文件夹 → 只收它这一层里认得的文件（子目录不钻，与 core 一致）
                found = False
                for c in sorted(entry.glob("*")):
                    if not (c.is_file() and c.suffix.lower() in exts):
                        continue
                    found = True
                    cp = str(c)
                    if cp in self.sources:
                        skipped += 1
                    else:
                        self.sources.append(cp)
                        added += 1
                if not found:
                    empty_dirs += 1
            elif p not in self.sources:
                # 单个文件（或不存在）—— 照旧直接加，认不认得交给 core 去判
                self.sources.append(p)
                added += 1
            else:
                skipped += 1
        if not added:
            # 一点没加进来，也要说一句为什么 —— 别让用户以为拖了个寂寞。
            if empty_dirs:
                self.status.config(
                    text="这个文件夹里没有当前模式能处理的文件"
                         f"（{caps_of(self.mode).kind_words}）。",
                    foreground=WARN)
            return
        self._refresh_source_list()
        self._sources_changed("来源变了")
        if not had:
            bits = [f"已加进 {added} 份"]
            if skipped:
                bits.append(f"跳过 {skipped} 份重复的")
            if empty_dirs:
                bits.append("有文件夹里没找到能处理的文件")
            self.status.config(text="，".join(bits), foreground=TEXT2)
        # ⚠️ 这里**不用**调度保存：清单不落盘（见 CONFIG_PERSIST_VARS）。

    # ⭐ 2026-10-08 删掉了 on_add_folder / _pick_many_folders（多选文件夹那套 osascript）。
    #    用户原话：一个副文件夹里「既有子文件夹又有文件」，还可能塞着十几二十份
    #    共表 / DWG —— 整个夹收进来是误收；真要批量挑，用「＋ 添加文件…」在同一层里
    #    一次框选更准。子文件夹里的文件就点进去再挑一次。
    #    （core 仍然支持「文件夹」这种来源 —— CLI 的 --src 与回归脚本都还在用，
    #      只是界面上不再开这个入口。）

    def on_add_files(self) -> None:
        """挑要处理的文件。**每个模式认的类型不一样**，所以选择框跟着模式变。"""
        caps = caps_of(self.mode)
        if caps.key == "swap":
            title = "添加要处理的 PDF（可按住 Cmd / Ctrl 多选）"
            types = [("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        else:
            title = f"添加要{'转' if caps.key == 'convert' else '拼'}的 " \
                    f"{caps.kind_words}（可按住 Cmd / Ctrl 多选）"
            # ⚠️ 扩展名**不按平台过滤**（收件口径一致，见 core 的 CONVERT_EXTS）：
            #    将来若加进「本机转不了、只有别的系统能转」的格式，这样也会**照收**，
            #    然后给一句明确的「本机转不了、去哪转」，比「选都选不中」有用得多。
            types = [(caps.kind_words, " ".join(f"*{e}" for e in caps.exts)),
                     ("所有文件", "*.*")]
        got = filedialog.askopenfilenames(
            title=title, initialdir=str(Path.home()), filetypes=types)
        self._add_sources(list(got or ()))

    def on_remove_sources(self) -> None:
        sel = sorted(self.lst_sources.curselection(), reverse=True)
        if not sel:
            return
        for i in sel:
            if 0 <= i < len(self.sources):
                del self.sources[i]
        self._refresh_source_list()
        self._sources_changed("来源变了")

    def on_clear_sources(self) -> None:
        if not self.sources:
            return
        self.sources = []
        self._refresh_source_list()
        self._sources_changed("来源清空了")

    def _sources_changed(self, reason: str) -> None:
        """来源一动，旧的结果就作废；有队列的两个模式顺手把队重新列一遍。

        ⚠️ 转换 / 合并模式**没有「预览」这一步**（用户要求「底下就直接开始」），
           所以这里必须自己把表刷出来 —— 不然加了文件表里还是空的。
           只认扩展名、不读文件内容，所以同步做也很快，不用开线程。
        """
        self._invalidate(reason)
        queue = caps_of(self.mode).queue
        if queue == "conv_plan":
            self._refresh_convert_plan()
        elif queue == "merge_plan":
            self._refresh_merge_plan()

    def _on_source_double(self, _e: tk.Event | None = None) -> None:
        """双击清单里的一行 → 在访达 / 资源管理器里定位它。"""
        sel = self.lst_sources.curselection()
        if not sel:
            return
        p = Path(self.sources[sel[0]]).expanduser()
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(p)], check=False)
            elif os.name == "nt":
                subprocess.run(["explorer", str(p)], check=False)
            else:
                subprocess.run(["xdg-open", str(p if p.is_dir() else p.parent)], check=False)
        except Exception:  # noqa: BLE001
            pass

    def _refresh_srcdir_label(self) -> None:
        """表格上方那排小地址 —— 这批文件是从哪些来源收来的、产物会写在哪。"""
        if not self.sources:
            self.lbl_srcdir.config(text=SRC_DIR_EMPTY, foreground=TEXT2)
            return
        gone = [s for s in self.sources if not Path(s).expanduser().exists()]
        if gone:
            self.lbl_srcdir.config(
                text=f"有 {len(gone)} 个来源已经不存在了：{Path(gone[0]).name}"
                     + ("…" if len(gone) > 1 else ""),
                foreground=DANGER)
            return
        caps = caps_of(self.mode)
        if caps.queue == "conv_plan":
            n = len(self.conv_plan.items) if (self.conv_plan and self.conv_plan.items) else 0
        elif caps.queue == "merge_plan":
            n = len(self.merge_plan.items) if (self.merge_plan and self.merge_plan.items) else 0
        else:
            n = len(self.plan.items) if (self.plan is not None and self.plan.items) else 0
        tail = caps.dir_tail
        head = (f"原件来自：{self.sources[0]}" if len(self.sources) == 1
                else f"原件来自：{len(self.sources)} 个来源")
        got = f"　·　共收到 {n} 份" if n else ""
        self.lbl_srcdir.config(text=f"{head}{got}　·　{tail}", foreground=TEXT2)

    # ── 模块①：扫描件页序（正序 ⇄ 倒序，整批）─────────────────────────────
    def _order_text(self) -> str:
        return "页序：倒序" if self.vars["reverse"].get() else "页序：正序"

    def on_flip_order(self) -> None:
        """一键把整批的页序反过来（扫描仪倒着出纸时用）。

        ⚠️ 是「按模式重算」不是「把当前结果再反一次」—— 来回点是幂等的，
           用户不用记自己按过几次；代价是手动调过的页码会被清回默认，
           所以顺序必须是：先定页序 → 再调个别页码。
        """
        if self.plan is None or self.plan.blocked:
            return
        rev = not self.vars["reverse"].get()
        self.vars["reverse"].set(rev)
        C.reindex(self.plan, rev)
        self._apply_rotations()
        self._fill_table()
        self.btn_order.config(text=self._order_text())
        self.status.config(
            text=f"页序已切成{'倒序' if rev else '正序'}（表格「扫描页」列已重算，"
                 f"手动调过的页码也一并清回默认）。",
            foreground=TEXT)

    # ── 模块②：页码对应（一页只许一个人用）───────────────────────────────
    def _swap_page(self, delta: int) -> None:
        """与相邻行对调扫描页 —— 行不动，只换「这一行用哪一页」。"""
        sel = self.tree.selection()
        if not sel or self.plan is None or self.plan.blocked:
            return
        idx = self.tree.index(sel[0]) + 1
        if not C.swap_pages(self.plan, idx, delta):
            return
        self._apply_rotations()
        pos = idx + delta - 1
        self._reselect_refresh(pos)

    def _mod_command(self, spec: BtnSpec):
        """把 BtnSpec 的「方法名 + 参数」拼成按钮回调。

        参数用默认值固化在闭包里 —— 直接在 for 循环里写 lambda 会踩晚绑定，
        五个按钮最后全指向最后一次迭代的值。
        """
        return lambda m=spec.method, a=spec.args: getattr(self, m)(*a)

    def _adjust_buttons(self) -> tuple:
        """①②③ 三个模块的全部按钮 —— 由 MODULE_ROWS 派生，一起控，不会漏掉某一个。"""
        return tuple(getattr(self, n) for n in ADJUST_BUTTONS)

    def _sync_adjust_buttons(self) -> None:
        """只在「已经有配好扫描件的队列」时才让①②③的按钮可点 ——

        只列队（还没选扫描件）时页码/方向根本无从谈起，按了也没意义。
        """
        on = (self.plan is not None and self.plan.paired and not self.plan.blocked
              and bool(self.tree.get_children()))
        state = "normal" if on else "disabled"
        for b in self._adjust_buttons():
            b.config(state=state)

    # ── 模块③：方向（整批转，没有单页入口）────────────────────────────────
    def on_turn_all(self, delta: int) -> None:
        """整份扫描件**所有页**统一左转 / 右转 90°（整批都躺倒了时用）。"""
        self._turn_pages(None, delta, "全部")

    def on_turn_land(self, delta: int) -> None:
        """只转「配给横向原件」的那几页（竖版没扫错、只有横版方向不对时用）。"""
        if self.plan is None or self.plan.blocked:
            return
        pages = {it.scan_index for it in self.plan.items if C.is_landscape(it.src_size)}
        if not pages:
            self.status.config(text="这批原件里没有横向的（首页都是竖向），横向按钮用不上。",
                               foreground=TEXT)
            return
        self._turn_pages(pages, delta, "横向")

    def _turn_pages(self, pages: set[int] | None, delta: int, scope: str) -> None:
        """在当前方向的基础上再转 —— 自动值就此固化成人工值，之后不再跟着原件横竖变。"""
        if self.plan is None or self.plan.blocked:
            return
        hit = 0
        for it in self.plan.items:
            if pages is not None and it.scan_index not in pages:
                continue
            self.rotations[it.scan_index] = C.turn_deg(it.rotation, delta)
            hit += 1
        if not hit:
            return
        self._apply_rotations()
        self._reselect_refresh()
        self.status.config(
            text=f"已把{scope} {hit} 页统一{'左' if delta < 0 else '右'}转 90°"
                 f"（看表格「方向」列；转错了点「方向复位」）。",
            foreground=OK_GREEN)

    def on_reset_rot(self) -> None:
        """清掉全部手动调整，回到默认规则（横向原件那一页左转 90°）。"""
        if not self.rotations:
            return
        self.rotations = {}
        self._apply_rotations()
        self._reselect_refresh()
        n = sum(1 for it in (self.plan.items if self.plan else [])
                if it.rotation)
        self.status.config(
            text="方向已复位到默认规则" + (f"（{n} 页按默认左转 90°）" if n else "（全部正向）"),
            foreground=TEXT)

    def _reselect_refresh(self, pos: int | None = None) -> None:
        """刷表并保住选中行（pos 给定时改选到那一行）。

        ⚠️ 只能按**位置**恢复，不能按行 id —— _fill_table 是整表重建，
           填完旧 id 就已经失效了（selection_set 会抛 "Item Ixxx not found"）。
        """
        if pos is None:
            sel = self.tree.selection()
            pos = self.tree.index(sel[0]) if sel else -1
        self._fill_table()
        kids = self.tree.get_children()
        if 0 <= pos < len(kids):
            self.tree.selection_set(kids[pos])
            self.tree.see(kids[pos])

    def _apply_rotations(self) -> None:
        """把「人工 > 自动 > 正向」写回 plan。

        口径统一走 core 的 apply_rotations，界面自己不算 —— 两边各算一套迟早漂移。
        """
        if self.plan is None:
            return
        C.apply_rotations(self.plan, self.rotations)

    # ── 输出文件名后缀的实时预览（替换模式必填 / 转换模式可不填）──────────
    def _sample_stem(self) -> str:
        """后缀预览拿谁当例子 —— 用**当前模式**队列里的第一份。"""
        plan = self._current_plan()
        items = getattr(plan, "items", None) if plan is not None else None
        if items:
            return items[0].src.stem
        return SUFFIX_SAMPLE_FALLBACK

    def _suffix_ready(self) -> bool:
        """后缀够不够用 —— 界面这一侧的唯一判据，core 那边还有同样一条。

        ⚠️ **只对替换模式有意义**：转换模式后缀可空（产物 .pdf vs 原件 .xlsx）。
        ⚠️ 读的是**当前模式自己那一份**后缀（见 SUFFIX_VAR_BY_MODE）。
        """
        if not caps_of(self.mode).suffix:
            return True                    # 没有后缀这一项 → 不存在「没填」这回事
        return bool(C.clean_output_suffix(self._suffix_text()))

    def _refresh_suffix_hint(self) -> None:
        """右边那格 + 那个「必填 / 可不填」标签，都随模式变。"""
        if not caps_of(self.mode).suffix:
            return                         # 合并模式：这一行不显示，不用刷
        suffix = C.clean_output_suffix(self._suffix_text())
        if caps_of(self.mode).key == "convert":
            self.lbl_suffix_need.config(text="可不填", foreground=TEXT2)
            self.lbl_suffix.config(text=f"{self._sample_stem()}{suffix}.pdf",
                                   foreground=PRIMARY_D)
            return
        self.lbl_suffix_need.config(text="必填", foreground=DANGER)
        if not suffix:
            self.lbl_suffix.config(text="不填不让跑（产物会跟原件同名、盖掉原件）",
                                   foreground=DANGER)
            return
        self.lbl_suffix.config(text=f"{self._sample_stem()}{suffix}.pdf",
                               foreground=PRIMARY_D)

    def _on_suffix_change(self) -> None:
        self._refresh_suffix_hint()
        # 转换模式表格里「产物名」这一列直接跟着后缀走 —— 后缀一变，
        # 上一轮转出来的结果（产物名 / 页数 / 成功失败）说的就是**别的目标**了。
        # 留着会让人以为「这些已经转好了」，所以整片进度清回「待转换」。
        if caps_of(self.mode).key == "convert":
            self._reset_convert_progress()
        self._sync_run_button()

    def _reset_convert_progress(self) -> None:
        """把转换队列的进度清回「待转换」（后缀一改就得清，理由见 _on_suffix_change）。"""
        self.conv_results = {}
        for it in (self.conv_plan.items if self.conv_plan else []):
            it.status = C.ConvStatus.PENDING.value
            it.note = ""
        self._fill_table()

    def _convertible_count(self) -> int:
        """队列里**还要转**的份数（被手动忽略的那几行不算）。"""
        if self.conv_plan is None:
            return 0
        return sum(1 for it in self.conv_plan.items
                   if it.status != C.ConvStatus.SKIPPED.value)

    def _sync_run_button(self) -> None:
        """三个模式的「开始…」按钮各自算自己的可用性，都在这一处判定。

        ⚠️ 只在这一处判定，`_set_busy` 与 `_after_preview` 都走它 ——
           判断散在多处必然漏掉某种组合（后缀填好但还没预览 / 预览好了又把后缀清空）。
        ⚠️ 按钮**都刷**（不管当前模式）：切模式时不用再补一刀，
           而且隐藏中的按钮状态也永远是对的。
        —— 替换模式：不忙 + 已配好扫描件 + 没阻断 + 后缀填了；
        —— 转换模式：不忙 + 有要转的份数 + 没阻断 + 找得到引擎（后缀可空）；
        —— 合并模式：不忙 + 有要拼的份数 + 没阻断 + 输出路径填了。
        """
        swap_ok = (not self.busy and self.mode == "swap" and self.plan is not None
                   and self.plan.paired and not self.plan.blocked and self._suffix_ready())
        self.btn_run.config(state="normal" if swap_ok else "disabled")
        conv_ok = (not self.busy and self.mode == "convert"
                   and self.conv_plan is not None and not self.conv_plan.blocked
                   and self._convertible_count() > 0
                   and bool(self.conv_plan.engines))
        self.btn_convert.config(state="normal" if conv_ok else "disabled")
        merge_ok = (not self.busy and self.mode == "merge"
                    and self.merge_plan is not None and not self.merge_plan.blocked
                    and self._merge_count() > 0
                    and bool(self.vars["merge_out"].get().strip()))
        self.btn_merge.config(state="normal" if merge_ok else "disabled")

    def _current_plan(self):
        """当前模式的主数据（队列）—— 能力表说了算，别再写 if。"""
        return getattr(self, caps_of(self.mode).queue, None)

    def _invalidate(self, reason: str = "") -> None:
        """来源 / 扫描件一变，之前那次的结果就作废了。

        ⚠️ 必须**清掉表格**，不能留着旧队接着调 —— 顺序配对是脆弱链，
           拿着「来源已经变了」的旧队去调页码，调到的一定是错的对应关系。
        """
        had = (self.plan is not None or self.conv_plan is not None
               or self.merge_plan is not None)
        self.plan = None
        self.conv_plan = None
        self.merge_plan = None
        self.merge_result = None
        self.rotations = {}
        self._clear_table()
        self._refresh_srcdir_label()
        self._sync_run_button()
        self._sync_adjust_buttons()
        if had and reason:
            tail = ("之前那次预览作废了，点「预览」重新来一遍。"
                    if caps_of(self.mode).key == "swap"
                    else "队列已经跟着重新列好了。")
            self.status.config(text=f"{reason} —— {tail}", foreground=WARN)

    # ── 配置持久化 ───────────────────────────────────────────────────────
    def _load_config(self) -> None:
        """只恢复白名单里的东西（两个模式各自的后缀 / 缩放 / 页序）+ 窗口大小。

        ⭐ v3.0 起**不再恢复任何路径** —— 「原件来源」清单和「扫描件 PDF」
        每次打开都是空的，重新挑一次即可（理由见 CONFIG_PERSIST_VARS）。
        老配置里的 sources / src / scan 这几个键会被**直接忽略**（读的时候就不认它们），
        下一次保存整个文件被重写，老键自然就没了，不用专门去清。
        """
        if not CONFIG_PATH.is_file():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            self._config_broken = True
            return
        if not isinstance(data, dict):
            self._config_broken = True
            return
        for key in CONFIG_PERSIST_VARS:
            if key not in self.vars or key not in data:
                continue
            try:
                self.vars[key].set(data[key])
            except Exception:  # noqa: BLE001  一条坏数据不该废掉整份配置，跳过它
                pass

        geo = fit_geometry(data.get("window"),
                           self.winfo_toplevel().winfo_screenwidth(),
                           self.winfo_toplevel().winfo_screenheight())
        if geo:
            self.winfo_toplevel().geometry(geo)

    def _save_config(self) -> None:
        """先写临时文件再原子替换；替换被拦就退一步直接写，并收掉 .tmp。

        ⚠️ 只写 CONFIG_PERSIST_VARS 白名单里的键 —— 「原件来源」清单和「扫描件 PDF」
           都不落盘（理由见那张表），所以这个小 JSON 里**永远不会出现路径**。
        """
        self._cancel_save()
        top = self.winfo_toplevel()
        data: dict[str, object] = {
            k: self.vars[k].get() for k in CONFIG_PERSIST_VARS if k in self.vars
        }
        data["_version"] = CONFIG_VERSION
        try:
            data["window"] = [top.winfo_width(), top.winfo_height()]
        except Exception:  # noqa: BLE001
            pass
        text = json.dumps(data, ensure_ascii=False, indent=1)
        tmp = CONFIG_PATH.with_suffix(CONFIG_PATH.suffix + ".tmp")
        try:
            tmp.write_text(text, encoding="utf-8")
            os.replace(tmp, CONFIG_PATH)
            return
        except Exception:  # noqa: BLE001
            pass
        try:
            tmp.unlink(missing_ok=True)
        except Exception:  # noqa: BLE001
            pass
        try:
            CONFIG_PATH.write_text(text, encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    def _schedule_save(self) -> None:
        self._cancel_save()
        self._save_after = self.after(SAVE_DEBOUNCE_MS, self._save_config)

    def _cancel_save(self) -> None:
        if self._save_after:
            try:
                self.after_cancel(self._save_after)
            except Exception:  # noqa: BLE001
                pass
            self._save_after = None

    def _announce_restore(self) -> None:
        """启动时状态栏的一句提示。

        ⭐ v3.0 起只剩「记录读不出来」这一种情况 —— 既然不再恢复任何路径，
           「已填好上次的 XX / XX 已经不存在」那两句就没有出场机会了。
        """
        if self._config_broken:
            self.status.config(text="上次的记录文件读不出来（已忽略，不影响使用）",
                               foreground=DANGER)

    # ── 预览 / 执行 ──────────────────────────────────────────────────────
    def _collect_options(self) -> C.Options:
        """⚠️ 别叫 `_options` —— tkinter.Widget 内部有同名方法，会被覆盖掉。

        ⚠️ 这里**不设 overwrite** —— 它由 `on_run` 里那次覆盖确认弹窗的结果决定
           （用户不要勾选框，要「点一次头」）。
        """
        return C.Options(
            sources=tuple(self.sources),
            scan_pdf=self.vars["scan"].get().strip(),
            reverse_order=bool(self.vars["reverse"].get()),
            scale_to_original=bool(self.vars["scale"].get()),
            suffix=self._suffix_text(),          # 当前模式自己那一份（见 SUFFIX_VAR_BY_MODE）
            rotations=dict(self.rotations),
        )

    def _clear_table(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)

    def _fill_table(self) -> None:
        """按**当前模式**把表格画一遍 —— 三套列定义、三套取值。"""
        self._clear_table()
        if self.mode == "convert":
            self._fill_convert_table()
            return
        if self.mode == "merge":
            self._fill_merge_table()
            return
        if self.plan is None:
            return
        # 「所在文件夹」列：只有**原件来自不止一个文件夹**时才填。
        # 同一层里的名字全挂着一模一样的目录，纯属噪音（这是 v3.0 之前的教训）。
        parents = {it.src.parent for it in self.plan.items}
        show_dir = len(parents) > 1
        for k, it in enumerate(self.plan.items):
            tags = [it.status] if it.status in STATUS_COLORS else []
            if k % 2:
                tags.append("odd")
            self.tree.insert(
                "", "end",
                values=(it.index, it.src.name,
                        (it.src.parent.name if show_dir else ""),
                        it.src_pages or "—",
                        _fmt_size(it.src_size),
                        it.scan_index or "—",
                        _fmt_size(it.scan_size) if it.scan_index else "—",
                        C.ROT_LABELS.get(it.rotation % 360, "正向") if it.scan_index else "—",
                        it.status, it.note),
                tags=tuple(tags))
        self._sync_adjust_buttons()

    def _fill_convert_table(self) -> None:
        """转换模式的表格 —— 列定义见 CONVERT_COLUMNS。

        ⚠️ 「产物名」在**转之前**是预测值（原件名 + 后缀 + .pdf），转之后换成
           **实际落点**（点了「否」让开时会是 xxx_2.pdf，跟预测值不一样）。
           「产物页数」转之前一律「—」—— 那时它还是个 Excel，没有页数可读。
        """
        if self.conv_plan is None:
            return
        suffix = self._suffix_text()             # 转换模式自己那一份后缀
        parents = {it.src.parent for it in self.conv_plan.items}
        show_dir = len(parents) > 1
        for k, it in enumerate(self.conv_plan.items):
            res = self.conv_results.get(it.src)
            out = (res.out.name if (res and res.out) else
                   C.out_pdf_name_for(it.src, suffix))
            pages = str(res.pages) if (res and res.pages) else "—"
            tags = [it.status] if it.status in STATUS_COLORS else []
            if k % 2:
                tags.append("odd")
            self.tree.insert(
                "", "end",
                values=(it.index, it.src.name,
                        (it.src.parent.name if show_dir else ""),
                        it.fmt_label, out, pages, it.status, it.note),
                tags=tuple(tags))

    # ── 合并打印：列队 / 画表 / 执行 ──────────────────────────────────────
    def _merge_count(self) -> int:
        """队列里**还要拼**的份数（被手动忽略的不算）。"""
        if self.merge_plan is None:
            return 0
        return sum(1 for it in self.merge_plan.items
                   if it.status != C.ConvStatus.SKIPPED.value)

    def _refresh_merge_plan(self) -> None:
        """把来源清单算成合并队列并铺到表格上。

        ⚠️ 与 `_refresh_convert_plan` 同一个理由：**同步**跑（只认扩展名、不读内容）。
        ⚠️ 被闸门拦住时**不弹窗**，状态栏说一句就够；真要跑时 `on_merge` 会再弹。
        """
        if self.mode != "merge":
            return
        # 还没有任何来源时，别急着报「没有可合并的文件」—— 那话该在用户真加了东西
        # 之后才说。空手点进来，状态栏给的就该是那句「就绪：把要拼的…」。
        if not self.sources:
            self.merge_plan = None
            self.merge_result = None
            self._fill_table()
            self._refresh_srcdir_label()
            self._refresh_engine_label()
            self._sync_run_button()
            self.status.config(text=caps_of(self.mode).ready, foreground=TEXT)
            return
        opts = self._collect_options()
        try:
            plan = M.plan_merge(opts)
        except Exception as exc:  # noqa: BLE001
            plan = M.MergePlan(blocked=True, block_reason=f"列队时出错：{exc}")
        self.merge_plan = plan
        self.merge_result = None
        # 队列里用到的引擎顺带回填进缓存（省一次探测）
        if plan.engines and self.engines is None:
            self.engines = list(plan.engines)
        self._fill_table()
        self._refresh_srcdir_label()
        self._refresh_engine_label()
        self._sync_run_button()
        if plan.blocked:
            first = (plan.block_reason or "还不能合并").splitlines()[0]
            self.status.config(text=first, foreground=WARN)
            return
        unk = len(plan.unknown)
        tail = f"，另有 {unk} 份认不出格式、已跳过" if unk else ""
        self.status.config(
            text=f"队列里有 {self._merge_count()} 份要拼{tail}。"
                 f"选好打印方式和输出位置，再点「开始合并」。",
            foreground=OK_GREEN)

    def _fill_merge_table(self) -> None:
        """合并模式的表格 —— 列定义见 MERGE_COLUMNS。

        ⚠️ 「页数 / 补白」在**跑之前**一律「—」：那时文件还是 .xlsx，
           有没有页、要不要补白都无从谈起；跑完才有真数。
        """
        if self.merge_plan is None:
            return
        parents = {it.src.parent for it in self.merge_plan.items}
        show_dir = len(parents) > 1
        for k, it in enumerate(self.merge_plan.items):
            pages = str(it.pages) if it.pages else "—"
            # 6.2：首页后插的那张和末尾补的那张分开说 —— 用户对账时看的是
            # 「第 1 页后面那张空白在不在」，跟「末尾凑偶数」不是同一个位置。
            if it.pad_front and it.padded:
                pad = "＋2 页"
            elif it.pad_front:
                pad = "＋1 页·首"
            elif it.padded:
                pad = "＋1 页"
            else:
                pad = "—" if not it.pages else "不补"
            tags = [it.status] if it.status in STATUS_COLORS else []
            if k % 2:
                tags.append("odd")
            self.tree.insert(
                "", "end",
                values=(it.index, it.src.name,
                        (it.src.parent.name if show_dir else ""),
                        it.fmt_label, pages, pad, it.status, it.note),
                tags=tuple(tags))

    def _refresh_merge_hint(self) -> None:
        """打印方式右边那句人话 + 输出路径的可用性，随口径变。"""
        if not hasattr(self, "lbl_print_tip"):
            return
        short = {
            M.PrintMode.DUPLEX.value: "奇数页的那份末尾补一张空白，每份都从正面开始",
            M.PrintMode.SIMPLEX.value: "直拼，一页都不补",
            M.PrintMode.HOME_SIMPLEX.value: "第 1 页后插一张空白，其余照双面走",
        }
        self.lbl_print_tip.config(
            text=short.get(self.vars["print_mode"].get(), ""))
        self._sync_run_button()

    def _refresh_orient_hint(self) -> None:
        """页面方向右边那句人话。

        ⚠️ 「内容会侧躺」必须写出来 —— 用户看到横版那几页文字歪着，
           第一反应是程序坏了，而不是「我开的方向开关的效果」。
        """
        if not hasattr(self, "lbl_orient_tip"):
            return
        short = {
            M.PageOrient.KEEP.value: "一张都不转，横竖混着就混着",
            M.PageOrient.PORTRAIT.value: "横版页转 90°，整批竖着印（内容会侧躺）",
            M.PageOrient.LANDSCAPE.value: "竖版页转 90°，整批横着印（内容会侧躺）",
        }
        self.lbl_orient_tip.config(
            text=short.get(self.vars["page_orient"].get(), ""))

    def on_pick_merge_out(self) -> None:
        """选合并输出到哪一份文件 —— 默认给桌面上的《合集打印.pdf》。"""
        cur = self.vars["merge_out"].get().strip()
        start = Path(cur).expanduser() if cur else Path.home() / "Desktop"
        got = filedialog.asksaveasfilename(
            title="合并输出到…", defaultextension=".pdf",
            initialdir=str(start if start.is_dir() else start.parent),
            initialfile=(start.name if start.suffix.lower() == ".pdf"
                         else M.DEFAULT_OUT_NAME),
            filetypes=[("PDF 文件", "*.pdf")])
        if got:
            self.vars["merge_out"].set(got)
            self._sync_run_button()

    def on_merge(self) -> None:
        """开始合并 —— 覆盖确认 → 后台线程 → 完事回填。

        ⚠️ 进度**不经过 `_jobs` 队列**（一次合并几十份会有几十条进度），
           而是写进 `self.merge_progress` 这个元组、由主线程定时读（见 `_pump_merge`）。
           这样后台线程一个 Tk 调用都不做 —— macOS 上跨线程碰 Tk 会直接炸。
        """
        if self.busy or self.mode != "merge":
            return
        plan = self.merge_plan
        if plan is None:
            self._refresh_merge_plan()
            plan = self.merge_plan
        if plan is None or plan.blocked:
            messagebox.showwarning(APP_TITLE,
                                   (plan.block_reason if plan else "") or "还不能合并。")
            return
        n = self._merge_count()
        if not n:
            messagebox.showwarning(
                APP_TITLE,
                "队列里没有要拼的文件。\n\n"
                "点右边的「＋ 添加文件…」把 " + MERGE_KINDS + " 挑进来；"
                "每一行也可以在右键里「忽略此行」。")
            return

        raw = self.vars["merge_out"].get().strip()
        if not raw:
            messagebox.showwarning(APP_TITLE, "还没选合并输出到哪一份文件。")
            return
        out = Path(raw).expanduser()
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")
            self.vars["merge_out"].set(str(out))
        mode = self.vars["print_mode"].get()
        orient = self.vars["page_orient"].get()

        head = (f"《{out.name}》已经存在，会被**覆盖**。" if out.exists()
                else f"将把 {n} 份文件合成一份 PDF：")
        if not messagebox.askokcancel(
                APP_TITLE,
                f"{head}\n{out}\n\n"
                f"打印方式：{M.print_mode_words(mode)}。\n"
                f"页面方向：{M.orient_words(orient)}。\n"
                "要转的先转成 PDF（写进临时目录，拼完就删）——\n"
                "**你的原件文件夹不会多出任何文件，原件一个字节都不动**。\n\n"
                "开始吗？"):
            self.status.config(text="已取消：没有生成合并文件。", foreground=DANGER)
            return

        try:
            out.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showwarning(APP_TITLE, f"这个位置写不进去：{exc}")
            return

        self._set_busy(True, "合并中……")
        self.pb["maximum"] = max(n + 1, 1)
        self.pb["value"] = 0
        self.merge_stop.clear()
        self.merge_progress = (0, max(n + 1, 1), "准备中…")
        self.status.config(text=f"正在合并（{n} 份）…… 中途别关窗口。", foreground=TEXT)
        if hasattr(self, "btn_merge_stop"):
            self.btn_merge_stop.config(state="normal")
        self._pump_merge()

        opts = self._collect_options()

        def work() -> None:
            try:
                res = M.run_merge(
                    plan, opts, out, mode=mode, orient=orient,
                    should_stop=self.merge_stop.is_set,
                    progress=lambda d, t, m: setattr(self, "merge_progress", (d, t, m)))
            except Exception as exc:  # noqa: BLE001 —— 线程死掉 = 界面永久卡忙
                res = M.MergeResult(message=f"合并时出错：{exc}", plan=plan)
            self._jobs.append(("merge", res))

        threading.Thread(target=work, daemon=True).start()
        self._kick_poll()

    def _pump_merge(self) -> None:
        """主线程定时把后台写的进度搬到界面上（见 on_merge 的说明）。"""
        if not self.busy or self.mode != "merge":
            return
        done, total, msg = self.merge_progress
        if total:
            self.pb["maximum"] = total
            self.pb["value"] = done
        if msg:
            self.status.config(text=msg, foreground=TEXT)
        self._merge_pump = self.after(120, self._pump_merge)

    def _after_merge(self, res: M.MergeResult) -> None:
        """合并跑完 —— 回填表格、报结果、把按钮和进度条复位。"""
        if getattr(self, "_merge_pump", None):
            try:
                self.after_cancel(self._merge_pump)
            except Exception:  # noqa: BLE001
                pass
            self._merge_pump = None
        self.merge_result = res
        self._set_busy(False)
        if hasattr(self, "btn_merge_stop"):
            self.btn_merge_stop.config(state="disabled")
        self._fill_table()
        self._sync_run_button()
        if res.ok:
            self.pb["maximum"] = max(self.pb["maximum"], 1)
            self.pb["value"] = self.pb["maximum"]
            self.status.config(text="✓ " + res.message, foreground=OK_GREEN)
            messagebox.showinfo(
                APP_TITLE,
                res.message + "\n\n" + str(res.out) +
                "\n\n原件一个字节都没动；中间 PDF 已经全部清掉了。")
            return
        self.pb["value"] = 0
        self.status.config(text="✗ " + res.message, foreground=DANGER)
        # 失败必须把原因摆到脸上（多半是转不出来 / 文件被占用 / 这份 PDF 读不了）。
        bad = [it for it in (res.plan.items if res.plan else [])
               if it.status == C.ConvStatus.FAILED.value]
        more = ""
        if len(bad) > 1:
            more = f"\n\n（另有 {len(bad) - 1} 份也失败，原因见表格「说明」列）"
        messagebox.showwarning(APP_TITLE, res.message + more)

    def on_merge_stop(self) -> None:
        """停止合并：引擎那一批转换打断不了，但补白 / 拼接与后续都会停下。"""
        if not self.busy or self.mode != "merge":
            return
        self.merge_stop.set()
        self.status.config(text="正在停止……（当前这一批转换跑完就停）", foreground=WARN)

    def on_preview(self) -> None:
        """列队（没选扫描件）或配对（选了）。

        两段合成一个按钮：`plan_batch` 自己看有没有扫描件决定做哪一段。
        """
        if self.busy or self.mode != "swap":
            return
        if not self.sources:
            messagebox.showwarning(APP_TITLE,
                                   "先把要处理的 PDF 加进「原件来源」——\n"
                                   "可以加一整个文件夹，也可以挑单个文件。")
            return
        self._start_swap_plan(auto=False)

    def _start_swap_plan(self, auto: bool) -> None:
        """跑一次 `plan_batch`，结果回到 `_after_preview`。

        `auto=True` 是「刚选完扫描件、自动把配对补上」（见 `_on_scan_change`）；
        `False` 是用户自己点了「预览」。两条路走的是**同一份代码**。

        ⚠️ 自动那条**不清表**：旧行先留在屏幕上，新计划一到整体换掉。
           先清表会闪一下白 —— 那正是用户说的「底下的原数据消失了」。
        """
        self._cancel_scan_after()     # 手点「预览」时，待触发的自动补配对就没必要了
        # 换了一份扫描件 → 上次人工判的方向作废（那批纸已经不在了）
        scan_now = self.vars["scan"].get().strip()
        if scan_now != self._rot_scan:
            self.rotations = {}
            self._rot_scan = scan_now
        self._preview_auto = auto
        opts = self._collect_options()
        self._set_busy(True, "正在读取……")
        self.status.config(text="正在读原件和扫描件……" if opts.scan_pdf else "正在读原件……",
                           foreground=TEXT)

        def work() -> None:
            try:
                plan = C.plan_batch(opts)
            except Exception as exc:  # noqa: BLE001
                plan = C.Plan(blocked=True, block_reason=f"读取时出错：{exc}")
            self._jobs.append(("preview", plan))

        threading.Thread(target=work, daemon=True).start()
        self._kick_poll()

    # ── 「扫描件 PDF」一变 ────────────────────────────────────────────────
    # ⭐ 2026-10-08 用户要求改的行为。原做法是 `_invalidate("扫描件换了")` ——
    #    清空表格 + 状态栏写「之前那次预览作废了，点预览重新来一遍」。
    #    用户原话：「那就显得像是之前的原件排序感觉有问题一样，容易重新出现排序」。
    #    真实流程是「先加原件 → 预览列队 → 照着扫 → 回来选扫描件」，
    #    这一步**只该把右边那几列（扫描页 / 尺寸 / 方向 / 状态）补上** ——
    #    原件那半边一个字都没变，没必要推倒重来。
    def _on_scan_change(self) -> None:
        """扫描件那一格变了 —— 先记一笔，停手一会儿再动手。

        ⚠️ 那一格是绑 StringVar 的**可输入** Entry，手打路径是**逐字符**触发的；
           不防抖的话每敲一个字母就重跑一遍配对（几十份原件要重新读一遍）。
        """
        self._cancel_scan_after()
        self._scan_after = self.after(SCAN_DEBOUNCE_MS, self._apply_scan_change)

    def _cancel_scan_after(self) -> None:
        if self._scan_after:
            try:
                self.after_cancel(self._scan_after)
            except Exception:  # noqa: BLE001
                pass
            self._scan_after = None

    def _apply_scan_change(self) -> None:
        """防抖到点了 —— 把配对补上（不是作废）。"""
        self._scan_after = None
        if self.busy or self.mode != "swap":
            return
        if self.plan is None or self.plan.blocked:
            # 还没预览过、或上一批被拦下来 —— 没有可补的东西，等用户点「预览」
            return
        self._start_swap_plan(auto=True)

    def _after_preview(self, plan: C.Plan) -> None:
        self._set_busy(False)
        self.plan = plan
        self._fill_table()
        self._refresh_srcdir_label()
        self._sync_run_button()
        if plan.blocked:
            self.status.config(text="来源/文件有问题，整批不执行（看弹出提示）",
                               foreground=DANGER)
            messagebox.showwarning(APP_TITLE, plan.block_reason)
            return
        skip = f"已自动跳过 {len(plan.products)} 份上一轮的产物；" if plan.products else ""
        if not plan.paired:
            # 第一段：只列了队。这就是「先看顺序、照着去扫」的那一步。
            self.status.config(
                text=f"已按文件名排好 {len(plan.items)} 份（{skip}还没选扫描件）。"
                     f"照着表格从上到下的顺序摊开纸质件去扫；"
                     f"扫完回来把「扫描件 PDF」选上，配对会**自动接上**。",
                foreground=OK_GREEN)
            return
        # ⭐ 自动补的那一次，先把「原件那半边没动」说出来 —— 免得看着像顺序被重排了
        #    （这正是当初把「作废重来」改成「自动补上」的原因）。
        if self._preview_auto:
            lead = (f"扫描件接上了，原件顺序没变 —— 原件 {len(plan.items)} 份、"
                    f"扫描件 {plan.scan_pages} 页，一一对应上了。")
        else:
            lead = (f"原件 {len(plan.items)} 份、扫描件 {plan.scan_pages} 页，"
                    f"一一对应上了。")
        if self._suffix_ready():
            self.status.config(
                text=lead + f"（{skip}产物会写在各自原件旁边）确认表格没错再点「开始替换」。",
                foreground=OK_GREEN)
        else:
            self.status.config(
                text=lead + "但「输出文件名后缀」还没填 —— 填上才让跑。",
                foreground=DANGER)

    def _ask_overwrite(self, n: int) -> bool | None:
        """产物会撞上同名文件时问一句。

        返回 True = 覆盖它们 / False = 不覆盖（自动加编号）/ None = 先别跑。

        用户两点口径（2026-10-08）：
          · 「打勾的话又太麻烦了……只有在这个界面点击确定了之后才会直接覆盖，
            否则就不覆盖」→ 界面上**不设勾选框**，只留这一个弹窗；
          · 「就不要说哪几份文件重名了……以后处理下来可能是十几二十份，太多了，
            你列不完，列出来也没有意义」→ **不列文件名清单**，只报个数。
        所以这里只有一句「有 N 份重名」+ 是 / 否 / 取消 三个选择。

        ⚠️ 原件是只读的，被覆盖的只会是「原件名 + 后缀」那些产物名，跟原件无关。
        ⚠️ 两个模式的**后果不同**（替换模式会盖掉上一轮的签字版；转换模式只是重转
           一份 PDF），所以那句话按模式换一个说法。
        """
        if self.mode == "convert":
            head = f"有 {n} 份的产物名（同名 PDF）已经存在，通常是上一轮转出来的。"
        else:
            head = f"有 {n} 份文件的产物名已经存在（通常是上一轮生成的产物）。"
        # ⚠️ 三个选项**各占一行且要短**（2026-10-09 用户要求）：原来那句
        #    「不覆盖，给产物加编号（xxx_2.pdf）」会折成两行，看着乱。
        return messagebox.askyesnocancel(
            APP_TITLE,
            head + "\n\n"
            "是 ＝ 覆盖它\n"
            "否 ＝ 加编号另存\n"
            "取消 ＝ 先别跑\n\n"
            "原件本身不会被改动。")

    def on_run(self) -> None:
        if self.busy or self.mode != "swap" or self.plan is None or self.plan.blocked:
            return
        if not self.plan.paired:
            messagebox.showwarning(APP_TITLE,
                                   "还没配对扫描件 ——\n"
                                   "先选好「扫描件 PDF」、点一次「预览」，配对上了再替换。")
            return
        opts = self._collect_options()
        why = C.suffix_error(opts)
        if why:
            messagebox.showwarning(APP_TITLE, why)
            self._sync_run_button()
            return
        n = len([it for it in self.plan.items if it.status != C.Status.SKIPPED.value])
        if not n:
            messagebox.showwarning(APP_TITLE, "没有要处理的文件。")
            return

        # ⑨ 覆盖必须让人看着清单点头（没有勾选框了）。
        # ⚠️ 回答要落成**三态**：是＝一律覆盖；否＝一律加编号让开（连自家上一轮的
        #    产物也不碰）—— 只写 bool 的话「否」会被 core 的 AUTO 分支吃掉。
        will = C.overwrite_list(self.plan, opts)
        if will:
            ans = self._ask_overwrite(len(will))
            if ans is None:
                self.status.config(text="已取消：一个文件都没动", foreground=DANGER)
                return
            opts.overwrite = C.OVERWRITE_YES if ans else C.OVERWRITE_NO
        elif not messagebox.askokcancel(
                APP_TITLE,
                f"将生成 {n} 份 PDF，写在各自原件旁边。\n\n原件不会被改动。开始吗？"):
            return
        self._set_busy(True, "替换中……")
        self.pb["maximum"] = max(n, 1)
        self.pb["value"] = 0

        def work() -> None:
            results, _notes = C.run_batch(self.plan, opts)
            self._jobs.append(("run", results))

        threading.Thread(target=work, daemon=True).start()
        self._kick_poll()

    # ── 转换模式（Office 转 PDF）───────────────────────────────────────────
    #
    # ⭐ 跟替换那套**共用**来源清单、后缀、覆盖三态，但流程短一截：
    #     加文件 → （列表自动列出来）→ 点「开始转换」→ 完事。
    #   没有「预览」这一步 —— 转换只认扩展名，加进来就已经是队列了。
    #   转换本身可能一份要好几秒（Office 起进程），所以执行**必须**走后台线程。
    @staticmethod
    def _probe_engines() -> list[C.ConvEngine]:
        """本机**每一族选中的那个**引擎（一族一条，见 core 的 ⑩a）。

        ⭐ 为什么不只挑「一个全局首选」：这台机器上能转的东西本来就不止一样 ——
           Windows 上 Word / Excel 走 Office，CAD 走浩辰 CAD（⑩d 品牌表），
           它们各走各的路，谁也不是谁的备胎。
        ⭐ 一族一个都挑不出来（Mac 上的 CAD 就是这样）→ 那一族**不列**，
           真正要用到它时由列队闸门与队列本身说清楚原因。
        """
        out: list[C.ConvEngine] = []
        for fam in C.ENGINE_FAMILIES:
            eng = C.pick_engine(fam)
            if eng is not None:
                out.append(eng)
        return out

    def _shown_engines(self) -> list[C.ConvEngine]:
        """那行小灰字显示哪些引擎。

        ⭐ 队列里已经有东西 → 显示**这一批真会用到的**（一族一条）；
           队列还是空的 → 显示这台机器**能用上的全部**（回答「我这台能转什么」）。
        """
        plan = self.conv_plan
        if plan is not None and plan.engines:
            return list(plan.engines)
        if self.engines is None:
            self.engines = self._probe_engines()
        return self.engines

    def _refresh_engine_label(self) -> None:
        """转换模式下操作行尾那行小灰字：现在能用哪些引擎。

        ⚠️ 探测是**带缓存**的（`self.engines` 不为 None 就不再问）——
           探针要摸文件系统 / 试 import，每刷一次界面就问一遍没必要。
           刚装完某个程序不会自己发现，所以另配一个「重新检测引擎」按钮。
        """
        if caps_of(self.mode).queue == "plan":
            return                     # 替换模式没有引擎这一说
        shown = self._shown_engines()
        if not shown:
            self.lbl_engine.config(text="⚠ 没找到可用的转换引擎（点「重新检测引擎」）",
                                   foreground=DANGER)
        else:
            self.lbl_engine.config(text="引擎：" + "、".join(e.label for e in shown),
                                   foreground=OK_GREEN)

    def on_recheck_engine(self) -> None:
        """丢掉缓存重探一次（用户刚装完某个程序时点）。

        ⚠️ 光重探不够：core 里 CAD 的「注册表里有没有这个名字」是**进程级缓存**的
           （见 core 的 find_cad），不清掉的话，刚装上的浩辰 CAD 照样看不见。
        """
        if self.busy:
            return
        C.reset_cad_cache()                   # 让「刚装的 CAD」这一次能被看见
        self.engines = self._probe_engines()
        if not self.engines:
            messagebox.showwarning(APP_TITLE, C.engine_no_engine_msg())
        # 引擎一变，闸门与「开始」的可用性都得重算（哪个模式在台上就重算哪个）
        q = caps_of(self.mode).queue
        if q == "conv_plan":
            self._refresh_convert_plan()
        elif q == "merge_plan":
            self._refresh_merge_plan()
        if self.engines:
            self.status.config(text="找到可用的转换引擎："
                                    + "、".join(e.label for e in self.engines),
                               foreground=OK_GREEN)

    def _refresh_convert_plan(self) -> None:
        """把来源清单算成队列并铺到表格上。

        ⚠️ **同步**跑（不开线程）：只认扩展名、不读文件内容，几十份也是毫秒级。
           开线程反而带来「表格还没填好按钮已经能点」的竞态。
        ⚠️ 被闸门拦住时**不弹窗** —— 每次加文件都弹一次太吵；状态栏说一句就够，
           真要跑时 `on_convert` 还会再弹一次完整的。
        """
        if self.mode != "convert":
            return
        opts = self._collect_options()
        try:
            plan = C.plan_convert(opts)
        except Exception as exc:  # noqa: BLE001
            plan = C.ConvertPlan(blocked=True, block_reason=f"列队时出错：{exc}")
        self.conv_plan = plan
        self.conv_results = {}
        # ⭐ 队列里用到的引擎顺带回填进缓存（省一次探测）；「重新检测引擎」仍可重探。
        if plan.engines and self.engines is None:
            self.engines = list(plan.engines)
        self._fill_table()
        self._refresh_srcdir_label()
        self._refresh_engine_label()
        self._sync_run_button()
        if plan.blocked:
            first = (plan.block_reason or "还不能转").splitlines()[0]
            self.status.config(text=first, foreground=WARN)
            return
        unk = len(plan.unknown)
        tail = f"，另有 {unk} 份认不出格式、已跳过" if unk else ""
        self.status.config(
            text=f"队列里有 {self._convertible_count()} 份待转{tail}。"
                 f"看一眼「产物名」再点「开始转换」。",
            foreground=OK_GREEN)

    def on_convert(self) -> None:
        """开始转换 —— 覆盖确认 → 后台线程 → 完事回填。"""
        if self.busy or self.mode != "convert":
            return
        plan = self.conv_plan
        if plan is None:
            self._refresh_convert_plan()
            plan = self.conv_plan
        if plan is None or plan.blocked:
            messagebox.showwarning(APP_TITLE, (plan.block_reason if plan else "")
                                   or "还不能转换。")
            return
        n = self._convertible_count()
        if not n:
            messagebox.showwarning(APP_TITLE,
                                   "队列里没有要转的文件。\n\n"
                                   "点左边的「＋ 添加文件…」把 " + CONVERT_KINDS + " 挑进来；"
                                   "每一行也可以在右键里「忽略此行」。")
            return
        opts = self._collect_options()

        # ⑨ 覆盖跟替换那边同一套三态（勾选框早就删了，要用户点一次头）。
        # ⚠️ 回答要落成**三态**：是＝一律覆盖；否＝一律加编号让开。
        will = C.convert_overwrite_list(plan, opts)
        if will:
            ans = self._ask_overwrite(len(will))
            if ans is None:
                self.status.config(text="已取消：一份都没转", foreground=DANGER)
                return
            opts.overwrite = C.OVERWRITE_YES if ans else C.OVERWRITE_NO
        elif not messagebox.askokcancel(
                APP_TITLE,
                f"将把 {n} 份 {CONVERT_KINDS} 转成 PDF，PDF 写在各自原件旁边。\n\n"
                f"文件名：原件名{C.clean_output_suffix(opts.suffix)}.pdf\n"
                "原件本身不会被改动。开始吗？"):
            return

        self._set_busy(True, "转换中……")
        self.pb["maximum"] = max(n, 1)
        self.pb["value"] = 0
        _use = "、".join(e.label for e in plan.engines) if plan.engines else "引擎"
        self.status.config(text=f"正在转换（{n} 份）…… 用 {_use}，"
                                f"同格式合成一批、只拉起一次它，中途别关窗口。", foreground=TEXT)

        def work() -> None:
            try:
                # ⚠️ **不要**把引擎注入进来（2026-10-09 改）：注入 = 钦定唯一一个，
                #    兜底梯队就被关掉了。交给 core 按「族」自己挑、自己兜底 ——
                #    这样 Office 失手时 LibreOffice 才真的接得上（旧写法下那段
                #    「已自动改用 XX 转好」的提示其实是死代码）。
                results, _notes = C.run_convert(plan, opts)
            except Exception as exc:  # noqa: BLE001
                results = [C.ConvertResult(src=it.src, ok=False,
                                           status=C.ConvStatus.FAILED.value,
                                           note=f"转换时出错：{exc}")
                           for it in plan.items]
            self._jobs.append(("convert", results))

        threading.Thread(target=work, daemon=True).start()
        self._kick_poll()

    def _after_convert(self, results: list[C.ConvertResult]) -> None:
        self._set_busy(False)
        self.pb["value"] = self.pb["maximum"]
        ok = [r for r in results if r.status == C.ConvStatus.OK.value]
        bad = [r for r in results if r.status == C.ConvStatus.FAILED.value]
        skip = [r for r in results if r.status == C.ConvStatus.SKIPPED.value]
        if self.conv_plan is not None:
            by_src = {r.src: r for r in results}
            for it in self.conv_plan.items:
                r = by_src.get(it.src)
                if r is not None:
                    it.status = r.status
                    it.note = r.note
                    self.conv_results[it.src] = r
            self._fill_table()
        self._sync_run_button()
        msg = f"转好 {len(ok)} 份"
        if skip:
            msg += f"，跳过 {len(skip)} 份"
        if bad:
            msg += f"，失败 {len(bad)} 份"
        tail = "　→　PDF 写在各自原件旁边"
        copies = sum(1 for r in ok if r.via_copy)
        if copies:
            # 走临时副本的（⑩b：来自网络标记 / 正被别的程序开着）—— 一句话带过，不弹窗。
            # 用户最关心的是「原件动没动」，所以那句必须留着。
            tail += f"（{copies} 份已用临时副本转，原件没动）"
        self.status.config(text=msg + tail,
                           foreground=DANGER if bad else OK_GREEN)
        # ⭐ 兜底梯队接手过 → 让用户知道「首选引擎在你这台机器上不灵」。
        #    不说的话，他只会看到「转好了」，下次换个文件又莫名其妙失败。
        planned = {e.key for e in self.conv_plan.engines} if self.conv_plan else set()
        fb = [r for r in ok if r.engine_key and planned and r.engine_key not in planned]
        if bad:
            # 失败原因**必须**摆到脸上来（多半是授权 / 没装引擎 / 文件被占用），
            # 只说「失败 3 份」用户不知道下一步该干嘛。
            head = bad[0]
            more = f"\n（另有 {len(bad) - 1} 份也失败，原因见表格「说明」列）" if len(bad) > 1 else ""
            # ⭐ 环境类失败（找不到 CAD / PDF 打印机 / 引擎）→ 引导去重新检测引擎。
            #    别的（文件被占用 / 授权）跑了也没用，不引导、免得误导。
            _note = head.note or ""
            _env_fail = ("CAD" in _note or "打印机" in _note or "引擎" in _note
                         or "没找到" in _note or "没装" in _note or "找不到" in _note)
            guide = ""
            if _env_fail:
                guide = "\n\n点界面上的「重新检测引擎」再试一次；还不行就看说明书「转换失败」那一段。"
            # ⭐ 出错写一份报告（用户点名要：失败原因 + 机器环境，方便自查或发人）。
            report_path = self._write_error_report(bad)
            if report_path is not None:
                guide += f"\n\n失败详情已写到：\n{report_path}"
            messagebox.showwarning(
                APP_TITLE,
                f"{msg}。\n\n第一份失败的原因：\n{head.src.name}\n{head.note}{guide}{more}")
        else:
            extra = ""
            if fb:
                names = "、".join(sorted({self._engine_label(r.engine_key) for r in fb}))
                extra = (f"\n\n提醒：有 {len(fb)} 份**不是首选引擎转的** ——\n"
                         f"首选那个没转成，已自动改用「{names}」接手。\n"
                         f"每份的具体原因见表格「说明」列。")
            messagebox.showinfo(
                APP_TITLE,
                msg + "\n\nPDF 就写在各自原件旁边，原件一个字节都没动。" + extra)
        # ⚠️ 这里**只到「弹一次确定」为止**，不再追加任何「是 / 否」询问。
        #    2026-10-09 用户拍板：使用的人会被 Yes/No 问糊涂（以为要装什么软件），
        #    而且公司电脑不让随便装东西。失败原因已经写在表格「说明」列里，
        #    够用了。**别再加回来**（回归 [13i] 盯着这条）。

    @staticmethod
    def _engine_label(key: str) -> str:
        e = C.engine_by_key(key)
        return e.label if e else key

    @staticmethod
    def _write_error_report(bad: list[C.ConvertResult]) -> Path | None:
        """转换失败时写一份错误报告（用户 2026-10-10 点名要：失败原因 + 机器环境）。

        写到**第一份失败文件所在目录**（跟产物同目录，用户自然能找到、不污染桌面）。
        写失败（无权限等）返回 None，别让「写报告」这个动作再炸一次转换。
        """
        if not bad:
            return None
        try:
            from datetime import datetime
            first = bad[0]
            out_dir = first.src.parent
            lines = [
                f"{C.TOOL_NAME} v{C.TOOL_VERSION} —— 转换失败报告",
                "=" * 40,
                f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"平台：{C.running_platform()}",
                f"失败：{len(bad)} 份",
                "",
                "── 失败清单 ──",
            ]
            for r in bad:
                lines.append(f"· {r.src.name}")
                lines.append(f"    原因：{r.note or '(未给原因)'}")
                if r.engine_key:
                    lines.append(f"    引擎：{r.engine_key}")
            lines.append("")
            lines.append("── 提示 ──")
            lines.append("若失败原因是「找不到 CAD / PDF 打印机 / 引擎」，")
            lines.append("点界面上的「重新检测引擎」再试一次。")
            report = out_dir / "错误报告.txt"
            report.write_text("\n".join(lines), encoding="utf-8")
            return report
        except Exception:                                     # noqa: BLE001
            return None

    def _kick_poll(self) -> None:
        """确保轮询链活着 —— 后台线程刚起来时调一次。

        ⚠️ 只有 `_poll_armed` 为假（链已经停了）才真的开一条；
           否则调了也是白调，还会多出一条抢同一个队列的链。
        """
        if not self._poll_armed:
            self._poll()

    def _poll(self) -> None:
        """主线程定时取后台结果 —— 后台线程绝不直接碰 Tk（macOS 上会炸
        “main thread is not in main loop”）。

        ⚠️ 全应用**只有一条轮询链**：`_poll_armed` 表示「已经排上了一次」。
           这里进来先把它清成 False（说明这一跳正在跑），做完再排下一跳 ——
           所以链子只会在「没人再排」时才真正停下。
        """
        self._poll_armed = False
        if not self._jobs:
            self._poll_armed = True
            self.after(50, self._poll)
            return
        kind, payload = self._jobs.pop(0)
        if kind == "preview":
            self._after_preview(payload)
        elif kind == "convert":
            self._after_convert(payload)
        elif kind == "merge":
            self._after_merge(payload)
        else:
            self._after_run(payload)
        # ⚠️ 处理完之后**必须再排一次**：原来这里没有，只要队列里攒了第二件事
        #    （比如合并失败后紧接着再跑一次），轮询链就断了，界面永远收不到结果。
        self._poll_armed = True
        self.after(50, self._poll)

    def _after_run(self, results: list[C.SwapResult]) -> None:
        self._set_busy(False)
        self.pb["value"] = self.pb["maximum"]
        ok = sum(1 for r in results if r.status == C.Status.OK.value)
        bad = [r for r in results if r.status == C.Status.FAILED.value]
        skip = sum(1 for r in results if r.status == C.Status.SKIPPED.value)
        # 把状态写回表格，用户能直接看到哪一份失败
        if self.plan is not None:
            by_name = {r.src.name: r for r in results}
            for it in self.plan.items:
                r = by_name.get(it.src.name)
                if r is not None:
                    it.status = r.status
                    it.note = r.note
            self._fill_table()
        msg = f"完成 {ok} 份"
        if skip:
            msg += f"，跳过 {skip} 份"
        if bad:
            msg += f"，失败 {len(bad)} 份（表格「说明」列有原因）"
        self.status.config(text=msg + "　→　产物写在各自原件旁边",
                           foreground=DANGER if bad else OK_GREEN)
        messagebox.showinfo(APP_TITLE,
                            msg + "\n\n产物就写在各自原件旁边（文件名 = 原名 + 后缀）。")

    def _set_busy(self, busy: bool, title: str = "") -> None:
        self.busy = busy
        self.btn_scan.config(state="disabled" if busy else "normal")
        self.btn_recheck.config(state="disabled" if busy else "normal")
        # 忙的时候把「扫描件 PDF」那一格也锁上 —— 半路改路径会触发一次自动补配对，
        # 而那次会被 busy 挡掉，等于把用户刚填的东西白吞了。
        row = self.rows.get("scan")
        if row is not None:
            st = "disabled" if busy else "normal"
            row.entry.config(state=st)
            row.button.config(state=st)
        self._sync_run_button()          # 忙 / 闲、有没有后缀，都在这一个地方判
        # 「停止」只在真跑合并时能点：开工由 on_merge 打开，这里负责**收回去**
        # （收不回去的话，干完之后还亮着一颗能点的停止键）。
        if hasattr(self, "btn_merge_stop") and not busy:
            self.btn_merge_stop.config(state="disabled")
        if busy:
            for b in self._adjust_buttons():
                b.config(state="disabled")
        else:
            self._sync_adjust_buttons()
        # 干活时标题栏报进度，干完**必须**改回来 —— 否则「替换中……」会一直挂在
        # 标题栏上，用户以为还在跑（转换一份要几秒，更容易误会）。
        self.winfo_toplevel().title(f"{APP_TITLE}　——　{title}" if busy else APP_TITLE)

    # ── 右键菜单 ─────────────────────────────────────────────────────────
    def _current_item(self):
        """当前模式表格里选中的那一行。

        替换模式给 `C.Item`，转换模式给 `C.ConvItem`，合并模式给 `M.MergeItem` ——
        三者都带 src / status / note，所以「忽略此行」和「定位原件」两个动作通用。
        """
        sel = self.tree.selection()
        if not sel:
            return None
        plan = self._current_plan()
        if plan is None or plan.blocked:
            return None
        i = self.tree.index(sel[0])
        return plan.items[i] if 0 <= i < len(plan.items) else None

    def _popup(self, event: tk.Event) -> None:
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        plan = self._current_plan()
        if plan is None or plan.blocked:
            return
        self.tree.selection_set(iid)
        # 各模式各有一份菜单（替换那份才有「改用扫描件第几页」）
        menu = getattr(self, caps_of(self.mode).menu)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def on_row_double(self, _e: tk.Event | None = None) -> None:
        """双击：替换模式＝直接指定用第几页；其余模式＝在访达里定位原件。"""
        if caps_of(self.mode).key == "swap":
            self.on_assign_page()
        else:
            self.on_reveal()

    def on_assign_page(self) -> None:
        """双击行 / 右键：直接指定这一行用扫描件第几页。

        ⭐ 一页只给一份原件：填的页已被别的行占用时自动**对调**，
           不会让两行抢同一页（那样必定有一份原件拿不到封面）。
        ⚠️ 只有替换模式有「扫描页」这一说，转换模式直接不进门。
        """
        if self.mode != "swap":
            return
        it = self._current_item()
        if it is None or self.plan is None:
            return
        idx = self.tree.index(self.tree.selection()[0]) + 1
        total = self.plan.scan_pages
        win = tk.Toplevel(self)
        win.title("改用扫描件第几页")
        win.configure(bg=BG)
        win.transient(self.winfo_toplevel())
        win.resizable(False, False)
        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=12)
        tk.Label(body, text=f"《{it.src.name}》", bg=BG, fg=TEXT,
                 font=(FONT, 11, "bold"), anchor="w").pack(anchor="w")
        tk.Label(body, text=f"现在用的是扫描件第 {it.scan_index} 页。", bg=BG,
                 fg=TEXT2, font=(FONT, 10), anchor="w").pack(anchor="w", pady=(2, 8))
        var = tk.IntVar(value=it.scan_index)
        ttk.Spinbox(body, from_=1, to=max(total, 1), textvariable=var,
                    width=8).pack(anchor="w")
        tk.Label(body, text=f"扫描件共 {total} 页（一页只给一份原件，占用的自动对调）。",
                 bg=BG, fg=MUTED, font=(FONT, 9), justify="left",
                 anchor="w").pack(anchor="w", pady=(8, 0))

        def ok() -> None:
            changed, note = C.assign_page(self.plan, idx, int(var.get()))
            win.destroy()
            if not changed:
                return
            self._apply_rotations()
            self._reselect_refresh()
            self.status.config(text=f"《{it.src.name}》{note}。看一眼表格再跑。",
                               foreground=TEXT)

        btns = tk.Frame(body, bg=BG)
        btns.pack(anchor="e", pady=(12, 0))
        # 按钮顺序按 macOS 习惯：左「取消」、右主操作；**都不设焦点**，
        # 焦点留在窗口本身（设了之后空格/回车会把焦点按钮按掉）。
        ttk.Button(btns, text="取消", style="Small.TButton",
                   command=win.destroy).pack(side="left")
        ttk.Button(btns, text="确定", style="Primary.TButton",
                   command=ok).pack(side="left", padx=(6, 0))
        win.focus_set()

    def on_ignore(self) -> None:
        """把选中的这一行标成「跳过 / 恢复」—— 两个模式同一套逻辑，只是状态字不同。"""
        it = self._current_item()
        if it is None:
            return
        if caps_of(self.mode).key == "swap":
            skipped, pending = C.Status.SKIPPED.value, C.Status.PENDING.value
        else:
            # 转换与合并共用同一套状态词（merge_core 直接沿用 ConvStatus）
            skipped, pending = C.ConvStatus.SKIPPED.value, C.ConvStatus.PENDING.value
        it.status = pending if it.status == skipped else skipped
        it.note = "你手动忽略了这一行" if it.status == skipped else ""
        self._fill_table()
        self._sync_run_button()

    def on_reveal(self) -> None:
        """在访达 / 资源管理器里定位这一行的原件（两个模式通用）。"""
        it = self._current_item()
        if it is None:
            return
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", "-R", str(it.src)], check=False)
            elif os.name == "nt":
                subprocess.run(["explorer", "/select,", str(it.src)], check=False)
            else:
                subprocess.run(["xdg-open", str(it.src.parent)], check=False)
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning(APP_TITLE, f"打不开：{exc}")

    # ── 其它 ─────────────────────────────────────────────────────────────
    def on_close(self) -> None:
        self._cancel_save()
        self._save_config()
        # 合并线程是 daemon，但让它早点看到停止信号 —— 免得关窗后还在后台转文件
        self.merge_stop.set()
        self.winfo_toplevel().destroy()


def main() -> int:
    root = tk.Tk()
    root.title(APP_TITLE)
    # 统一浅色系：整个应用只在这里定一次样式（见 apply_theme）
    apply_theme(root)
    root.configure(bg=BG)
    root.geometry(default_geometry(root.winfo_screenwidth(), root.winfo_screenheight()))
    # ⭐ 先给主窗口补上「接拖入」的能力（见 enable_dnd），再建界面 ——
    #    App 建好之后要往窗口和框架上注册 drop target（见 _setup_dnd）。
    #    ⚠️ 挂不上**不是错误**：只是少一个方便功能，界面照常起来，
    #       里面的提示也不会提「可以拖进来」。
    enable_dnd(root)
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
