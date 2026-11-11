#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF转换拼接工具 v6.2 —— 核心层（顺序配对，不做名字匹配）

⚠️ 6.0 只在这一层加了 `Options.out_dir`（合并打印的中间产物要落到临时目录），
   以及版本号。转换 / 首页替换的全部规则与 5.0 **一字未改** ——
   合并打印那条链路在 merge_core.py 里，它复用本文件的引擎梯队与闸门。

和 v1（PDF 首页替换工具）的区别
------------------------------
v1：一份原件 ↔ 一份扫描件，靠文件名相似度匹配。
v2：现实里不可能一份一份扫 —— 把十几二十份签字首页**一口气扫成一份多页 PDF**，
    原件**按文件名排队**（就是文件夹里按「名称」看到的那个顺序，不用改名成 001/002），
    然后**按顺序硬对应**：
        表格第 1 行 ← 扫描件第 1 页
        表格第 2 行 ← 扫描件第 2 页
        ……
    不做任何名字匹配（也没有相似度门槛）。
    顺序 = 你扫的顺序 = 你把纸质文件摊开的顺序，所以界面里能 ↑↓ 微调、也能双重指定页号。

⚠️ 顺序配对是一条「脆弱链」：中间任何一份原件出问题（读不出、页数不齐），
   后面全部会错位。所以校验做成**阻断式** —— 有问题就整批不输出，
   而不是"跳过那一份继续跑"（跳过 = 从第 N 份起全部贴错封面）。

规则都收在下面的表里，加规则只改表：
    ① 排序规则表   ORDER_RULES   —— 原件怎么排队
    ② 阻断闸门表   GATE_RULES    —— 什么情况下整批拒绝执行
    ③ 贴合决策表   FIT_*         —— 扫描页缩放到原件首页 vs 原样保留
    ④ 后缀规范化   SUFFIX_*      —— 界面与命令行同一口径（**必填**，见下）
    ⑤ 状态枚举     Status        —— 预览表格里能出现的全部状态
    ⑥ 方向表       ROT_*         —— 扫描页左转/右转（人工判，机器判不了）
    ⑦ 自动方向     AUTO_ROT_*    —— 原件首页是横向的 → 它那一页默认左转 90°
    ⑧ 产物识别     sibling_source_of / is_product —— 认出上一轮的产物
    ⑨ 落点决策     resolve_target —— 目标名已被占用时怎么办
    ⑩ 格式转换表   CONVERT_FORMATS —— 认哪些扩展名、用哪个应用转、在哪几个平台上有
    ⑩b 转前预处理  PREP_RULES      —— 原件不能动，但有时得换个方式喂（「来自网络」标记 / 正被占用）
    ⑩c 分页规则表  PAGE_RULES      —— 一份文件该出几页（Word 分页 / Excel 一表一页）
    ⑩d CAD 品牌表   CAD_APPS        —— 认哪家 CAD、用哪台打印机出图（靠 COM 名字找）
    ⑪ 转换引擎表   CONVERT_ENGINES —— 谁来执行转换；每行标了**平台**与**族**
    ⑫ 转换闸门     CONVERT_GATES   —— 转换模式下什么情况整批拒绝

⭐ 「族（family）」是什么（2026-10-09 用户问「引擎梯队」时收口的）
--------------------------------------------------------------
⑩ 表每行写它是**哪一族**的格式；⑪ 表每行写它**伺服哪一族**、**在哪个平台**上成立。
于是「这条格式该走哪条路」不需要任何 if：同族的引擎按表序排成梯队，兜底只在**同族内**换。
· Mac 版**永远看不到** Windows 的引擎（平台不符的行直接在表外）；
· CAD 图纸失手**不会**掉到 LibreOffice / Word 上去试（不同族）；
· 只在 Windows 上有的格式（CAD / BarTender）写在 `platforms=("win",)` —— Mac 上收得进队列，
  但会在列队时就说清「这台机器上没有能转它的程序」，而不是让某个应用白试一遍。

⭐ v3.0 加了第二种模式：**Office 转 PDF**（2026-10-08 用户提出）
--------------------------------------------------------------
真实流程里原件并不总是 PDF：源头常常是 Word / Excel，改完数据 → 重新上传 →
**打印** → 签字 → 扫描。打印需要 PDF，所以在本工具的「首页替换」之前还有一步
「把几十份 Word / Excel 批量转成 PDF」。

于是 core 里现在有两套**互不干扰**的流程，共用同一份 `Options`：

| | 首页替换（原） | 格式转换（新，⑩⑪⑫） |
|---|---|---|
| 认的输入 | `.pdf` | Word / Excel 的扩展名（见 ⑩） |
| 要不要扫描件 | 要（第二段） | **不要**，也用不到①②③ |
| 后缀 | **必填**（产物也是 .pdf，不填就盖原件） | **可以不填**（产物 .pdf vs 原件 .xlsx，扩展名不同） |
| 产物落点 | 每份原件旁边 | 每份原件旁边（同一口径） |
| 原件 | 只读 | 只读 |

⚠️ 两套流程**各自独立判定**（`plan_batch` / `plan_convert`），谁也不调用谁 ——
   「转换完的 PDF 再拿去替换」是使用者串起来的两步，不是代码里的串联。

⭐ v3.0 换了输入输出模型（2026-10-08 用户拍板，与 v1/v2 都不同）
--------------------------------------------------------------
**没有「输出文件夹」了**：产物直接写在**每份原件自己所在的目录**里
（原件在子文件夹里，产物也跟着落在那个子文件夹里，天然不会同名打架）。

代价与对策，三条都要记牢：
    1. 后缀**必填** —— 后缀为空时产物名 = 原件名，会直接覆盖原件。
       界面禁用按钮 + 命令行拒绝 + `run_batch` 再兜一道，见 `suffix_error()`。
    2. 第二次跑时上一轮的产物就躺在原件旁边，会被当成原件收进来 → 数量翻倍。
       所以有 ⑧：按「原件名 + 当前后缀」认出产物、跳过，并记进 `Plan.products`。
    3. 目标名已被占用时按 ⑨ 的表处理：**只覆盖自己生出来的产物**，
       认不出是谁的一律加编号让开（绝不会碰到原件本身）。

⭐ v3.0 又一处改动（2026-10-08 下午）：原件来源从「一个文件夹」改成**清单**
--------------------------------------------------------------------
`Options.sources` 是一串来源，每个可以是：
    · 一个**文件夹** —— 收它**这一层**里的 PDF（不再往子目录钻：「连同子文件夹
      一起处理」已删，见下）
    · 一个**PDF 文件** —— 你手动挑的
    · 混着来都行；同一个文件被指到两次会自动去重（既加了文件夹、又单加了里面的文件）

排队顺序 = **全部按文件名**（`DEFAULT_ORDER="name"`，见 ①）：
不看来源、不给来源分组，所以「先加哪个文件夹」不影响队列 ——
同一条队在任何机器上都是同一条。

为什么删掉递归（用户 2026-10-08 拍板）：
    它原本是 v1「防输出件落进子文件夹后、第二轮按名字匹配抓重」用的；
    v3.0 输出在原位、匹配靠顺序，这条用途作废。剩下的「一次收一批散在子目录
    里的原件」现在由**来源清单**显式表达 —— 谁参与、一目了然，不会误收无关文件。

预览是**两态**的（`Plan.paired`）：
    · 没选扫描件 → 只**列队**：把原件按名字排好、读页数/尺寸，`paired=False`；
    · 选了扫描件 → **配对**：再叠上「扫描件能否读 / 数量是否相等」，`paired=True`。
    界面据此先让你看队、照着摊纸质件去扫，扫完再回来配对。
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable, NamedTuple, Sequence

# 目录重构（2026-11-02）：本文件挪进 app/engine/，bootui 在 app/shell/。
# 直接运行本文件（命令行自测）也要能解析到 bootui，先铺好 sys.path。
_APP = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP not in sys.path:
    sys.path.insert(0, _APP)
import paths as _paths  # noqa: E402
_paths.install()

# ⭐ 启动/体检的统一版式（2026-10-23 抽出来的共用模块，各项目逐字节同一份）。
#    ⚠️ 只许 import 标准库的东西，别往 bootui 里塞项目逻辑 —— 它要被拷到每个包里。
import bootui as B

TOOL_NAME = "PDF转换拼接工具"
TOOL_VERSION = "6.2"

PDF_EXTS = frozenset({".pdf"})


# ══════════════════════════════════════════════════════════════════════════
# ① 排序规则表
# ══════════════════════════════════════════════════════════════════════════

_NUM_RE = re.compile(r"(\d+)")


def natural_key(name: str) -> tuple[tuple[int, int, str], ...]:
    """「文件名顺序」的排序键 —— 力求与文件夹里按「名称」看到的顺序一致。

    两条规则，缺一不可：
        ① 字母不区分大小写（Apple 与 apple 不会被大小写劈成两堆）
        ② 数字按数值大小排（WI-9 在 WI-10 前面，不会变成 WI-10 < WI-9 的字典序错误）

    每段统一成 (类别, 数值, 小写文本) 三元组，保证任意两个名字比较时
    类型永远一致 —— 否则 str 和 int 撞在一起会抛 TypeError。

    ⚠️ 已知的边界：中文名按 Unicode 码点排，可能与资源管理器/访达的拼音序不完全一致。
       真遇到就到表格里用 ↑↓ 手动调一次（手动调过的位置以表格为准）。
    """
    return tuple(
        (0, int(p), "") if p.isdigit() else (1, 0, p.lower())
        for p in _NUM_RE.split(name)
    )


def _key_by_name(p: Path) -> tuple:
    """按**文件名**排 —— 全局一把尺，不看来路。

    ⭐ 2026-10-08 用户拍板（方式二）：来源可以是一堆散在各处的文件夹/文件，
       队列一律按**文件名**统一排，**不按来源分组** ——
       加进来的先后、来自哪个文件夹，都不影响队伍。

    ⚠️ 不同文件夹里可能有同名文件：再按所在目录比一次，保证结果唯一确定
       （否则同名两份的先后会随文件系统返回顺序飘）。
    """
    return (natural_key(p.name), natural_key(str(p.parent)))


class OrderRule(NamedTuple):
    """① 一条排序规则。加一种排法就往 ORDER_RULES 里加一行。"""

    key: str
    label: str
    desc: str
    sorter: Callable[[list[Path]], list[Path]]


ORDER_RULES: tuple[OrderRule, ...] = (
    OrderRule(
        "name",
        "文件名顺序",
        "全部原件统一按**文件名**排（不看来自哪个文件夹）：字母不分大小写、数字按数值排。"
        "原件不用改名成 001/002，WI-1032、WI-2001 这类原名直接排",
        lambda paths: sorted(paths, key=_key_by_name),
    ),
    OrderRule(
        "mtime",
        "修改时间顺序",
        "按文件修改时间从早到晚排（原件没有编号、但拷贝/生成有先后的情况用）",
        lambda paths: sorted(paths, key=lambda p: (_safe_mtime(p), _key_by_name(p))),
    ),
)
DEFAULT_ORDER = "name"


def _safe_mtime(p: Path) -> float:
    """读不出修改时间就当作 0，别让一份文件把整个预览搞崩。"""
    try:
        return p.stat().st_mtime
    except OSError:
        return 0.0


def sort_originals(paths: Sequence[Path],
                   rule_key: str = DEFAULT_ORDER) -> list[Path]:
    for rule in ORDER_RULES:
        if rule.key == rule_key:
            return rule.sorter(list(paths))
    raise ValueError(f"没有这条排序规则：{rule_key}")


# ══════════════════════════════════════════════════════════════════════════
# ③ 贴合决策表（沿用 v1 已验证的口径）
# ══════════════════════════════════════════════════════════════════════════

FIT_ASPECT_TOLERANCE = 0.15
FIT_FIT = "贴合"
FIT_ASIS = "原样"


# ══════════════════════════════════════════════════════════════════════════
# ⑥ 扫描页方向（人工判定 —— 机器判不出来）
# ══════════════════════════════════════════════════════════════════════════
#
# 现实情况：扫描仪不管纸是横是竖，出来一律竖版。一张**横版表格**被竖着扫进去后，
# 内容就是躺着的（而且可能头朝左、也可能头朝右）。这活儿没法自动判：
# 躺着的字和正着的字在程序看来都是一堆墨迹，判方向必然有误判。
# 所以交给人在表格里点一下 —— 规则表只管"怎么转"，不管"该不该转"。

ROT_ANGLES = (0, 90, 180, 270)
ROT_LABELS = {
    0: "正向",
    90: "右转 90°",
    180: "倒转 180°",
    270: "左转 90°",
}

# 角度一律按 PDF 的 /Rotate 语义：**正数 = 顺时针**（向右转）。
# 实测确认：pypdf 的 Page.rotate(angle) 是顺时针，transfer_rotation_to_content()
# 会把 /Rotate 烘进内容流并自动交换 mediabox 宽高 —— 两者配套才生效。
ROT_LEFT = -90
ROT_RIGHT = 90

# ── 自动方向规则（⑦）—— 用户 2026-10-08 拍板 ──────────────────────────────
# 前提：**扫描件一律是竖着扫的**（扫描仪不管纸横竖，出来都是竖版）。
# 推论：如果一份原件的首页是**横向**的，那它的封面被竖着扫进来后内容必然是躺着的，
#       所以对应的那一页扫描件**默认左转 90°（逆时针，= 270）**把它扳正。
# ⚠️ 这条只是「默认值」，不是结论：人一眼就能看出对不对，改起来也就一个按钮。
# ⚠️ 自动值每次都按当前的【行 → 页】映射重算：页码一动（调页/反转页序），
#    「这一页是谁的封面」就变了，自动值跟着变才对。
AUTO_ROT_LANDSCAPE = 270          # 左转 90°（逆时针）
AUTO_ROT_UPRIGHT = 0              # 竖向原件：不用转

# 用户 2026-10-08 明确：「扫描件我人为定义必须竖着去打印机扫描的，不会存在横向的」。
# 所以下面这条守卫平时永远不会触发；真触发了说明这批纸不是竖扫的（扫描仪设置变了、
# 纸放错了），值得出声提醒一句 —— 静默跳过等于把异常藏起来。
NOTE_SCAN_LANDSCAPE = "这一页是横向的（说好扫描件都竖扫），已跳过自动左转"


def is_landscape(size: tuple[float, float]) -> bool:
    """原件首页是不是横向（宽 > 高）—— 用**显示尺寸**，/Rotate 90 的横页也算横。"""
    if not size or min(size) <= 0:
        return False
    return size[0] > size[1]


def auto_rotations(items: Sequence["Item"],
                   raw_sizes: Sequence[tuple[float, float]] | None = None
                   ) -> dict[int, int]:
    """按当前映射算出每页的**默认**方向：横向原件用的那一页 → 左转 90°。

    ⚠️ 前提「扫描件一律竖扫」不成立时要让路：这一页**本来就是横向**的，
       说明它已经是横着扫出来的（不是躺着的内容），再转一次反而转反。
    """
    out: dict[int, int] = {}
    for it in items:
        if not is_landscape(it.src_size):
            out[it.scan_index] = AUTO_ROT_UPRIGHT
            continue
        raw = (raw_sizes[it.scan_index - 1]
               if raw_sizes and 0 < it.scan_index <= len(raw_sizes) else it.scan_size)
        out[it.scan_index] = AUTO_ROT_UPRIGHT if is_landscape(raw) else AUTO_ROT_LANDSCAPE
    return out


def resolve_rotations(items: Sequence["Item"], manual: dict[int, int],
                      auto_on: bool = True,
                      raw_sizes: Sequence[tuple[float, float]] | None = None
                      ) -> dict[int, int]:
    """最终方向 = 人工 > 自动 > 正向。界面与命令行共用这一个口径（不许各算各的）。"""
    auto = auto_rotations(items, raw_sizes) if auto_on else {}
    return {
        it.scan_index: int(manual.get(it.scan_index, auto.get(it.scan_index, 0))) % 360
        for it in items
    }


def turn_deg(cur: int, delta: int) -> int:
    """左转 -90 / 右转 +90，循环回 0~270（转四次回到原样）。"""
    return (int(cur) + int(delta)) % 360


def rotated_size(size: tuple[float, float], deg: int) -> tuple[float, float]:
    """旋转后的显示尺寸 —— 90/270 时宽高互换，180 不变。"""
    if not size or min(size) <= 0:
        return size
    deg = int(deg) % 360
    return (size[1], size[0]) if deg in (90, 270) else size


def parse_rotations(raw: str) -> dict[int, int]:
    """命令行用的解析：「3:90,5:270」→ {3: 90, 5: 270}（页码 1 起）。"""
    out: dict[int, int] = {}
    for part in (raw or "").split(","):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise ValueError(f"看不懂这段旋转设置：{part}（应该写成 页码:角度，例如 3:90）")
        page, deg = part.split(":", 1)
        d = int(deg.strip()) % 360
        if d not in ROT_ANGLES:
            raise ValueError(f"角度只能是 90 的倍数：{part}")
        out[int(page.strip())] = d
    return out


# ══════════════════════════════════════════════════════════════════════════
# ⑤ 状态枚举
# ══════════════════════════════════════════════════════════════════════════

class Status(str, Enum):
    PENDING = "待替换"
    OK = "替换成功"
    FAILED = "替换失败"
    SKIPPED = "已跳过"


# ══════════════════════════════════════════════════════════════════════════
# ⑨ 撞名时的态度 —— **三态**，不是一个 bool
# ══════════════════════════════════════════════════════════════════════════
#   AUTO(None) —— 没被问过 / 命令行没表态：只覆盖「认得出是自己上一轮产物」的，
#                 认不出的（别人手动放的、别的软件生成的）加编号让开。
#                 这一态就是「重跑＝原地更新」，不会攒出一堆 _2。
#   YES(True)  —— 用户看着清单点了「是」：一律覆盖。
#   NO(False)  —— 用户看着清单点了「否」：一律让开加编号，**什么都不覆盖**。
#
# ⚠️ 为什么必须是三态：原来只有 True/False 两态时，界面弹窗点「否」会置 False，
#    但 resolve_target 里「认得出是产物就覆盖」这条**与 overwrite 无关**地生效，
#    于是用户明确说了「不覆盖」，目标文件还是被盖掉了 —— 弹窗等于摆设。
# ⚠️ 它跟弹窗的「取消」不是一回事：取消＝这次根本不跑（在界面层就 return 了），
#    不会走到这里。
OVERWRITE_AUTO: None = None
OVERWRITE_YES: bool = True
OVERWRITE_NO: bool = False


@dataclass
class Options:
    """选项开关（界面与命令行共用一套默认值）。"""

    # ⭐ 原件来源清单：每项可以是「一个文件夹」或「一个 PDF 文件」，可以混着来。
    #    （2026-10-08 之前这里是一个 `src_dir` 字符串；换成清单后就能一次收一批
    #      散在不同文件夹里的原件，也能只挑某几份。详见模块文档。）
    sources: tuple[str, ...] = ()
    # 扫描件 PDF。可以留空 —— 那时预览只**列队**、不配对（见 Plan.paired）。
    scan_pdf: str = ""
    # 扫描仪出的是倒序（最后扫的在最前面）时打开
    reverse_order: bool = False
    scale_to_original: bool = True
    # ⑨ 目标名被占用时的态度，三态（见上面 OVERWRITE_* 的说明）：
    #    AUTO  = 只覆盖认得出的自家产物（默认，重跑＝更新）
    #    True  = 一律覆盖
    #    False = 一律加编号让开
    overwrite: bool | None = OVERWRITE_AUTO
    # ④ **必填**：产物就写在原件旁边，后缀为空会让产物跟原件同名、把原件覆盖掉
    suffix: str = ""
    order_rule: str = DEFAULT_ORDER
    keep_metadata: bool = True
    dry_run: bool = False
    # ⑥ 人工判的方向：{扫描页序号(1起): 角度}，没记的就用⑦的自动规则
    rotations: dict[int, int] = field(default_factory=dict)
    # ⑦ 自动方向：原件首页是横向的 → 它那一页默认左转 90°（前提：扫描件一律竖扫）
    auto_rotate: bool = True
    # ⑬ 转换产物落在哪个目录。留空 = 老口径「写在原件旁边」。
    #    ⭐ 6.0 新增，只给合并打印用：那一轮的中间 PDF 是给拼接吃的，
    #       不该散落在用户的原件文件夹里，全部丢进临时目录，拼完即删。
    #    ⚠️ 填了就**不再**做「认不认得出是自家产物」那套原地更新判定 ——
    #       临时目录每轮都是新的，撞名只可能来自同名原件，加编号让开才是对的。
    out_dir: str = ""


# ══════════════════════════════════════════════════════════════════════════
# 数据模型
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class Item:
    """一份原件 + 它对应的那一页扫描页。"""

    index: int                  # 1 起的排队序号
    src: Path
    src_pages: int = 0          # 0 = 读不出来
    src_size: tuple[float, float] = (0.0, 0.0)
    scan_index: int = 0         # 1 起的扫描页序号
    scan_size: tuple[float, float] = (0.0, 0.0)
    rotation: int = 0           # ⑥ 这一页人工判的方向（0 正向 / 90 右转 / 270 左转 / 180 倒转）
    status: str = Status.PENDING.value
    note: str = ""
    error: str = ""


@dataclass
class Plan:
    """一次预览的结果。blocked=True 时**一个文件都不许输出**。

    `paired` 区分这是「只列了队」还是「已经配上扫描件」：
        False → 只有原件队列（没给扫描件），表格里扫描页/方向都是空的；
        True  → 完整配对，可以开始替换。
    """

    items: list[Item] = field(default_factory=list)
    blocked: bool = False
    block_reason: str = ""
    paired: bool = False
    scan_pages: int = 0
    scan_name: str = ""
    # 扫描件各页的**原始**尺寸（旋转前的）。⚠️ 不能靠 item.scan_size 反推 ——
    # 那个字段每次 apply_rotations 都会被改写，拿它当原始值会「转一次变一次」
    scan_raw_sizes: list[tuple[float, float]] = field(default_factory=list)
    # ⑧ 收集原件时被认出来、跳过掉的上一轮产物（界面拿它告诉用户跳过了几份）
    products: list[Path] = field(default_factory=list)


@dataclass
class SwapResult:
    src: Path
    out: Path
    scan_index: int = 0
    rotation: int = 0
    ok: bool = False
    status: str = Status.PENDING.value
    fit_mode: str = ""
    fit_reason: str = ""
    pages_in: int = 0
    pages_out: int = 0
    note: str = ""


# ══════════════════════════════════════════════════════════════════════════
# 读 PDF
# ══════════════════════════════════════════════════════════════════════════

def _reader(path: Path):
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"PDF 已加密，读不了（{exc}）") from exc
    return reader


def display_size(page) -> tuple[float, float]:
    """页面「显示尺寸」：/Rotate 90/270 的页面按横放后的尺寸返回。"""
    box = page.mediabox
    w, h = float(box.width), float(box.height)
    if abs(float(getattr(page, "rotation", 0) or 0)) % 180 == 90:
        w, h = h, w
    return w, h


def read_pdf_info(path: Path) -> tuple[int, tuple[float, float], str]:
    """(页数, 首页显示尺寸, 错误信息)。出错返回 (0, (0,0), 原因)。"""
    try:
        reader = _reader(path)
        n = len(reader.pages)
        if n == 0:
            return 0, (0.0, 0.0), "里面没有任何页面"
        return n, display_size(reader.pages[0]), ""
    except Exception as exc:  # noqa: BLE001
        return 0, (0.0, 0.0), str(exc) or exc.__class__.__name__


def scan_page_sizes(path: Path) -> tuple[list[tuple[float, float]], str]:
    """扫描件每一页的显示尺寸。用于预览表格。"""
    try:
        reader = _reader(path)
        return [display_size(p) for p in reader.pages], ""
    except Exception as exc:  # noqa: BLE001
        return [], str(exc) or exc.__class__.__name__


# ══════════════════════════════════════════════════════════════════════════
# ⑧ 产物识别 —— 认出上一轮留下的产物，别把它当原件收
# ══════════════════════════════════════════════════════════════════════════

# 「加编号让开」留下的尾巴：X_已替换.pdf 撞名时会让成 X_已替换_2.pdf
_NUMBERED_TAIL = re.compile(r"^(?P<base>.+)_(?P<n>\d+)$")


def sibling_source_of(p: Path, suffix: str) -> Path | None:
    """p 是不是「同目录里某份原件的产物」？是就返回那份原件，否则 None。

    两条判据缺一不可（宁可漏判，也不能误伤真原件）：
        ① 文件名 = 原件名 + 当前后缀（后缀非空）
           —— 也接受「再带一个 _2 / _3 编号尾巴」的形态（见下面的 ⚠️）
        ② 同目录**确实存在**那份原件

    ⚠️ 为什么必须有这条：产物现在写在原件旁边（不再有独立输出文件夹），
       第二次跑时它会跟原件一起被收进来 —— 数量直接翻倍、顺序配对整批报废。
    ⚠️ 为什么要连 `_2` 尾巴一起认：⑨ 表里「让开」那一支产出的正是
       `X_已替换_2.pdf`。只认无编号形态的话，用户在上一步选了「否」之后，
       这些编号件下一轮就会被当成原件收进来 → 数量翻倍 → 整批拒绝，
       他还得自己动手去删。认了它们，这条路才是闭环的。
    ⚠️ 为什么只认这两条、不做更聪明的猜测：猜错的代价是把用户真正的原件藏起来。
       「多收一份」会被数量闸门拦住（出声），「藏起一份」是静默少一份（不出声）——
       后者的危害大得多，所以判据宁严勿宽。多认一个 `_N` 尾巴不会新增误判类别：
       它要求的 `X_已替换.pdf` 本身按无编号规则就已经被当成产物跳过了。
    ⚠️ 已知边界：用户中途换后缀 → 旧后缀的产物认不出来，会被当成原件。
       后果只是数量对不上（闸门拦住并给出提示），不会静默出错。
    """
    sfx = clean_output_suffix(suffix)
    if not sfx:
        return None              # 后缀为空时压根不该有产物（suffix_error 会挡住执行）
    stem = p.stem
    if not stem.endswith(sfx):
        m = _NUMBERED_TAIL.match(stem)
        if not m or not m.group("base").endswith(sfx):
            return None
        stem = m.group("base")   # 剥掉 _2 / _3，回到「原件名 + 后缀」
    origin = stem[: -len(sfx)]
    if not origin:
        return None              # 名字整个就是后缀（如「_已替换.pdf」），不认
    cand = p.with_name(origin + p.suffix)
    return cand if cand.is_file() else None


def is_product(p: Path, suffix: str) -> bool:
    """⑧ p 是不是上一轮的产物（收集原件时要跳过它）。"""
    return sibling_source_of(p, suffix) is not None


# ══════════════════════════════════════════════════════════════════════════
# 收集原件
# ══════════════════════════════════════════════════════════════════════════

def collect_originals(sources: Sequence[str | Path], options: Options,
                      skipped: list[Path] | None = None) -> list[Path]:
    """把来源清单里的原件收拢成**一队**：排除扫描件自身、排除上一轮的产物、去重。

    每个来源按类型处理：
        · 文件夹 → 只收它**这一层**里的 PDF（`.glob("*")`，不往子目录钻）
        · 单个 PDF 文件 → 直接用

    `skipped` 传进来的话，会把「看得出是上次产物才跳过」的文件记进去 ——
    界面拿它告诉用户跳过了几份，免得用户以为文件凭空少了。
    """
    scan_r = (Path(options.scan_pdf).expanduser().resolve()
              if str(options.scan_pdf).strip() else None)

    hit: list[Path] = []
    seen: set[Path] = set()
    for s in sources:
        entry = Path(s).expanduser()
        if entry.is_dir():
            cands = [c for c in sorted(entry.glob("*"))
                     if c.is_file() and c.suffix.lower() in PDF_EXTS]
        elif entry.is_file() and entry.suffix.lower() in PDF_EXTS:
            cands = [entry]
        else:
            continue                      # 不存在 / 不是 PDF → 交给闸门去报
        for p in cands:
            try:
                pr = p.resolve()
            except OSError:
                continue
            if scan_r and pr == scan_r:
                continue                  # 扫描件本身不算原件
            if pr in seen:
                continue                  # 同一个文件被指到两次（既加了文件夹又单加了它）
            if is_product(p, options.suffix):
                if skipped is not None:
                    skipped.append(p)
                continue
            seen.add(pr)
            hit.append(p)
    return sort_originals(hit, options.order_rule)


# ══════════════════════════════════════════════════════════════════════════
# ② 阻断闸门表 —— 任一条命中，整批拒绝执行
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class Ctx:
    """闸门判断要看的东西，一次算好传给所有闸门。"""

    sources: list[Path]
    scan_pdf: Path | None
    originals: list[Path]
    scan_pages: int
    scan_err: str
    broken: list[tuple[Path, str]]     # 读不出页数的原件


def _g_no_src(ctx: Ctx) -> str:
    if not ctx.sources:
        return ("还没有添加任何原件来源。\n"
                "点「＋ 添加文件夹…」把一整个文件夹的 PDF 收进来，"
                "或「＋ 添加文件…」挑单个 PDF。")
    missing = [p for p in ctx.sources if not p.exists()]
    if missing:
        names = "、".join(str(p) for p in missing[:3])
        more = f" 等 {len(missing)} 个" if len(missing) > 3 else ""
        return (f"这些来源已经不存在了：{names}{more}\n"
                "（被挪走了？在清单里把它移除、再重新添加一下。）")
    return ""


def _g_empty(ctx: Ctx) -> str:
    if not ctx.originals:
        return ("这些来源里没有找到任何 PDF。\n"
                "（文件夹来源只收它**这一层**里的 PDF —— 子文件夹里的不算；"
                "扫描件本身、以及认得出的上一轮产物也会被排除。）")
    return ""


def _g_scan(ctx: Ctx) -> str:
    if ctx.scan_pdf is None:
        return ""                      # 只列队（没给扫描件）时不检查
    if not ctx.scan_pdf.is_file():
        return f"扫描件 PDF 不存在：{ctx.scan_pdf}"
    if ctx.scan_err:
        return f"扫描件读不出来：{ctx.scan_err}"
    if ctx.scan_pages <= 0:
        return "扫描件里没有任何页面"
    return ""


def _g_count(ctx: Ctx) -> str:
    n_src, n_scan = len(ctx.originals), ctx.scan_pages
    if n_src == n_scan:
        return ""
    diff = abs(n_src - n_scan)
    if n_src > n_scan:
        tail = ("请补扫几页，让两边数量一致再跑。\n"
                "（原件偏多的常见原因：上一轮的产物被当成原件了。按「原件名 + 当前后缀」\n"
                "　认得出的会自动跳过；中途换过后缀的话旧产物认不出来，请把它挪走或统一后缀。）")
    else:
        tail = "请补上漏掉的原件（或删掉多余的扫描页），让两边数量一致再跑。"
    return (f"数量对不上：原件 {n_src} 份、扫描件 {n_scan} 页（差 {diff}）。\n"
            f"顺序配对下缺一个或多余一个都会让后面全部错位，所以整批不执行。\n"
            f"{tail}")


def _g_broken(ctx: Ctx) -> str:
    if not ctx.broken:
        return ""
    names = "、".join(p.name for p, _e in ctx.broken[:5])
    more = f" 等 {len(ctx.broken)} 份" if len(ctx.broken) > 5 else ""
    return (f"这几份原件读不出页数，顺序会错乱，整批不执行：{names}{more}\n"
            f"　{ctx.broken[0][1]}")


class Gate(NamedTuple):
    name: str
    check: Callable[[Ctx], str]


# ⭐ 闸门分两组：只列队时只过 QUEUE_GATES，配对时再叠上 PAIR_GATES。
QUEUE_GATES: tuple[Gate, ...] = (
    Gate("有没有来源", _g_no_src),
    Gate("有没有原件", _g_empty),
)
PAIR_GATES: tuple[Gate, ...] = (
    Gate("扫描件能否读取", _g_scan),
    Gate("数量是否相等", _g_count),
)
BROKEN_GATE = Gate("原件是否都读得出", _g_broken)
# 全量清单（文档 / 回归拿它核对「一条都没漏」）；实际执行顺序见 plan_batch
GATE_RULES: tuple[Gate, ...] = QUEUE_GATES + PAIR_GATES + (BROKEN_GATE,)


# ══════════════════════════════════════════════════════════════════════════
# 预览
# ══════════════════════════════════════════════════════════════════════════

def _scan_index_for(i: int, total: int, reverse: bool) -> int:
    """第 i 份原件（1 起）用扫描件第几页（1 起）。"""
    return (total - i + 1) if reverse else i


def plan_batch(options: Options) -> Plan:
    """把来源清单 + 扫描件算成一份计划。

    两态（见 Plan.paired）：
        · **没给扫描件** → 只列队：收原件、按名字排、读页数/首页尺寸，闸门只过
          「有没有来源 / 有没有原件 / 原件是否读得出」。界面拿它让你先看队。
        · **给了扫描件** → 完整配对：再叠「扫描件能否读 / 数量是否相等」，
          并算好每行用第几页、方向。
    """
    src_paths = [Path(s).expanduser() for s in options.sources if str(s).strip()]
    scan_pdf = Path(options.scan_pdf).expanduser() if str(options.scan_pdf).strip() else None
    paired = scan_pdf is not None

    skipped: list[Path] = []
    originals = collect_originals(options.sources, options, skipped)

    sizes: list[tuple[float, float]] = []
    scan_err = ""
    if scan_pdf is not None:
        if scan_pdf.is_file():
            sizes, scan_err = scan_page_sizes(scan_pdf)
        else:
            scan_err = "文件不存在"

    ctx = Ctx(sources=src_paths, scan_pdf=scan_pdf, originals=originals,
              scan_pages=len(sizes), scan_err=scan_err, broken=[])
    plan = Plan(paired=paired, scan_pages=len(sizes),
                scan_name=(scan_pdf.name if scan_pdf else ""), products=skipped)

    # 读每一份原件的页数/首页尺寸（损坏清单也是闸门之一，所以先读、再过闸门）
    infos: list[tuple[int, tuple[float, float], str]] = []
    for p in originals:
        n, size, err = read_pdf_info(p)
        infos.append((n, size, err))
        if n == 0:
            ctx.broken.append((p, err))

    gates = QUEUE_GATES + (PAIR_GATES if paired else ()) + (BROKEN_GATE,)
    for gate in gates:
        why = gate.check(ctx)
        if why:
            plan.blocked = True
            plan.block_reason = why
            return plan

    total = len(originals)
    for i, p in enumerate(originals, start=1):
        n, size, _err = infos[i - 1]
        si = _scan_index_for(i, total, options.reverse_order) if paired else 0
        plan.items.append(Item(
            index=i, src=p, src_pages=n, src_size=size,
            scan_index=si, scan_size=sizes[si - 1] if 0 < si <= len(sizes) else (0.0, 0.0),
            status=Status.PENDING.value,
        ))
    plan.scan_raw_sizes = list(sizes)
    if paired:
        apply_rotations(plan, options.rotations, auto_on=options.auto_rotate)
    return plan


def apply_rotations(plan: Plan, manual: dict[int, int],
                    raw_sizes: Sequence[tuple[float, float]] | None = None,
                    auto_on: bool = True) -> None:
    """把「最终方向」写回 plan（人工 > 自动 > 正向），并同步扫描页的显示尺寸。

    可以反复调用：每次都从**原始**尺寸重新算，不会把上一次的结果当输入。

    ⚠️ 人工值记在**扫描页**上（那张纸确实是横版内容被竖着扫进来的），页码怎么调都跟着纸走。
    ⚠️ 自动值每次按当前的【行→页】映射**重算** —— 这一页到底是谁的封面变了，
       默认方向就得跟着变（横版原件的封面才需要转）。
    """
    raw = list(raw_sizes) if raw_sizes else plan.scan_raw_sizes
    final = resolve_rotations(plan.items, manual, auto_on, raw)
    for it in plan.items:
        it.rotation = final.get(it.scan_index, 0)
        base = raw[it.scan_index - 1] if 0 < it.scan_index <= len(raw) else it.scan_size
        it.scan_size = rotated_size(base, it.rotation)
        # 横向原件 + 横向扫描页：自动左转被跳过，出声说明原因（见 NOTE_SCAN_LANDSCAPE）
        if auto_on and is_landscape(it.src_size) and is_landscape(base) \
                and it.scan_index not in manual:
            it.note = NOTE_SCAN_LANDSCAPE


def reindex(plan: Plan, reverse: bool) -> None:
    """切「正序 / 倒序」：整批重算扫描页序号（原件顺序不动）。

    ⚠️ 这是「重来一遍」而不是「把当前结果再反一次」：手动调过的页码会被覆盖掉 ——
       按钮上写着当前是正序还是倒序，来回点是幂等的，猜不出自己按了几次。
    """
    total = len(plan.items)
    for it in plan.items:
        it.scan_index = _scan_index_for(it.index, total, reverse)


def assign_page(plan: Plan, item_index: int, page: int) -> tuple[bool, str]:
    """把第 item_index 行（1 起）改成用扫描件第 page 页。

    ⭐ 一页只许一个人用：目标页已经被别的行占着时**直接对调**，
       而不是让两行都用同一页（那样必然有一份原件拿不到自己的封面）。
    返回 (有没有变化, 人话说明)。
    """
    n = len(plan.items)
    if not (1 <= item_index <= n) or page <= 0 or page > plan.scan_pages:
        return False, ""
    it = plan.items[item_index - 1]
    if it.scan_index == page:
        return False, ""
    old = it.scan_index
    other = next((x for x in plan.items
                  if x is not it and x.scan_index == page), None)
    it.scan_index = page
    if other is not None:
        other.scan_index = old
        return True, f"和《{other.src.name}》对调了（它改成用第 {old} 页）"
    return True, f"改用第 {page} 页（第 {old} 页现在没人用）"


def swap_pages(plan: Plan, item_index: int, delta: int) -> bool:
    """与相邻行对调扫描页（delta=-1 与上一行、+1 与下一行）。

    行本身不动 —— 行的顺序就是原件在文件夹里的名称顺序，那是唯一的排队依据；
    能调的只有「这一行用哪一页」。
    """
    n = len(plan.items)
    j = item_index + delta
    if not (1 <= item_index <= n) or not (1 <= j <= n):
        return False
    a, b = plan.items[item_index - 1], plan.items[j - 1]
    a.scan_index, b.scan_index = b.scan_index, a.scan_index
    return True


# ══════════════════════════════════════════════════════════════════════════
# ④ 输出文件名后缀 —— 现在是**必填**，不是"建议填"
# ══════════════════════════════════════════════════════════════════════════

SUFFIX_FORBIDDEN = frozenset('\\/:*?"<>|') | {chr(i) for i in range(32)}
SUFFIX_MAX_LEN = 40

# 为什么必填：产物就写在原件旁边。后缀为空 → 产物名 = 原件名 → 直接覆盖原件本身。
# 所以这是**硬闸门**，三道：界面禁用按钮 / 命令行拒绝 / run_batch 再兜一道。
SUFFIX_REQUIRED_MSG = (
    "必须先填「输出文件名后缀」。\n\n"
    "产物现在直接写在原件旁边（不再有单独的输出文件夹），"
    "后缀为空会让产物跟原件一模一样、把原件覆盖掉。\n\n"
    "例如填 _已替换：\n"
    "　原件 WI-1032.pdf  →  产物 WI-1032_已替换.pdf（原件原封不动）"
)


def clean_output_suffix(raw: str) -> str:
    """后缀规范化唯一口径（界面与命令行共用，避免规则漂移）。

    刻意不做「自动补下划线」的猜测 —— 填什么就是什么。
    """
    s = (raw or "").strip()
    s = "".join(ch for ch in s if ch not in SUFFIX_FORBIDDEN)
    s = s.strip(".")
    return s[:SUFFIX_MAX_LEN]


def suffix_error(options: Options) -> str:
    """后缀够不够用。返回 "" = 没问题；否则是人话原因。"""
    if clean_output_suffix(options.suffix):
        return ""
    return SUFFIX_REQUIRED_MSG


def out_name_for(src: Path, suffix: str) -> str:
    s = clean_output_suffix(suffix)
    return f"{src.stem}{s}{src.suffix}"


def _unique_out_path(base: Path, used: set[Path]) -> Path:
    """加编号避开：xxx_2.pdf、xxx_3.pdf。

    ⚠️ 占用判定必须**同时**看两边：批内已经排过的（used）+ 磁盘上真实存在的。
       只查 used 是不够的 —— 原来那份 used 是靠预扫输出目录填的，现在没有输出目录了。
    """
    def taken(p: Path) -> bool:
        return p in used or p.exists()

    if not taken(base):
        used.add(base)
        return base
    stem, suf, parent = base.stem, base.suffix, base.parent
    k = 2
    while True:
        cand = parent / f"{stem}_{k}{suf}"
        if not taken(cand):
            used.add(cand)
            return cand
        k += 1


# ══════════════════════════════════════════════════════════════════════════
# ⑨ 落点决策 —— 产物写哪儿、目标名被占了怎么办
# ══════════════════════════════════════════════════════════════════════════

def _resolve_target(base: Path, options: Options, used: set[Path], own: bool) -> Path:
    """⑨ 表的**共用实现**：目标名已被占用时怎么走。

    `own` = 这个目标名认得出是「本工具自己上一轮生成的产物」。
    两套流程（首页替换 / 格式转换）各有一个 `*_target` 外壳算 base 与 own，
    判定表本身只有这一份 —— 免得两边漂移。
    """
    if base not in used and not base.exists():
        used.add(base)                       # 无人占用 → 直接用，三态都一样
        return base
    if options.overwrite is OVERWRITE_YES:   # 用户点了「是」→ 一律覆盖
        used.add(base)
        return base
    if options.overwrite is OVERWRITE_NO:    # 用户点了「否」→ 一律让开
        return _unique_out_path(base, used)
    if own:                                  # AUTO：认得出是自家产物 → 原地更新
        used.add(base)
        return base
    return _unique_out_path(base, used)      # AUTO：认不出 → 不碰，加编号


def resolve_target(src: Path, options: Options, used: set[Path]) -> Path:
    """算出这份原件该写到哪个路径。一张表说清（`overwrite` 是**三态**）：

    | 目标名的状态                     | AUTO（默认）   | YES（点是） | NO（点否） |
    |----------------------------------|----------------|-------------|------------|
    | 根本没这个文件                   | 写它           | 写它        | 写它       |
    | 存在，且认得出是上一轮产物       | **覆盖**（更新）| 覆盖        | 加编号让开 |
    | 存在，但认不出是谁               | 加编号让开      | 覆盖        | 加编号让开 |

    ⭐ AUTO 一句话：**工具只动自己生出来的东西**。上一轮的产物认得出就更新它；
       认不出的（用户手动放的、别的软件生成的）一律不碰，加 _2 绕开。
    ⭐ NO 一句话：用户说「不覆盖」就是**什么都不覆盖** —— 撞上就加编号。
    ⚠️ 第三行**绝不会是原件本身** —— 产物名 = 原件名 + 非空后缀（suffix_error 保证非空），
       跟原件名必然不同。这也是后缀必填换来的安全性。
    """
    base = src.parent / out_name_for(src, options.suffix)
    return _resolve_target(base, options, used, is_product(base, options.suffix))


# ══════════════════════════════════════════════════════════════════════════
# 首页替换（沿用 v1 已验证的贴合逻辑，区别：取扫描件【第 N 页】）
# ══════════════════════════════════════════════════════════════════════════

def decide_fit(src_disp: tuple[float, float], scan_disp: tuple[float, float],
               tolerance: float = FIT_ASPECT_TOLERANCE) -> tuple[str, str]:
    """③ 贴合决策：返回 (FIT_FIT / FIT_ASIS, 人话原因)。"""
    if min(src_disp) <= 0 or min(scan_disp) <= 0:
        return FIT_ASIS, "尺寸信息缺失，原样保留扫描页"
    r_src = src_disp[0] / src_disp[1]
    r_scan = scan_disp[0] / scan_disp[1]
    diff = abs(r_src - r_scan) / max(r_src, r_scan)
    if diff <= tolerance:
        return FIT_FIT, f"版面比例接近（宽高比差异 {diff:.0%}），按原件首页尺寸等比缩放居中"
    return FIT_ASIS, (f"版面比例不一致（原件 {r_src:.2f} vs 扫描页 {r_scan:.2f}），"
                      f"扫描页原样作为第 1 页，不硬塞进原页面")


def _prepare_scan_page(scan_page, tw: float, th: float):
    """把扫描页等比缩放并居中到 (tw, th) 上。"""
    from pypdf import Transformation

    try:  # pypdf >= 3.16：把 /Rotate 烘进内容流，避免旋转页错位
        scan_page.transfer_rotation_to_content()
    except Exception:  # noqa: BLE001
        pass

    box = scan_page.mediabox
    sw, sh = float(box.width), float(box.height)
    if sw <= 0 or sh <= 0:
        return scan_page
    k = min(tw / sw, th / sh)
    ox = (tw - sw * k) / 2.0
    oy = (th - sh * k) / 2.0
    scan_page.add_transformation(Transformation().scale(k, k).translate(ox, oy))
    scan_page.mediabox.lower_left = (0, 0)
    scan_page.mediabox.upper_right = (tw, th)
    return scan_page


def swap_first_page(src_path: Path, scan_path: Path, scan_index: int,
                    out_path: Path, options: Options,
                    rotation: int = 0) -> SwapResult:
    """核心动作：输出 = [扫描件第 scan_index 页] + [原件第 2..N 页]。

    ⚠️ 扫描件**每份原件重新打开一次**：_prepare_scan_page 会就地改页面对象，
       复用同一个 reader 的 page 会让第二份输出被缩放两次。

    ⑥ rotation：人工判的方向（正数顺时针）。必须先 rotate 再 transfer_rotation_to_content
       —— 后者会把 /Rotate 烘进内容流并交换 mediabox，之后 display_size 才是转好的尺寸，
       贴合决策要用转好的尺寸去比，否则会拿竖着的尺寸去套横着的原件。
    """
    from pypdf import PdfReader, PdfWriter

    res = SwapResult(src=src_path, out=out_path, scan_index=scan_index,
                     rotation=int(rotation) % 360)
    try:
        # 🛡 最后一道保险：产物名跟原件名撞上了就绝不动手。
        #    正常路径下 suffix_error 已经挡住了空后缀，这里只是兜底 ——
        #    「原件只读」是这工具最硬的一条承诺，宁可不跑也不能盖原件。
        if out_path.resolve() == src_path.resolve():
            raise RuntimeError("产物名与原件同名（后缀是不是空的？）—— 已阻止，不覆盖原件")
        src_reader = _reader(src_path)
        n_in = len(src_reader.pages)
        if n_in == 0:
            raise RuntimeError("原 PDF 内没有任何页面")
        res.pages_in = n_in

        scan_reader = _reader(scan_path)
        if not (1 <= scan_index <= len(scan_reader.pages)):
            raise RuntimeError(f"扫描件只有 {len(scan_reader.pages)} 页，取不到第 {scan_index} 页")
        scan_page = scan_reader.pages[scan_index - 1]

        # ⑥ 先转向，再把旋转烘进内容流（此后 mediabox 已是转好的尺寸）
        if res.rotation:
            scan_page.rotate(res.rotation)
        try:
            scan_page.transfer_rotation_to_content()
        except Exception:  # noqa: BLE001
            pass

        mode, reason = (FIT_ASIS, "已关闭缩放，扫描页原样作为第 1 页")
        if options.scale_to_original:
            mode, reason = decide_fit(display_size(src_reader.pages[0]), display_size(scan_page))
        res.fit_mode, res.fit_reason = mode, reason

        writer = PdfWriter()
        if mode == FIT_FIT:
            tw, th = display_size(src_reader.pages[0])
            writer.add_blank_page(width=tw, height=th)
            writer.pages[-1].merge_page(_prepare_scan_page(scan_page, tw, th))
        else:
            writer.add_page(scan_page)

        for idx, page in enumerate(src_reader.pages):
            if idx == 0:
                continue
            writer.add_page(page)

        if options.keep_metadata and src_reader.metadata:
            try:
                writer.add_metadata({k: str(v) for k, v in src_reader.metadata.items()})
            except Exception:  # noqa: BLE001
                pass

        res.pages_out = len(writer.pages)

        if options.dry_run:
            res.ok = True
            res.status = Status.OK.value
            res.note = "演练：没有真的写文件"
            return res

        out_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = out_path.with_suffix(out_path.suffix + ".part")
        with open(tmp, "wb") as fh:
            writer.write(fh)
        os.replace(tmp, out_path)

        res.ok = True
        res.status = Status.OK.value
        res.note = reason if mode == FIT_ASIS else ""
        return res
    except Exception as exc:  # noqa: BLE001
        res.ok = False
        res.status = Status.FAILED.value
        res.note = str(exc) or exc.__class__.__name__
        return res


def overwrite_list(plan: Plan, options: Options) -> list[tuple[Path, Path]]:
    """**哪些已有的文件会被盖掉** → [(已有文件, 对应原件)]。

    给界面弹确认窗用：把清单摊给人看，他点了确定才真的覆盖。
    ⚠️ 这里**不看 `options.overwrite`** —— 它回答的是「如果选择覆盖，会动到谁」，
       是对着 `是` 那个分支算的。真正的落点由 `resolve_target` 按三态决定：
         YES  → 清单上这些全盖
         NO   → 一个都不盖，全加编号
         AUTO → 只有「认得出是自己上一轮产物」的那些被盖
    """
    out: list[tuple[Path, Path]] = []
    seen: set[Path] = set()
    for it in plan.items:
        if it.status == Status.SKIPPED.value:
            continue
        t = it.src.parent / out_name_for(it.src, options.suffix)
        if t.exists() and t not in seen:
            seen.add(t)
            out.append((t, it.src))
    return out


def run_batch(plan: Plan, options: Options) -> tuple[list[SwapResult], list[str]]:
    """按预览的顺序逐个替换。返回 (结果, 提示语)。

    ⚠️ 落点：不再是统一的输出文件夹，而是**每份原件自己所在的目录** ——
       原件在子文件夹里，产物也跟着落进那个子文件夹，天然不会同名打架，
       所以 v2 那套「递归时复刻子目录结构」(`_out_sub`) 整段不需要了。
    ⚠️ 后缀必填在这里**再兜一道**：不指望调用方（界面 / 命令行）都记得检查。
    ⚠️ 三道闸门的**先后**有意义：后缀是「这次请求本身就没法落地」（跟数据无关），
       排最前；然后是「还没配对扫描件」；最后才是 `plan.blocked`（数据对不上）。
       空后缀时如果先报数量问题，用户会去数页数，白忙一场 —— 真正该改的是填后缀。
    """
    why = suffix_error(options)
    if why:
        return [], [why]
    if not plan.paired:
        return [], ["还没配对扫描件 —— 先选好扫描件 PDF、点一次预览，配对上了再替换。"]
    if plan.blocked:
        return [], [plan.block_reason or "被阻断，未执行"]

    used: set[Path] = set()
    scan = Path(options.scan_pdf).expanduser()
    results: list[SwapResult] = []
    for it in plan.items:
        if it.status == Status.SKIPPED.value:
            results.append(SwapResult(src=it.src, out=Path(""), scan_index=it.scan_index,
                                      ok=True, status=Status.SKIPPED.value,
                                      note="你手动忽略了这一行"))
            continue
        target = resolve_target(it.src, options, used)
        results.append(swap_first_page(it.src, scan, it.scan_index, target, options,
                                       rotation=it.rotation))
    return results, []


# ══════════════════════════════════════════════════════════════════════════
# ⑩ 格式转换表 —— 认哪些扩展名、用哪个应用、LibreOffice 用哪个过滤器
# ══════════════════════════════════════════════════════════════════════════
#
# 加一种格式（比如 PPT）＝ 往 CONVERT_FORMATS 加一行，别处不用动：
#   · 收件（collect_convertibles）按 exts 认；
#   · macOS 走 AppleScript 的 tell application <mac_app>；
#   · Windows 走 COM 的 DispatchEx(<win_app>)；
#   · LibreOffice 走 --convert-to pdf:<soffice_filter>。
#
# ⭐ 表里每一行都要求**两个平台各自说得通**。万一将来加进来一种「只在 Windows
#    上有的格式」（比如只出了 Windows 版的那类行业软件），就把它缺的那一边**整个留空**
#    —— mac_app 与 mac_ 三步全空，由 assert_format_complete() 认「空 ⇔ 空」这个组合
#    （半空半填＝表自己矛盾，当场炸）。在缺的那一边碰到它时，工具**不会去瞎叫一个
#    不存在的应用**，而是直接说明原因（见 _NO_MAC_APP_NOTE）；若那台机器上恰好有
#    LibreOffice，兜底梯队照样会接过去试。
#    ⚠️ 别为了「填满表」给不存在的平台硬编一套语句 —— 那会变成一句看不懂的报错。
#
#    ⚠️ 2026-10-09：清掉 Visio 之后，表里**暂时没有**这样的格式（Word / Excel 两边都有）。
#       这套「一边留空」的机制**留着**（上面那条规矩、_NO_MAC_APP_NOTE、
#       platform_formats 的平台派生都还在），因为它是通用的扩展点，不是 Visio 的私货 ——
#       将来真要加 .btw / .dwg 那种「靠本机某个专用程序」的格式，正好复用。
#       回归里用**合成格式**（把 Word 那一行拷贝出来清空 mac_ 侧）盯着它，见 selftest [13]。
#
# ⚠️ Excel 的**分页细则**（按 sheet 拆 / 按打印区域裁 …）已由 ⑩c 分页规则表落地：
#    表里每行写一个 `page_rule` 的 key，由引擎读完产物后核对页数（不符才出声）。
#    再要加口径就往 ⑩c 加一行，收件 / 闸门 / 落点 / 命名都不用改。
#    现在这一版对外的那句话见 CONVERT_RULE_NOTE（它自己也是从 ⑩c 派生的）。


# ⚠️⚠️ 这一段踩过两次坑，第二次（2026-10-09）把它从**结构上**堵死了：
#
#   1. `tell application` 后面的 app 名**必须是字面量**，绝不能写成变量。
#      AppleScript 是编译型语言：写 `tell application appName` 时编译期不知道这是
#      哪个 app、该用哪本词典，于是 `file format format PDF` 这些词全被当成裸标识符
#      → 整段脚本 -2741「syntax error」→ **一份都转不出来**，而且报错看着像
#      「Word 转失败了」，把人往授权 / 沙箱方向带（真凶其实在语法）。
#      ⇒ 现在的做法：app 名从 ⑩ 表的 `mac_app` 字段来，由 `build_mac_script()`
#        **拼进字面量**。不再有「先写 __APP__ 占位、运行时再字符串替换」那一步 ——
#        拼装结构本身就保证是字面量，这个坑从结构上不存在了。
#
#   2. Word 和 Excel 的语句**不能塞进同一个 tell 块**（一个 tell 块只认一本词典，
#      另一种 app 的命令在里面照样编译不过）。⇒ ⑩ 表里一个格式一行，三步各写各的。
#
#   ⚠️ 还有一条不是坑、但一定要知道：macOS 上让别的 app 替你转文件，走的是
#      「自动化」授权（TCC）。用户双击启动器 → Terminal 发事件 → 系统弹一次
#      「允许 Terminal 控制 Microsoft Word」。**从别的进程（IDE、后台服务、
#      被代跑的窗口）拉起来的实例拿不到这个授权**，会一路 -10004。
#      → 所以这个功能请用户**自己双击启动器**用，别在别处代跑。
#
#   3. ⭐⭐ **`set d to open file name …` 这个「一句话打开并拿到文档」的写法不成立**
#      （2026-10-09 的真实事故，报 `-2753 变量"d"没有定义`）。
#      Word / Excel 没在跑的时候，`open` 是**跟着启动事件一起发出去**的，
#      这一句**不给回文档**，于是下一句 `save as d …` 里的 `d` 根本不存在。
#      报错位置指在第 6 行的那个 `d` 上 —— 看着像「Word 转失败了」，
#      真凶却是「文档参考没拿到」。社区里能跑通的写法都是三段式：
#          先让应用起来  →  open 路径  →  等文档真出现，再锁定它
#      现在 ⑩ 表里 Word / Excel 的 mac_open 就是按这个套路写的，
#      并且**用「打开前有几份文档」做参照**锁定这一次新打开的那一份 ——
#      免得用户自己正开着 Word 时，我们去存错了别人的文档。
#      ⚠️ 顺带：Office 的 `save as … file name` 吃 HFS（冒号）路径，
#         所以路径转换统一放在 _MAC_OSA_HEADER 里、**在 tell 块外面**做。
#
#   4. ⭐⭐ **「一份文件一段脚本」这个粒度本身就是错的**（2026-10-09 第二轮）。
#      用户实测反馈两条症状，根子都在这里：
#        · 「每一份都要点一下」—— 授权按应用记，一份一段脚本就等于每份重启应用、
#          每份重问一遍；用户以为「每份文件都要授权一次」，其实是被粒度坑了；
#        · 「3 份只成功 1 份 / 试到第三次才全成 / 隔一会儿又全成」—— 每份都要先把
#          Word 拉起来，撞上「应用刚启动、文档还没交出来」的竞态，纯看运气。
#      ⇒ 现在 build_mac_script() 产出的是**整批循环**：一个格式一段脚本、
#        一个 apply 会话、`activate` 只在循环外一次；
#        Windows 侧对应的是 _run_office_win_many 里整批复用一个 DispatchEx 实例。
#      ⚠️ 别再退回「一份一次」。要加格式就往 ⑩ 表加一行，粒度不归格式表管。

# ⑪a 转换「三步曲」—— 平台无关，只在这里声明一次。
#     两种方言（AppleScript / COM）都必须按这个顺序走；少一步、多一步都算违规。
#     ⚠️ 新增格式（比如 PPT）＝ 往 ⑩ 表加一行 + 在下面两张方言表里补齐三步，
#        别处（收件、闸门、落点、命名、界面）一律不用动。
OFFICE_STEPS: tuple[str, ...] = ("open", "export", "close")
OFFICE_STEP_LABELS: dict[str, str] = {
    "open": "只读打开",
    "export": "导出为 PDF",
    "close": "丢弃改动并关闭",
}

# ⭐⭐ 2026-10-09（第二轮）：从「一份文件一段脚本」改成「**一个格式一段脚本、整批循环**」。
#
#   为什么要改（用户实测反馈，别改回去）：旧写法每份文件都新起一个 osascript，于是
#     · 每份都重新 `activate` 一次 Word / Excel —— 窗口忽前忽后地跳；
#     · 每份都重新走一遍 macOS 的「自动化」授权检查 —— 用户说的「每一份都要点一下允许」
#       就是这么来的（授权本来是**按应用**记的，一格式一段脚本才问得到一次）；
#     · 每份都带一个「应用刚被拉起来、文档还没交出来」的竞态 ——
#       这正是「3 份只成功 1 份 / 试到第三次才全成 / 隔一会儿又全成」那种时好时坏的根源。
#   ⇒ 现在整批塞进**同一个 `tell application` 块**里循环：应用只激活一次、
#     授权最多问一次、竞态只剩「这份文档等它出现」这一个。
#
# AppleScript 的外壳分两段：头（argv 两两一对 → HFS 路径列表）和尾（收结果）。
# 两段都**不含任何用户数据**（路径全从 argv 来），所以整段脚本可以安全地当
# 命令行参数传，不用落地成临时文件。
#
# ⚠️ POSIX → HFS（`Macintosh HD:Users:…`）的转换**必须在 tell 块外面**做。两条理由：
#    · Office 的 `file name` / `workbook file name` 是 text 参数，历史上只认冒号分隔；
#      有人拿 POSIX 斜杠喂它，`save as` 会失败（社区里反复出现过）。
#    · 块内 `POSIX file` 会被当成本 app 词典里的 `file` 类去解析，容易出意外；
#      放外面走的是 AppleScript 自己的强制转换，稳。
#    `POSIX file` 不要求文件存在，所以产物路径（还不存在）照样转得出来（实测过）。
_MAC_OSA_HEADER = (
    "on run argv\n"
    "  set srcList to {}\n"
    "  set dstList to {}\n"
    "  repeat with i from 1 to (count of argv) by 2\n"
    "    set end of srcList to ((POSIX file (item i of argv)) as text)\n"
    "    set end of dstList to ((POSIX file (item (i + 1) of argv)) as text)\n"
    "  end repeat\n"
    "  set nFiles to (count of srcList)\n"      # 上界在 tell 块**外面**算好
)

# 收尾：把每份的结果拼成「一行一份」的文本交回 Python 解析。
# 每行长这样 ——  `2|OK`   或   `2|ERR|-10004|权限违例。`
# ⚠️ 报错文本里可能带换行 / 制表符（AppleScript 的 error message 常带），
#    所以专门有个 neat() 处理器把它们压成空格 —— **一行一份**才解析得干净。
# ⚠️ neat() 是 `on run` 之外的顶层处理器；tell 块里要用 `my neat(...)` 才叫得到它。
_MAC_OSA_FOOTER = (
    "  set AppleScript's text item delimiters to linefeed\n"
    "  return out as text\n"
    "end run\n"
    "\n"
    "on neat(t)\n"
    "  set AppleScript's text item delimiters to {return, linefeed, tab}\n"
    "  set parts to text items of t\n"
    "  set AppleScript's text item delimiters to \" \"\n"
    "  set flat to parts as text\n"
    "  set AppleScript's text item delimiters to \"\"\n"
    "  return flat\n"
    "end neat\n"
)


class ConvFormat(NamedTuple):
    """⑩ 一条格式规则 —— **认哪些扩展名、靠谁转、在哪些平台上转得了**。

    ⭐ **这一行只说「是什么」，不说「怎么转」**：
       · 「怎么转」在 ⑪ 表里（按 `family` 认领）—— 本行只声明自己属于哪一族；
       · 「在哪些平台上转得了」由 `platforms` 说死，**不再靠 `mac_app` 是不是空来暗示**。
         以前在 Mac 上碰到只有 Windows 能转的格式，只能看出「mac_app 空」；
         现在是明写 `platforms=("win",)`，界面文案、列队闸门、回归断言都直接读它。

    Office 那一族（word / excel）才用得上 mac_ 三步与 win_kind；别的族（cad / btw）
    这几个字段**必须留空**（半填半空＝表自己矛盾，`assert_format_complete` 当场炸）。
    """

    key: str                  # "word" / "excel" / "cad" / "btw"（也当引擎认领的「格式」用）
    label: str                # 界面上「类型」列显示的名字，也是「Microsoft ？？」里的那个词
    exts: tuple[str, ...]     # 认这些扩展名（小写，含点）
    # macOS：tell application 的名字（会被拼成**字面量**）。
    # ⭐ "" ＝ 这一族不走 macOS 的 Office 三步曲（CAD / BarTender 就是；
    #    Word / Excel 在 Mac 上也有，所以它们非空）。
    #    「在不在这个平台上」**不由它承担** —— 那是下面 platforms 的事，别混为一谈。
    mac_app: str
    mac_open: str             # macOS：三步曲① 只读打开
    mac_export: str           # macOS：三步曲② 导出为 PDF
    mac_close: str            # macOS：三步曲③ 丢弃改动并关闭
    win_app: str              # Windows：COM 的 ProgID
    win_kind: str             # Windows：用哪一对方法 —— "document" / "workbook"
    soffice_filter: str       # LibreOffice：--convert-to pdf:<这个过滤器>
    # ⑩c 的 key —— 「这一份该出几页」。一族只用一条口径，所以在这里声明一次；
    # 能核对页数的规则顺带核对。
    page_rule: str
    tip: str                  # 悬浮说明
    # ── ⭐ 2026-10-09 新增（用户问「引擎梯队」时收的口）────────────────────
    # family：**谁伺候我** —— ⑪ 表里 family 相同的引擎才来转我（"office" / "cad" / "btw"）。
    #         加一族格式 ＝ ⑩ 表加一行 + ⑪ 表加一条引擎，别处一个字不用动。
    # platforms：这个格式在**哪些平台**上转得了 —— 界面文案 / 列队闸门 / 回归断言都读它。
    #         空元组是错的（没有哪个格式两边都转不了）→ assert_format_complete 当场炸。
    # ⚠️ 这两个字段有默认值，是为了让 `_replace(...)` 与回归里手搓的临时行少写几个字；
    #    **⑩ 正式那两行必须显式写出来**（回归 [13a] 盯着）。
    family: str = "office"
    platforms: tuple[str, ...] = ("mac", "win")


CONVERT_FORMATS: tuple[ConvFormat, ...] = (
    ConvFormat(
        key="word", label="Word", exts=(".doc", ".docx", ".docm", ".rtf"),
        mac_app="Microsoft Word",
        # ⚠️ 不要再改回 `set d to open file name …` —— 那是 2026-10-09 的事故写法，
        #    见下面 ⑩a 的注释。正确姿势是「open → 等它真把文档交出来
        #    （活动文档可能还是 missing value）→ 存 → 关」。
        # ⚠️ `activate` 已经不在这里了 —— 它由 build_mac_script() 提到**循环外面**，
        #    整批只激活一次（每份都 activate 一次正是「窗口乱跳 + 反复问授权」的来源）。
        mac_open=(
            "set beforeCount to (count of documents)\n"
            "open srcHFS\n"
            "set d to missing value\n"
            "repeat 40 times\n"
            "  delay 0.25\n"
            "  try\n"
            "    if (count of documents) > beforeCount then "
            "set d to document (beforeCount + 1)\n"
            "  end try\n"
            "  if d is not missing value then exit repeat\n"
            "end repeat\n"
            "if d is missing value then error "
            "\"Word 没能打开这份文档（多半是弹了对话框在等人点）\""
        ),
        mac_export="save as d file name dstHFS file format format PDF",
        mac_close="close d saving no",
        win_app="Word.Application", win_kind="document",
        soffice_filter="writer_pdf_Export", page_rule="flow",
        tip="Word 文档整份转成一份 PDF（按它自己的分页一页一页）。",
        family="office", platforms=("mac", "win"),
    ),
    ConvFormat(
        key="excel", label="Excel", exts=(".xls", ".xlsx", ".xlsm"),
        mac_app="Microsoft Excel",
        # 与 Word 同一条思路；差别只在 Excel 的词典写法（workbook / filename）
        # 和它打开时要用 HFS 路径（`workbook file name` 那条历来只认冒号形式）。
        # ⚠️ 同上，`activate` 由 build_mac_script() 提到循环外面，整批只来一次。
        mac_open=(
            "set beforeCount to (count of workbooks)\n"
            "open workbook workbook file name srcHFS\n"
            "set wb to missing value\n"
            "repeat 40 times\n"
            "  delay 0.25\n"
            "  try\n"
            "    if (count of workbooks) > beforeCount then "
            "set wb to workbook (beforeCount + 1)\n"
            "  end try\n"
            "  if wb is not missing value then exit repeat\n"
            "end repeat\n"
            "if wb is missing value then error "
            "\"Excel 没能打开这份工作簿（多半是弹了对话框在等人点）\""
        ),
        mac_export="save as wb filename dstHFS file format PDF file format",
        mac_close="close wb saving no",
        win_app="Excel.Application", win_kind="workbook",
        soffice_filter="calc_pdf_Export", page_rule="sheet",
        tip="Excel 整个文件转成一份 PDF —— **底下每个工作表出成一页**。\n"
            "打印区域 / 缩放按文件里存的来，所以每个子表格正好一张纸。",
        family="office", platforms=("mac", "win"),
    ),
    # ── ⭐ Visio 流程图（2026-10-23 用户点名要的第五种格式）──────────────────
    # ⚠️ 它**自己一族**（family="visio"），不并进 office：Visio 常常压根没装，
    #    并进 office 会让「这台机器没 Visio」变成「Office 那一族全军覆没」，
    #    把同一批里的 .docx 也一起拖下水；⑪ 表的引擎名字也会变成一句假话
    #    （「Microsoft Word / Excel / Visio（本机）」—— 那是三个程序。
    # ⚠️ 与 CAD / BarTender 同一个写法：Office 那几个字段全空，
    #    「怎么转」在 ⑪ 表的 visio 引擎里（见 ⑩e）。
    ConvFormat(
        key="visio", label="Visio",
        # .vsd/.vsdx/.vsdm 是图纸；.vst/.vstx/.vstm 是它的模板（一样能导成 PDF）。
        # ⚠️ .vss / .vssx 是**形状库**（模具），不是能出图的东西，故意不收 ——
        #    收进来只会让人把一堆模具当成图纸拖进来。.vswx 同理（网页绘图）。
        exts=(".vsd", ".vsdx", ".vsdm", ".vst", ".vstx", ".vstm"),
        mac_app="", mac_open="", mac_export="", mac_close="",
        win_app="", win_kind="", soffice_filter="",
        page_rule="visio",
        tip="Visio 流程图交给本机装的 Visio 导出 PDF —— 只读打开、导出、不保存关掉。\n"
            "⭐ 走 Visio 自己的「导出为 PDF」，**不经过打印**：不依赖本机装没装 PDF 打印机，\n"
            "　 也不会被打印驱动改掉尺寸或边距。\n"
            "⭐ 纸张 / 缩放 / 页数**都用图纸里自己存的**（跟 Word / Excel 一个口径）。\n"
            "⚠️ 需要这台 Windows 电脑装有 Microsoft Visio；Mac 上没有它。",
        family="visio", platforms=("win",),
    ),
    # ── ⭐ 只在 Windows 上成立的两种（2026-10-09 用户点名要的）──────────────
    # ⚠️ 这两行的 Office 那几个字段**全空**（它们不是 Office 文档，不走三步曲，
    #    LibreOffice 也开不了它们）—— `assert_format_complete` 盯着「要么全空、要么全填」。
    # ⚠️ 它们的「怎么转」不在本行，而在 ⑪ 表里同族的引擎（cad / bartender）。
    ConvFormat(
        key="cad", label="CAD",
        # ⚠️ .dwg / .dxf 是正常图纸；.dwt（模板）/ .dws（标准文件）也是 CAD 自己
        #    能打开的东西，用户说「能囊括就都囊括起来」，所以一并认下。
        #    真出了意外（某一份打不开）会在「说明」列里点名，不会静默漏掉。
        exts=(".dwg", ".dxf", ".dwt", ".dws"),
        mac_app="", mac_open="", mac_export="", mac_close="",
        win_app="", win_kind="", soffice_filter="",
        page_rule="own",
        tip="CAD 图纸交给本机装的 CAD 出 PDF —— 打开图纸、出图到 PDF、不保存关掉。\n"
            "⭐ 布局 / 纸张 / 比例**都用图纸里自己存的**（跟 Word / Excel 一个口径）。\n"
            "⭐ 认程序**靠名字不靠位置**（COM）：装了浩辰 CAD 就用它（见 ⑩d 品牌表）。\n"
            "⚠️ 需要这台 Windows 电脑装有浩辰 CAD；Mac 上没有它。",
        family="cad", platforms=("win",),
    ),
    ConvFormat(
        key="btw", label="BarTender",
        exts=(".btw",),
        mac_app="", mac_open="", mac_export="", mac_close="",
        win_app="", win_kind="", soffice_filter="",
        page_rule="own",
        tip="BarTender 标签模板交给本机的 BarTender 导出 —— 一份 .btw 出**两个文件**：\n"
            "一个标签图片（JPG）+ 一份标签 PDF（走 BarTender 自己的导出接口，不打印、\n"
            "不旋转不调整，模板里什么尺寸就导什么尺寸）。\n"
            "⚠️ 需要这台 Windows 电脑装了 BarTender（装时要勾「Automation」组件）；Mac 上没有它。",
        family="btw", platforms=("win",),
    ),
)

# 收件时认的全部扩展名（由上面那张表派生，别另写一份）
CONVERT_EXTS: tuple[str, ...] = tuple(sorted({e for f in CONVERT_FORMATS for e in f.exts}))


# ══════════════════════════════════════════════════════════════════════════
# ⑩a 「族」表 —— 一种格式归谁管
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ 为什么要这一层（2026-10-09 用户问「引擎梯队」时收的口）：
#    以前 ⑪ 表是一根**全局梯队**（office_mac → office_win → libreoffice），
#    「谁伺候哪种格式」这件事**根本没写在任何地方** —— 全靠探测函数里藏一句
#    `if sys.platform != "darwin": return False`。于是：
#      · 表看起来像「Mac 版里装着 Windows 引擎」（其实跑不到，但读的人会这么以为）；
#      · 真加了 CAD / BarTender 就会出事：Excel 失手后梯队会把 CAD 交给 LibreOffice 试。
#    现在把「谁管谁」明写出来：「族」= 一组**互相可以顶替**的引擎。
#      · 加一种格式 ＝ ⑩ 表那一行写个 family；⑪ 表里同族的引擎自动认领；
#      · 兜底**只在同族内**换（Mac 的 Office 失手 → 同族的 LibreOffice，
#        绝不会跳去 AutoCAD）；
#      · 平台不符的引擎**根本不进梯队**（不是「进去了再判 False」）。
ENGINE_FAMILIES: dict[str, str] = {
    "office": "Word / Excel 这类 Office 文档",
    "visio": "Visio 流程图",
    "cad": "CAD 图纸",
    "btw": "BarTender 标签模板",
}

# 体检表里用的**短名**（长句留给正文与提示）。
# ⚠️ 表格里一列只能这么宽：`Word / Excel 这类 Office 文档` 占 29 列，
#    放进清单会把「用哪个程序」那一栏挤到第二行，整片对齐就散了。
FAMILY_SHORT: dict[str, str] = {
    "office": "Office 文档",
    "visio": "Visio 流程图",
    "cad": "CAD 图纸",
    "btw": "BarTender 标签",
}

# 某一族**靠哪个程序**（只用在「给人看的文案」里）。
# ⚠️ 别拿它去启动进程 —— 「怎么转」是 ⑪ 表的事，这里纯粹是说法。
# ⚠️ CAD 那一族的名字**必须跟 ⑩d 品牌表一致**：表里只剩浩辰了，这里也就写一家。
#    加一家时两处一起改 —— 回归里有一条断言盯着它俩不许各说各的。
FAMILY_MAKER: dict[str, str] = {
    "office": "Microsoft Office",
    "visio": "Microsoft Visio",
    "cad": "浩辰 CAD",
    "btw": "Seagull Scientific BarTender",
}


def family_maker(family: str) -> str:
    return FAMILY_MAKER.get(family, family)


def _maker_of(fmt: ConvFormat) -> str:
    """这一行靠哪个程序转（文案用）——⑩ 表写了 Mac 应用名就用它，否则看族。"""
    return fmt.mac_app or family_maker(fmt.family)

# 某一族的程序**这台机器上没有**时要说的话（列队闸门用）。
# ⚠️ 别笼统地写「去装一个」：CAD / BarTender 在 macOS 上**根本装不了**，
#    把人支使去找一个不存在的东西比不说还糟。分平台说清楚「去哪台机器上做」。
FAMILY_NO_ENGINE: dict[str, str] = {
    "office": ("没找到可用的 {kinds} —— 转换要靠本机装的 Office。\n\n"
               "装好 Microsoft {kinds} 之后，点一下「重新检测引擎」。"),
    "visio": ("这台机器上没有能转「{what}」的程序 —— 要靠本机装的 Microsoft Visio。\n\n"
              "Visio 是单独买、单独装的（Office 套件里不一定带它）；\n"
              "macOS 上没有 Visio 桌面版，请在装了它的 Windows 电脑上转。"),
    # ⚠️ `{maker}` 由 family_maker() 填（＝⑩d 品牌表里那几家的名字），
    #    别再在这里手写一遍品牌名 —— 表一改文案就漂。
    "cad": ("这台机器上没有能转「{what}」的程序 —— 要靠本机装的 CAD（{maker}）。\n\n"
            "macOS 上没有它们的桌面版；请在装了它的 Windows 电脑上转。"),
    "btw": ("这台机器上没有能转「{what}」的程序 —— 要靠本机装的 BarTender。\n\n"
            "macOS 上没有 BarTender 桌面版；请在装了它的 Windows 电脑上转。"),
}


def family_label(family: str) -> str:
    """族的「人话」名字；表里漏了也不许抛异常（只是文案）。"""
    return ENGINE_FAMILIES.get(family, family)


# 现在这一版适用的转换规则 —— 界面上直接显示，用户知道「按什么在转」。
# ⭐ 2026-10-10 改口径：CAD **也**改成「按图纸里存的来」了。
#    原来写的是「CAD 按工具内固定的出图参数来」—— 因为旧版走 accoreconsole 的 -PLOT，
#    要逐行喂 17 个答案（纸型/方向/比例全由工具定死）。现在换成 COM 自动化：
#    只把**出图设备**换成 PDF 打印机（图纸里存的多半是实体绘图仪，照它打出不了 PDF），
#    布局、纸张、比例、打印区域**一律用图纸里自己存的** —— 跟 Word / Excel 同一个口径，
#    也就是用户那句「他们内部已经把文件编辑好了，导出什么就是什么，不需要人为微调」。
CONVERT_RULE_NOTE = ("每个文件转一份 PDF —— 都按文件里存的打印设置来"
                     "（CAD 只把出图设备换成 PDF 打印机，其余照图纸里存的）。")


# ══════════════════════════════════════════════════════════════════════════
# ⑩c 分页规则表 —— 「一份文件该出几页」
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ 2026-10-09 用户给的规律（配 WI-2019 实例，13 个子表格 ↔ 13 页，一一对应）：
#    · **Excel：底下每个子表格（工作表）= 一张纸一页**。用户原话「它大概来说就是
#      底下有子表格，那每一个子表格相当于打印一页出来」；每个子表内容都不多、
#      只占一页，所以整本导出就是「一个子表一页」。
#      Excel 的导出**本来就是整本按这个口径出的**（Mac 的 `save as … PDF` 与
#      Windows 的 `ExportAsFixedFormat` 都是整本，不是只导活动表 —— 已核实：
#      Mac 上想只导一个表反而做不到），所以规则天然成立；这里把它**声明出来**，
#      顺带核对页数，免得哪天子表内容溢出了没人发现。
#    · **Word：正常文档，按它自己的分页一页一页。**
#
# 加格式 ＝ 在 ⑩ 表那行写一个 page_rule 的 key；要新口径就往这里加一行。
# `expect(原件)` 能给得出「应该几页」就返回它（转换后拿去核对），给不出返回 None
# （＝这道核对不做，绝不瞎猜）。它**只读原件**，绝不改动任何东西。

_XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _xlsx_sheet_count(path: Path) -> int | None:
    """.xlsx / .xlsm 里有几个**可见**工作表 —— Excel 导 PDF 只导可见的那些。

    ⚠️ 老 .xls（二进制）读不了 → None（这一份就不做这项核对，而不是猜一个）。
    ⚠️ 只读 zip 里的 workbook.xml，**不写、不解压**，原件一个字节不动。
    ⚠️ 任何异常一律 None —— 这只是个核对，绝不能因为它把本来能转的搞成不能转。
    """
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return None
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = root.find(_XLSX_NS + "sheets")
        if sheets is None:
            return None
        return sum(1 for sh in sheets
                   if (sh.get("state") or "visible") == "visible")
    except Exception:  # noqa: BLE001
        return None


class PageRule(NamedTuple):
    """⑩c 一条分页规则。

    label ：给人看的一句话（进界面的规则提示）
    expect：能给「应该几页」就给（转换后核对用）；给不出就 None
    """

    key: str
    label: str
    expect: Callable[[Path], int | None] | None = None


PAGE_RULES: dict[str, PageRule] = {
    "flow": PageRule("flow", "Word 按自己的分页一页一页"),
    "sheet": PageRule("sheet", "Excel 每个工作表一页", _xlsx_sheet_count),
    # CAD / BarTender：页数由它们自己排（图纸几个布局、标签排几联），
    # 我们**判不了**「该几页」→ expect 留空，只声明口径、不做核对（绝不瞎猜）。
    "own": PageRule("own", "图纸 / 标签按它自己排的页"),
    # Visio：图纸里画的每个绘图页 = 一页 PDF（背景页不单独出，跟 Visio 自己导出一致）。
    # ⚠️ 同上不给 expect：.vsd 是二进制容器、.vsdx 的页数还牵扯背景页，估错了
    # 会变成「假告警」—— 宁可只声明口径。
    "visio": PageRule("visio", "Visio 每个绘图页一页"),
}


def page_rule_of(fmt: ConvFormat) -> PageRule | None:
    return PAGE_RULES.get(fmt.page_rule)


def page_rule_label() -> str:
    """本次平台能用到的分页规则，拼成一句话（给界面提示 / 使用说明用）。"""
    seen: list[str] = []
    for f in platform_formats():
        r = page_rule_of(f)
        if r and r.label not in seen:
            seen.append(r.label)
    return "；".join(seen)


def page_rule_report(fmt: ConvFormat, src: Path, pages: int) -> str:
    """转出来的页数**不符合这条规则**时给一句人话；符合 / 判不了 → 空串。

    ⭐ 只在**不符**时出声：符合就是默认预期，不必报（用户嫌啰嗦）。
    ⚠️ src 传**原件**（不是临时副本）；核对只是读，不阻断、不改状态。
    """
    rule = page_rule_of(fmt)
    if rule is None or rule.expect is None or pages <= 0:
        return ""
    want = rule.expect(src)
    if not want or want == pages:
        return ""
    return (f"{rule.label}：这份有 {want} 个工作表，转出来却是 {pages} 页 —— "
            f"对不上，多半是有子表的内容超过一页（或空表没出页），打开原件看一眼。")


def format_of(path: Path) -> ConvFormat | None:
    """⑩ 这份文件属于哪种格式？认不出返回 None（不当成要转的东西）。"""
    ext = Path(path).suffix.lower()
    for f in CONVERT_FORMATS:
        if ext in f.exts:
            return f
    return None


def format_by_key(key: str) -> ConvFormat | None:
    for f in CONVERT_FORMATS:
        if f.key == key:
            return f
    return None


# ── 「这一版支持什么」的说法，一律从 ⑩ 表 + **当前平台**派生 ──────────────
#
# ⭐ 为什么不直接在界面文案里写「Word / Excel」：加一种格式就要去改七八处文案，
#    改漏的那一处从此对用户说假话（比如在一个转不了它的平台上写「支持某某」）。
#    收口成函数之后，加格式只动 ⑩ 表那一行。
# ⚠️ 收件口径（CONVERT_EXTS）**刻意不按平台过滤** —— 认得出才给得出
#    「这份本机转不了，去有它的那台机器上转」这种明确反馈；
#    悄悄不认反而让人以为文件坏了 / 工具漏读了。
#    2026-10-09 起这条**真的看得出差别了**：CAD / BarTender 只在 Windows 上成立，
#    而 Mac 版照样收得进队列（收得进 = 才说得出那句明确的话）。
def running_platform() -> str:
    """当前跑在哪一边。只看「这一版支持什么」，与引擎探测无关。"""
    return "mac" if sys.platform == "darwin" else "win"


def platform_formats(platform: str | None = None) -> tuple[ConvFormat, ...]:
    """⑩ 表里**在这个平台上有对应程序**的那些行（`platforms` 说了算）。"""
    plat = platform or running_platform()
    return tuple(f for f in CONVERT_FORMATS if plat in f.platforms)


def convert_kind_words(platform: str | None = None) -> str:
    """界面文案里「支持哪些类型」的说法：Mac `Word / Excel`、Windows 再多 `CAD / BarTender`。"""
    return " / ".join(f.label for f in platform_formats(platform))


def family_kind_words(family: str, platform: str | None = None) -> str:
    """**只数某一族**的格式名 —— 引擎自己的名字用这个，别用上面那个。

    ⚠️ 引擎的名字里绝不能出现别族的格式：`office_win` 那条要是写成
       `Microsoft Word / Excel / CAD / BarTender（本机）`，就等于对用户说
       「这个引擎能转 CAD」—— 而它根本没这个本事。2026-10-09 加 CAD 时踩到的。
    """
    plat = platform or running_platform()
    return " / ".join(f.label for f in CONVERT_FORMATS
                      if f.family == family and plat in f.platforms)


# ══════════════════════════════════════════════════════════════════════════
# ⑪ 转换引擎表 —— 谁真的去执行转换。按顺序探测，第一个可用的胜出
# ══════════════════════════════════════════════════════════════════════════
#
# 一条引擎 = 「怎么判断它能不能用」+「怎么转一份」。
# 统一约定：run() 返回 "" 表示成功，返回非空字符串 = 失败原因（直接给人看）。
#
# 优先级按**保真度**排：本机 Office 最懂自家文件的打印设置，
# 所以先试它；没有才退到 LibreOffice（免装 Office，但分页口径是它自己那套）。

# 单份文件最多等多久 —— 引擎弹了对话框没人点时不能把整批挂死
CONVERT_TIMEOUT = 300

# LibreOffice 的常见安装位置（Mac / Win），PATH 上也找一遍
SOFFICE_CANDIDATES: tuple[str, ...] = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


def find_soffice() -> str | None:
    for c in SOFFICE_CANDIDATES:
        if Path(c).is_file():
            return c
    return shutil.which("soffice")


def _mac_app_path(app: str) -> Path:
    return Path("/Applications") / f"{app}.app"


def _probe_office_mac() -> bool:
    if sys.platform != "darwin":
        return False
    # 只看**在 macOS 上有对应应用**的格式（万一某行 mac_app 是空的，跳过它）
    return any(_mac_app_path(f.mac_app).exists() for f in platform_formats("mac"))


# ── Windows 方言：⑩ 表的 `win_kind` 落到这两个函数 ───────────────────────
# 它们与 macOS 的 mac_open / mac_export / mac_close 是**同一套三步曲**，
# 只是换了种写法（COM 调用 vs AppleScript 语句）。表里加格式时，
# 这里要么复用已有的 kind，要么补一个新 kind —— 补漏由 assert_format_complete 拦。
#
# ⭐⭐ 2026-10-23 治「假失败」（用户实测：一批 3 份报「失败 2 份」，
#    可那两份 PDF 全都好好转出来了、内容也正常）。
#    根因不在转换，在**收尾**：Office 忙的时候 `doc.Close()` 会抛
#    RPC_E_CALL_REJECTED（「被调用方拒绝调用」0x80010001 / 0x8001010A）——
#    这时候 SaveAs 早就把文件写完了，异常却一路冒到批量口，被记成「这份没转成」。
#    为什么常常只有第一份成功：第一个文档关闭时实例还闲着，从第二份起 Word
#    正忙着做后台保存 / 恢复，Close 就接连被拒 —— 正好对上「3 份成 1 份」的形状。
#    三道防线，缺一条假失败就会回来：
#      ① `Close()` 失败**一律吞掉**（收尾不是成败判据）—— 见 _quiet_close
#      ② Open / SaveAs / Export 撞上「Office 忙」**退一步重试** —— 见 _com_retry
#      ③ 万一还是抛了：只要**产物确实新生成了**就当成功 —— 见 _win_salvaged
#    ⚠️ 别把 ③ 改回「有异常就算失败」——那就是这个 bug 本身。
#    ⚠️ 也别把 ③ 简化成「文件存在就算成功」—— 覆盖那次原件旁本来就有同名文件。

# 「Office 忙着呢，等一下再来」这一类 HRESULT —— 撞上它重试基本都能过，不该算失败。
COM_BUSY_HRESULTS: frozenset[int] = frozenset({
    0x80010001,   # RPC_E_CALL_REJECTED —— Word 上最常见的那条
    0x8001010A,   # RPC_E_SERVERCALL_RETRYLATER —— 「忙，稍后重试」
    0x800AC472,   # Excel 版的「被调用方拒绝调用」
    0x800706BA,   # RPC 服务器不可用（Office 正在退出 / 重启，下一跳通常就好）
})

# `DISP_E_BADPARAMCOUNT` —— 「参数个数不对」。
# ⚠️ 它**不是**「东西坏了」，而是「这一版不认这套写法」：重试没用、换参数个数才管用。
#    2026-10-23 Visio 三份 .vsd 全报的就是它（少给了 Intent / PrintRange 两个必填参数）。
#    报错时也特意不优先把它甩给用户 —— 那句话对用户没有任何解释力。
COM_BADPARAMCOUNT = 0x8002000E


def _exc_hresult(exc: BaseException) -> int | None:
    """从 pywin32 的 com_error 里挖出 HRESULT（挖不到 → None）。

    ⚠️ pywin32 把 hresult 放在 `args[0]`（**有符号**整数），新版另有 `.hresult` 属性；
       两处都要试，否则新老版本各漏一半，重试就形同虚设。
    """
    h = getattr(exc, "hresult", None)
    if not isinstance(h, int) and exc.args and isinstance(exc.args[0], int):
        h = exc.args[0]
    return (h & 0xFFFFFFFF) if isinstance(h, int) else None


def _com_busy(exc: BaseException) -> bool:
    """这个错是不是「Office 忙，等下再来」那种（重试有意义的那种）。"""
    return _exc_hresult(exc) in COM_BUSY_HRESULTS


def _com_retry(call: Callable[..., object], *args, **kwargs) -> object:
    """调一次 COM 接口；撞上「Office 忙」就退一步再试（最多 3 次，共约 1.2 秒）。

    ⚠️ **只**重试瞬时错误。「文件被占用」「没装 Office」这类重试也没用的，
       原样抛出去，交给 ⑪c 归类 —— 别在这里把真错误也拖成 3 倍耗时。
    """
    last: BaseException | None = None
    for i in range(3):
        try:
            return call(*args, **kwargs)
        except Exception as exc:                      # noqa: BLE001
            last = exc
            if i == 2 or not _com_busy(exc):
                raise
            time.sleep(0.4 * (i + 1))                 # 0.4 → 0.8 秒
    if last is not None:                              # 走不到；留着让静态检查满意
        raise last
    raise AssertionError("_com_retry 不该走到这里")


def _quiet_close(obj: object) -> None:
    """收尾关闭：**关不上不算失败**。

    ⚠️ 绝对不要在这里 raise —— 调用方（批量口）拿到的异常会被记成「这份没转成」，
       而文件其实早就写好了。用户报的那个「报失败、PDF 却是好的」就是这条。
    """
    if obj is None:
        return
    try:
        obj.Close(False)          # False = 不保存（只读打开的，本来也没改动）
    except Exception:             # noqa: BLE001
        pass


def _product_stamp(p: Path) -> tuple[int, int] | None:
    """产物在**动手之前**长什么样（mtime_ns, size）；不存在 → None。"""
    try:
        st = p.stat()
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def _looks_like_pdf(p: Path) -> bool:
    """这份文件现在是不是一份「像样的 PDF」（不是 0 字节 / 不是半截垃圾）。"""
    try:
        if p.stat().st_size <= 0:
            return False
        with open(p, "rb") as f:
            return b"%PDF" in f.read(1024)
    except OSError:
        return False


def _win_salvaged(dst: Path, before: tuple[int, int] | None) -> bool:
    """抛了异常，但这份 PDF **确实是这次转出来的**吗？

    ⚠️ 必须跟 `before` 比对：原件旁边可能**本来就有**一份同名产物（覆盖那一次），
       光看「文件存在」会把真失败也判成成功 —— 那比报假失败更糟。
    """
    now = _product_stamp(dst)
    return now is not None and now != before and _looks_like_pdf(dst)


# 「转成了、但收尾报了一句」的现场（最多留 20 条）。
# ⭐ 这些**不往界面上放**（用户不需要看，看见了反而以为出事了），
#    只给 `--check-convert` 用：下次再有人说「它说失败、可文件是好的」，
#    一句 `python seq_core.py --check-convert 某份.docx` 就能把这句原文打出来。
WIN_SALVAGED_NOTES: list[str] = []


def _note_salvaged(src: Path, exc: BaseException) -> None:
    """记一条「转成了、但收尾报了句错」的现场（只给自检用，见上面那个列表）。"""
    try:
        WIN_SALVAGED_NOTES.append(f"{Path(src).name}：{exc}"[:200])
        del WIN_SALVAGED_NOTES[:-20]          # 只留最近 20 条，别无限长
    except Exception:                         # noqa: BLE001
        pass


def _win_convert_document(app, src: Path, dst: Path) -> None:
    """COM：Word 文档 → PDF。"""
    doc = None
    try:
        doc = _com_retry(app.Documents.Open, str(src),
                         ReadOnly=True, AddToRecentFiles=False)
        _com_retry(doc.SaveAs, str(dst), FileFormat=17)     # 17 = wdFormatPDF
    finally:
        _quiet_close(doc)                                  # 见文件上方「假失败」那段


def _win_convert_workbook(app, src: Path, dst: Path) -> None:
    """COM：Excel 工作簿 → PDF。"""
    wb = None
    try:
        wb = _com_retry(app.Workbooks.Open, str(src), ReadOnly=True)
        _com_retry(wb.ExportAsFixedFormat, 0, str(dst))     # 0 = xlTypePDF
    finally:
        _quiet_close(wb)


WIN_OFFICE_RUNNERS: dict[str, Callable[[object, Path, Path], None]] = {
    "document": _win_convert_document,
    "workbook": _win_convert_workbook,
}


# ── Windows 方言：Office 起来之后统一要设的几个「应用级」属性 ──────────────
#
# ⭐ 为什么是一张表、而不是两行 setattr：**不同 Office 认的属性不一样**。
#    Word / Excel 有 DisplayAlerts（不关它，「是否保存」这类弹窗就会把流程挂住）。
#    将来若加一门别的 Office（认的是它自己那套成员名），往这张表加一行即可 ——
#    「设不上就跳过」比「设不上就整批失败」合理得多：这几个属性只影响弹窗行为，
#    设不上照样能转。谁认哪个由这张表定。
APP_INIT_PROPS: tuple[tuple[str, object], ...] = (
    ("Visible", False),        # 两边都认：后台干活，别弹主窗口
    ("DisplayAlerts", False),  # Word / Excel 专用
)


def _init_office_app(app: object) -> None:
    """把 APP_INIT_PROPS 逐个设一遍，**哪个设不上就跳过哪个**（理由见上表）。"""
    for name, value in APP_INIT_PROPS:
        try:
            setattr(app, name, value)
        except Exception:  # noqa: BLE001
            pass


# ══════════════════════════════════════════════════════════════════════════
# ⑪c 失败归类表 —— **macOS 与 Windows 共用同一套口径**
# ══════════════════════════════════════════════════════════════════════════
#
# 两个平台吐出来的报错长得完全不一样（AppleScript 给 -10004，
# COM 给「拒绝访问」/ HRESULT / Python 异常文本），但它们**要表达的其实是同一批事**。
# 与其在两个平台里各写一串 if，不如把「原始文本 → 一个稳定的 code」收在这里，
# 再由 code 决定给用户看什么、给什么补救入口（界面还会据此弹「打开设置」按钮）。
#
# ⚠️ 加一类失败 ＝ 往下面加一行，两个平台自动都认。
OFFICE_ERR_RULES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("no_permission", ("-10004", "权限违例", "not permitted", "not authorized",
                       "permission denied", "拒绝访问", "access is denied",
                       "0x80070005")),
    ("syntax", ("-2741", "syntax error", "expected end of line", "cannot compile",
                "语法错误")),
    # ⭐ 「叫起来了但没交出文档」—— 2026-10-09 的 -2753 就是这一类：
    #    报错文本长成「变量"d"没有定义」，看着像脚本写错，其实是 Office 没就绪 / 弹窗。
    #    末两条是我们自己在脚本里 error 出来的原话（等不到文档时），别删。
    ("no_doc", ("-2753", "没有定义", "-1728", "is not defined", "missing value",
                "can't get", "cannot get", "无效的引用", "没能打开")),
    ("no_app", ("无效的类字符串", "class not registered", "0x80040154",
                "found no application", "application isn't running",
                "找不到应用程序", "not installed", "没有安装")),
    ("busy", ("正在使用", "already open", "已打开", "被占用", "locked",
              "0x800a0046", "sharing violation")),
)

OFFICE_ERR_TITLE: dict[str, str] = {
    "no_permission": "系统没放行「自动化控制」",
    "syntax": "给 Office 的那段脚本没编译过",
    "no_doc": "Office 起来了，但没把文档交出来",
    "no_app": "本机没找到这个 Office 程序",
    "busy": "文件正被别的程序占用",
    "other": "转换失败",
}

# (平台, code) → 补救说明。{app} / {raw} 会被替换掉。
OFFICE_ERR_HINTS: dict[tuple[str, str], str] = {
    ("mac", "no_permission"): (
        "macOS 的「自动化」授权没放行这次操作。按顺序试，每步做完**重开一次本工具**：\n"
        "① 系统设置 → 隐私与安全性 → 自动化：给「终端」勾上允许控制 "
        "Microsoft Word / Excel；\n"
        "② 系统设置 → 隐私与安全性 → 完全磁盘访问权限：把 Microsoft Word / Excel 勾上。\n\n"
        "⚠️ 还有一种情况：本工具**不是你自己双击启动的**（比如从别的程序里代开）。\n"
        "　 那种窗口拿不到这个授权，请关掉、自己双击启动器再跑一次。"),
    ("mac", "syntax"): (
        "这段脚本是工具自己生成的，跟你的文件无关 —— 把下面原文发给维护的人：\n{raw}"),
    ("mac", "no_doc"): (
        "Word / Excel 被叫起来了，但**没把文档交给我们**（不是你的文件有问题）。\n"
        "按顺序试：\n"
        "① 切到 Word / Excel 窗口看一眼 —— 是不是弹着对话框在等人点"
        "（「启用编辑」/ 只读 / 修订 / 登录激活）；点掉它再重试；\n"
        "② 把你自己已经打开的 Word / Excel 窗口全关掉，再重试。"),
    ("mac", "no_app"): "这台电脑上没装 {app} —— 装好它再试。",
    ("mac", "busy"): "文件正被别的程序打开着（可能就是你自己的 {app}）。\n"
                     "建议先在那边关掉这份文件，再回来重试。",
    ("win", "no_permission"): (
        "Office 拒绝了这次自动化。常见两条：\n"
        "① Office 正卡在一个对话框上（宏 / 只读 / 修订）—— 手动打开这份文件看一眼；\n"
        "② 公司策略禁用了自动化 —— 问一下 IT 能不能给 Office 放行。"),
    ("win", "no_doc"): (
        "Office 起来了但没交出文档。多半是它卡在对话框上（只读 / 宏 / 修订），\n"
        "手动打开这份文件看一眼、点掉提示，再重试。"),
    ("win", "no_app"): "这台电脑上没装 {app}（或它注册坏了）—— 装好 / 修复它再试。",
    ("win", "busy"): "文件正被别的程序打开着（可能就是 {app} 自己）。\n"
                     "建议先在那边关掉这份文件，再回来重试。",
}


def classify_office_error(raw: str) -> str:
    """原始报错文本 → 稳定的 code（Mac / Windows 同一套规则，命中第一条即返回）。"""
    low = (raw or "").lower()
    for code, pats in OFFICE_ERR_RULES:
        if any(p.lower() in low for p in pats):
            return code
    return "other"


def office_error_note(platform: str, fmt: ConvFormat, raw: str) -> str:
    """统一的失败文案：抬头（含 code 对应的头一句）+ 补救说明 + 原始报错尾巴。

    ⭐ 两个平台的**所有**失败都必须从这里出 —— 这样「同一种病」在两个平台上
       说出来的话是一致的，界面也能靠同一个 code 决定弹什么补救入口。
    """
    app = fmt.mac_app if platform == "mac" else fmt.win_app
    code = classify_office_error(raw)
    head = f"{app} 转换失败：{OFFICE_ERR_TITLE.get(code, OFFICE_ERR_TITLE['other'])}"
    hint = OFFICE_ERR_HINTS.get((platform, code), "")
    hint = hint.replace("{app}", app)
    hint = hint.replace("{raw}", (raw or "").strip()[:200] or "（没有报错信息）")
    tail = (raw or "").strip()[:300]
    parts = [head]
    if hint:
        parts.append(hint)
    if tail and tail not in hint:
        parts.append("原始报错：" + tail)
    return "\n".join(parts)


def assert_format_complete(fmt: ConvFormat) -> None:
    """⑩ 表的一行是不是「自己说得通」—— 缺哪一步、哪边自相矛盾都当场报清楚。

    ⚠️ 为什么要这道自检：漏一步的后果是「转出来的 PDF 不对」或「脚本编译不过」，
       而报错文本只会说 -2741 / 权限违例，指向别处。宁可在这里当场炸。

    ⭐ 2026-10-09 起分两步看：
       ① **通用**（哪一族都要）：family 认得、platforms 非空且合法、page_rule 认得、tip 有字；
       ② **按族**：
          · `office` 族 —— 「macOS 三步曲」与 `win_app`/`win_kind` 必须跟 `platforms`
            **对齐**（说 Mac 上能转就得有应用名与三步；说不转就得全空）；
          · 别的族（cad / btw）—— Office 那几个字段**必须全空**：
            它们描述的是「怎么用 Office 转」，对 CAD / BarTender 没有意义，
            填了就是表自己矛盾。这些行「怎么转」在 ⑪ 表里同族的引擎身上。
    """
    if fmt.family not in ENGINE_FAMILIES:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行的 family={fmt.family!r} 不在 ⑩a 族表里"
                         f"（可选：{'、'.join(ENGINE_FAMILIES)}）")
    if not fmt.platforms:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行没写 platforms —— "
                         f"没有哪个格式是两边都转不了的")
    bad_plat = [p for p in fmt.platforms if p not in ("mac", "win")]
    if bad_plat:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行的 platforms 里有认不出的：{'、'.join(bad_plat)}"
                         f"（只认 mac / win）")
    if fmt.page_rule not in PAGE_RULES:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行的 page_rule={fmt.page_rule!r} "
                         f"不在 ⑩c 分页规则表里（可选：{'、'.join(PAGE_RULES)}）")
    if not fmt.tip.strip():
        raise ValueError(f"⑩ 表里 {fmt.key} 这行没有 tip（界面上会是个空悬浮框）")

    steps = {s: getattr(fmt, f"mac_{s}", "").strip() for s in OFFICE_STEPS}
    filled = [s for s, v in steps.items() if v]

    if fmt.family != "office":
        # 非 Office 族：那几个字段一个都不该有。
        ugly = ((["mac_app"] if fmt.mac_app else [])
                + [f"mac_{s} ({OFFICE_STEP_LABELS.get(s, s)})" for s in filled]
                + (["win_app"] if fmt.win_app else [])
                + (["win_kind"] if fmt.win_kind else [])
                + (["soffice_filter"] if fmt.soffice_filter else []))
        if ugly:
            raise ValueError(
                f"⑩ 表里 {fmt.key} 这行是 «{fmt.family}» 族（不靠 Office 转），"
                f"那几个 Office 字段就不该填 —— 现在填了：{'、'.join(ugly)}")
        return

    # ── office 族：三步曲 / win_kind 必须跟 platforms 对齐 ──────────────
    if "mac" in fmt.platforms:
        if not fmt.mac_app:
            raise ValueError(f"⑩ 表里 {fmt.key} 这行说 Mac 上能转，却没有 macOS 应用名")
        if len(filled) != len(OFFICE_STEPS):
            missing = [s for s in OFFICE_STEPS if not steps[s]]
            names = "、".join(OFFICE_STEP_LABELS.get(s, s) for s in missing)
            raise ValueError(f"⑩ 表里 {fmt.key} 这行少了 macOS 的这几步：{names}")
    elif fmt.mac_app or filled:
        names = "、".join(
            (["mac_app"] if fmt.mac_app else [])
            + [f"mac_{s}" for s in filled])
        raise ValueError(
            f"⑩ 表里 {fmt.key} 这行 platforms 里没有 mac（＝ Mac 上转不了），"
            f"macOS 那几项就不该填 —— 现在填了：{names}")
    if "win" in fmt.platforms:
        if not fmt.win_app:
            raise ValueError(f"⑩ 表里 {fmt.key} 这行说 Windows 上能转，却没有 Windows 应用名")
        if fmt.win_kind not in WIN_OFFICE_RUNNERS:
            raise ValueError(f"⑩ 表里 {fmt.key} 这行的 win_kind={fmt.win_kind!r} "
                             f"没有对应实现（可选：{'、'.join(sorted(WIN_OFFICE_RUNNERS))}）")
    elif fmt.win_app or fmt.win_kind:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行 platforms 里没有 win，"
                         f"Windows 那几项就不该填")
    if not fmt.soffice_filter:
        raise ValueError(f"⑩ 表里 {fmt.key} 这行没有 LibreOffice 的过滤器 —— "
                         f"Office 那两条都失手时就没法兜底了")


def _indent_block(text: str, pad: str) -> str:
    """把一段多行的步骤统一缩进 —— 表里的写法只关心内容，缩进交给这里。

    ⚠️ 每个**非空行**都加 pad（空行原样留着），这样拼出来的脚本层次一致，
       以后对着 `--check-convert` 打出来的原文排查时，眼睛能跟得上。
       AppleScript 不在乎缩进，所以这只是可读性，不是语法要求。
    """
    return "\n".join(pad + ln if ln.strip() else ln for ln in text.splitlines())


def build_mac_script(fmt: ConvFormat) -> str:
    """按 OFFICE_STEPS 的顺序，把 ⑩ 表里这个格式的三句拼成**一段整批 AppleScript**。

    脚本结构（一个格式一段，一次转完整批）：

        on run argv                 ← argv = src1, dst1, src2, dst2 …
          逐对转成 HFS 路径          ← 在 tell 块外做，见 _MAC_OSA_HEADER
          tell application "<字面量>"
            activate                ← 整批只激活一次（循环外面）
            repeat 每一份
              换上这一份的 srcHFS / dstHFS
              mac_open → mac_export → mac_close   ← 顺序由 OFFICE_STEPS 定
              记一行 "N|OK"
              on error → 记一行 "N|ERR|号|话"      ← 这份崩了不影响后面的
          return 每行一份的结果

    ⭐ app 名是**拼进字面量**的（`tell application "Microsoft Word"`），
       而不是 `tell application <变量>` —— 后者 AppleScript 编译期认不出词典，
       整段 -2741，一份都转不出来（2026-10-08 的真实事故，理由见上面的长注释）。
       拼装写法让这个错误模式在结构上无法复现。
    ⭐ 一份失败只记一行错误、**继续转后面的** —— 这是整批模式带来的额外好处：
       旧的「一份一段脚本」写法里，一份崩了整批就停在那儿。
    """
    # ⚠️ 在 macOS 上没有对应应用的格式**不该走到这里** ——
    #    调用方（_run_office_mac_many）会先把它挡掉并给出专门说明。
    #    真漏过来了就当场炸，别去拼一段 `tell application ""` 的怪脚本。
    if not fmt.mac_app:
        raise ValueError(f"⑩ 表里 {fmt.key} 在 macOS 上没有对应的应用，"
                         f"不该调 build_mac_script")
    assert_format_complete(fmt)
    lines = [
        _MAC_OSA_HEADER.rstrip("\n"),
        f'  tell application "{fmt.mac_app}"',
        "    activate",
        "    set out to {}",
        "    repeat with i from 1 to nFiles",
        "      set srcHFS to item i of srcList",
        "      set dstHFS to item i of dstList",
        "      try",
    ]
    for step in OFFICE_STEPS:
        lines.append(_indent_block(getattr(fmt, f"mac_{step}"), "        "))
    lines += [
        '        set end of out to (i as text) & "|OK"',
        "      on error errMsg number errNum",
        '        set end of out to (i as text) & "|ERR|" '
        '& (errNum as text) & "|" & my neat(errMsg)',
        "      end try",
        "    end repeat",
        "  end tell",
        _MAC_OSA_FOOTER.rstrip("\n"),
        "",
    ]
    return "\n".join(lines)


def _parse_mac_batch_output(text: str, n: int) -> list[str | None]:
    """把脚本返回的「一行一份」结果解析成 n 项。

    "" = 这份转好了；非空 = 这份的失败原因；None = 脚本压根没给出这份的结果
    （整段没跑起来，或半路被打断）—— 由调用方补一句能看懂的话。
    """
    out: list[str | None] = [None] * n
    for ln in (text or "").splitlines():
        ln = ln.strip()
        if not ln or "|" not in ln:
            continue
        parts = ln.split("|", 3)
        try:
            idx = int(parts[0].strip())
        except ValueError:
            continue
        if not (1 <= idx <= n):
            continue
        if len(parts) > 1 and parts[1].strip().upper() == "OK":
            out[idx - 1] = ""
            continue
        num = parts[2].strip() if len(parts) > 2 else ""
        msg = parts[3].strip() if len(parts) > 3 else ""
        note = msg or "转换失败（脚本没给出原因）"
        if num and num not in ("0", "-0"):
            note = f"{note} ({num})"            # 号挂在尾巴上，⑪c 的归类靠它认
        out[idx - 1] = note
    return out


# macOS 上「这个格式根本没有对应应用」时说的话。
# ⚠️ 别写成「没装 xxx，装好再试」—— 那等于把人支使去找一个 Mac 上不存在的东西。
#    这台机器上真装不了，就把话说明白：去有它的那台机器上转。
#    （表里现在没有这样的格式 —— Word / Excel 两边都有；这是给「只在 Windows 上有的
#      行业软件格式」留的兜底话术，由 selftest 的合成格式用例盯着。）
_NO_MAC_APP_NOTE = (
    "{label} 这类文件要靠它自己的那个程序（{maker}）导出 PDF，"
    "而 macOS 上没有它 ——\n本机转不了，请在装了它的 Windows 电脑上转。"
)


# macOS：让 Word / Excel 自己「另存为 PDF」——保真度最高，也尊重文件自带的打印设置。
# ⚠️ 首次使用 macOS 会弹一次自动化授权（给「终端」），点允许即可；
#    Office 是沙箱应用，若仍报 -10004 就是「完全磁盘访问权限」那条（见 ⑪c 的提示）。
# ⚠️ 失败一律交给 office_error_note() —— 别在这里自己拼文案，否则 Mac / Win 会跑偏。
def mac_unavailable_note(fmt: ConvFormat) -> str:
    """⑩ 表这一行在 macOS 上**根本没有对应程序**时要说的话。

    返回 ""＝没有这回事（这一行在 Mac 上转得了，照常去叫它）。
    ⭐ 判据是 `fmt.platforms`（**明写**的），不是「mac_app 是不是空」那种暗示。
    ⭐ 抽成纯函数是为了能回归：`_run_office_mac_many` 拿它当第一道门，
       回归则可以不打任何命令、直接检查「这句说得对不对」。

    ⚠️ 这只是**最后一道兜底**：真正的第一道门在列队闸门（同族没有引擎就不让排队，
       见 ⑫ 的 _cg_engine）。留在这里是因为「万一真漏过来」，也必须有句人话。
    """
    if "mac" in fmt.platforms:
        return ""
    return _NO_MAC_APP_NOTE.format(label=fmt.label, maker=fmt.mac_app or _maker_of(fmt))


def _run_office_mac_many(jobs: list[tuple[Path, Path]], fmt: ConvFormat,
                         timeout: int = CONVERT_TIMEOUT) -> list[str]:
    """macOS 的批量口：**一批同格式的文件，一次 osascript、一个应用会话**。

    返回与 jobs 等长的失败原因（"" = 这份转好了）。失败的是**单独某一份**时，
    整批照样往下转 —— 一份坏了不该拖着其余的全转不出来。
    """
    n = len(jobs)
    if n == 0:
        return []
    note = mac_unavailable_note(fmt)           # Mac 上压根没这个应用的那种格式
    if note:
        return [note] * n
    if not _mac_app_path(fmt.mac_app).exists():
        return [office_error_note("mac", fmt, f"found no application {fmt.mac_app}")] * n
    try:
        script = build_mac_script(fmt)
    except ValueError as exc:                 # ⑩ 表自己写漏了（工具的问题）
        return [f"工具配置有问题：{exc}"] * n
    argv = ["osascript", "-e", script]
    for src, dst in jobs:
        argv += [str(src), str(dst)]
    # 整批共用一条超时：每份最多 CONVERT_TIMEOUT，再按份数给点富余（首次拉应用偏慢）。
    limit = max(30, timeout) + 15 * n
    try:
        r = subprocess.run(argv, capture_output=True, text=True, timeout=limit)
    except subprocess.TimeoutExpired:
        return [(f"{fmt.mac_app} 超过 {limit} 秒没把这一批转完 —— "
                 "多半是弹了对话框在等人点（只读打开 / 有修订记录 / 首次运行的提示）。\n"
                 "去 Office 窗口那边把它处理掉，再重试。")] * n
    except OSError as exc:
        return [f"叫不动 {fmt.mac_app}：{exc}"] * n

    raw = ((r.stderr or "") + (r.stdout or "")).strip()
    errs = _parse_mac_batch_output(r.stdout or "", n)
    if all(e is None for e in errs):
        # 一行结果都没回来（编译不过 / 授权被拒 / 应用直接退出）→ 整批报同一句
        return [office_error_note("mac", fmt,
                                  raw or "没有报错信息，osascript 直接退出了")] * n
    for i, (_src, dst) in enumerate(jobs):
        e = errs[i]
        if e is None or e:                    # 没给结果 / 给了个错 → 都被归一地翻译
            errs[i] = office_error_note("mac", fmt,
                                        e or raw or "脚本没给出这份的结果")
        elif not dst.is_file():
            errs[i] = office_error_note("mac", fmt, "没有报错信息，但也没生成 PDF")
    return ["" if e is None else e for e in errs]


def _run_office_mac(src: Path, dst: Path, fmt: ConvFormat,
                    timeout: int = CONVERT_TIMEOUT) -> str:
    """单份版 —— 就是批量口的退化情形（命令行自检 / 回归里按一份调时用它）。"""
    return _run_office_mac_many([(src, dst)], fmt, timeout)[0]


def _probe_office_win() -> bool:
    if os.name != "nt":
        return False
    try:                                   # 只查有没有装 pywin32，不去启动 Office
        import importlib.util

        return importlib.util.find_spec("win32com") is not None
    except Exception:  # noqa: BLE001
        return False


def _run_office_win_many(jobs: list[tuple[Path, Path]], fmt: ConvFormat) -> list[str]:
    """Windows 的批量口：**一批同格式的文件，只启一个 Office 实例**。

    ⭐ 跟 macOS 那条是**同一个三步曲**（OFFICE_STEPS：只读打开 → 导出 → 丢弃关闭），
       连失败文案都走同一个 office_error_note()。差异只剩最底层的方言：
       这边是 COM 调用，那边是 AppleScript 语句。用哪个具体函数由 ⑩ 表的
       `win_kind` 决定（见 WIN_OFFICE_RUNNERS），所以这里不再有 `if key == "excel"`。

    ⭐ 整批复用一个 `DispatchEx` 实例（而不是每份启一次、转完就 Quit）：
       反复启停 Office 既慢，又最容易在「上一个实例还没退干净」时撞出占用 / 弹窗。
       Windows 上没有 macOS 那套「自动化授权」，所以这里不存在「每份都要点一下」的问题 ——
       但**每份重开一次 Word 一样会让它偶尔卡在启动画面上**，一并治掉。

    ⚠️ 必须在**这个线程**里 CoInitialize —— 转换是扔到后台线程跑的，
       Windows 上不初始化 COM 会直接报「尚未调用 CoInitialize」。
    """
    n = len(jobs)
    if n == 0:
        return []
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        return [f"缺 pywin32（在环境里执行 pip install pywin32）—— {exc}"] * n
    runner = WIN_OFFICE_RUNNERS.get(fmt.win_kind)
    if runner is None:                        # ⑩ 表自己写漏了（工具的问题）
        return [f"工具配置有问题：{fmt.key} 的 win_kind={fmt.win_kind!r} 没有对应实现"] * n
    pythoncom.CoInitialize()
    app = None
    errs: list[str] = []
    try:
        app = win32.DispatchEx(fmt.win_app)
        _init_office_app(app)                 # 见 APP_INIT_PROPS：认哪个设哪个
        for src, dst in jobs:
            # ⭐ 先记一笔产物现在什么样：一会儿要看「这次的 PDF 是不是这次写的」
            #    （覆盖那一次原件旁本来就有同名文件，见 _win_salvaged）
            before = _product_stamp(dst)
            try:
                runner(app, src, dst)
                errs.append("" if dst.is_file()
                            else office_error_note("win", fmt, "没有输出，但也没报错"))
            except Exception as exc:  # noqa: BLE001
                # ⭐ 有异常 ≠ 没转成：Office 常把「收尾」的错误抛在文件写完之后
                #    （见文件上方那段「假失败」）。产物确实新生成了 → 这份算成功。
                if _win_salvaged(dst, before):
                    _note_salvaged(src, exc)
                    errs.append("")
                else:
                    # 一份坏了只记这一份，接着转后面的（整批不必陪着一起失败）
                    errs.append(office_error_note("win", fmt, str(exc)))
    except Exception as exc:  # noqa: BLE001
        # 连 Office 都没起来（没装 / 被策略挡住）→ 整批报同一句
        errs = [office_error_note("win", fmt, str(exc))] * n
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
    if len(errs) < n:                         # 防御：任何意外都按 jobs 对齐
        errs += ["转换结果没拿到"] * (n - len(errs))
    return errs[:n]


def _run_office_win(src: Path, dst: Path, fmt: ConvFormat) -> str:
    """单份版 —— 就是批量口的退化情形（命令行自检 / 回归里按一份调时用它）。"""
    return _run_office_win_many([(src, dst)], fmt)[0]


def _probe_libreoffice() -> bool:
    return find_soffice() is not None


def _run_libreoffice(src: Path, dst: Path, fmt: ConvFormat,
                     timeout: int = CONVERT_TIMEOUT) -> str:
    """LibreOffice headless：转完落在临时目录，再搬到目标位置。

    ⭐ 这是**唯一一个 Mac / Windows 一字不差**的引擎（命令行、参数、产物处理全同）——
       所以它同时也是跨平台行为一致的那条退路。

    ⚠️ 一律用**独立的用户配置目录**（-env:UserInstallation）：否则会跟用户自己
       开着的 LibreOffice 抢同一个 profile，headless 会莫名卡住。
    ⚠️ --outdir 用临时目录：产物名固定是「原件名.pdf」，先转出去再改名搬到目标，
       免得同名文件在原件目录里被 LibreOffice 自己覆盖/加编号。
    """
    exe = find_soffice()
    if not exe:
        # 探针说能用、真跑却找不到 —— 属于本机状态变了。文案保持中性，不劝人装东西。
        return "本机这个转换程序找不到了（刚被卸载 / 挪走？）—— 点一次「重新检测引擎」重探。"
    outdir = Path(tempfile.mkdtemp(prefix="pdfseq_conv_"))
    profile = Path(tempfile.gettempdir()) / "pdfseq_lo_profile"
    cmd = [exe, "--headless", "--norestore", "--nolockcheck", "--nodefault",
           f"-env:UserInstallation={profile.as_uri()}",
           "--convert-to", f"pdf:{fmt.soffice_filter}",
           "--outdir", str(outdir), str(src)]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        shutil.rmtree(outdir, ignore_errors=True)
        return f"LibreOffice 超过 {timeout} 秒没转完"
    except OSError as exc:
        shutil.rmtree(outdir, ignore_errors=True)
        return f"叫不动 LibreOffice：{exc}"
    made = outdir / f"{src.stem}.pdf"
    if r.returncode == 0 and made.is_file():
        try:
            os.replace(made, dst)
        except OSError as exc:
            shutil.rmtree(outdir, ignore_errors=True)
            return f"转出来了但搬不过去：{exc}"
        shutil.rmtree(outdir, ignore_errors=True)
        return ""
    err = ((r.stderr or "") + (r.stdout or "")).strip()
    shutil.rmtree(outdir, ignore_errors=True)
    return f"LibreOffice 转换失败：{err[:300] or '没有输出'}"


class ConvEngine(NamedTuple):
    """⑪ 一个转换引擎。

    probe   ：现在这台机器上能用吗（要便宜，界面每次刷新都可能问）
    run     ：转**一份**。返回 "" = 成功；非空 = **给人看的**失败原因
    run_many：转**一批同格式的**。返回与这批等长的原因列表（"" = 这份成了）。

    ⭐ 为什么要有 run_many（2026-10-09 加）：把一批交给**同一个应用会话**，
       而不是「一份文件一次拉起应用」。这直接解决用户实测的两个毛病 ——
         · macOS 上「每一份都要点一下授权」：授权是按应用记的，一份一段脚本
           等于每份都重新问一次；
         · 「3 份只成功 1 份 / 试三次才全成」：每份都要重新拉起 Word，
           撞上「应用刚启动、文档还没交出来」的竞态。
    ⚠️ None ＝ 这个引擎没有批量实现，流程按单份循环调 `run`。
       回归脚本注入的假引擎就是这种（5 个字段），照样跑得通。
    """

    key: str
    label: str
    detail: str
    probe: Callable[[], bool]
    run: Callable[[Path, Path, ConvFormat], str]
    run_many: Callable[[list[tuple[Path, Path]], ConvFormat], list[str]] | None = None
    # ── ⭐ 2026-10-09 新增（用户问「引擎梯队」时收的口）────────────────────
    # family  ：伺服哪一族格式（⑩ 表的 family 字段）。**兜底只在同族内换** ——
    #            Excel 失手绝不会把活交给 AutoCAD，CAD 失手也不会掉到 LibreOffice。
    # platform：这条引擎在哪个平台上才成立。"mac" / "win" / "any"。
    #           ⚠️ 平台不符的行**根本不进梯队**（不是「进去了再靠 probe 返回 False」）——
    #              以前全靠探测函数里藏一句 `if sys.platform != "darwin": return False`，
    #              读表的人会以为 Mac 版里装着 Windows 引擎。现在明写在这一列上。
    # ⚠️ 两个都有默认值，是为了让回归里手搓的临时引擎（只给前 5 个字段）照样构造得出来；
    #    ⑪ 正式那几条**必须显式写**（回归 [13b] 盯着）。
    family: str = "office"
    platform: str = "any"


# ══════════════════════════════════════════════════════════════════════════
# ⑩d CAD 品牌表 + 调用层 —— 敲哪家的门、怎么让它出图
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ CAD「转 PDF」的真相：它**不是格式转换，是打印** ——
#    每一家 CAD 都只认同一句话：「把这张图，用某台打印机，打到这个文件」。
#    所以这一节要解决的只有两件事：**敲哪家的门**、**用哪台打印机**。
#
# ⭐⭐ 认程序靠 **COM ProgID，绝不猜安装路径**（2026-10-10 在用户的 Windows 上收口）。
#    下面这些家**共用同一套 ActiveX 对象模型**（`Documents.Open` / `ActiveLayout` /
#    `Plot.PlotToFile` …），所以「挑哪家转」的全部差异 = 这张表里的一行。
#
#    ⚠️ 教训（第一版就是这么错的）：最早写的是**按路径找 exe**
#       （`C:\Program Files\Autodesk\AutoCAD *\accoreconsole.exe`）。
#       现场实测：那台机器上根本没有 Autodesk 目录 —— 装的是**浩辰 CAD 2025**
#       （`C:\Program Files\浩辰软件\浩辰CAD2025`，中文目录名、没有公司英文名），
#       按路径 glob **必然漏**。用户一句话点破了要害：
#          「像 Word / Excel 你根本不需要知道它装在哪个位置，CAD 为什么需要？
#            而且每台公司的电脑都不同。」
#       ⇒ COM 是**安装时自己写进注册表**的：`reg query HKCR /f "GStarCAD" /k`
#          一查就摆着 `HKEY_CLASSES_ROOT\GStarCAD.Application`。
#          **注册表已经替我们记了它装在哪**，Dispatch 自己去问 —— 一行路径都不用写。
#       ⇒ 以后「本机有没有这个程序」一律先问注册表 / COM，别再猜安装目录。
#
# ⚠️ 表里的**顺序就是优先级**（同时装了两家就用靠前那个）。
#    2026-10-10 用户拍板收口：**只留浩辰 CAD 一家**。
#    原先表里还有 AutoCAD，另有三条注释掉的「预留位」（中望 / BricsCAD / nanoCAD）——
#    用户说公司这台实际只有浩辰，其余「不用管」，留着就是没人维护的死表项，已全删。
#    ⭐ **加一家 ＝ 照下面那行再加一行 `CadApp(...)`**：`progids` 填它的 COM 名字
#       （开 CMD 敲 `reg query HKCR /f "<它的英文名>" /k` 一查就有），
#       其余代码一个字都不用动 —— 品牌在这儿、调用逻辑不认品牌。
class CadApp(NamedTuple):
    key: str                      # 内部代号：gstarcad
    name: str                     # 给用户看的品牌名
    progids: tuple[str, ...]      # COM 名字，**按优先级**逐个试，登记着的那个赢


CAD_APPS: tuple[CadApp, ...] = (
    # ⚠️ progids 给多个是**刻意的**：有的机器只注册了带版本号那个
    #    （`GStarCAD.Application.25`），有的只注册不带版本的 —— 逐个问，命中即用。
    #    用户 2026-10-10 在真机上 `reg query HKCR /f "GStarCAD" /k` 实测过这三条。
    CadApp("gstarcad", "浩辰 CAD",
           ("GStarCAD.Application", "GStarCAD.Application.25", "GstarCAD.Application")),
)


def cad_brand_words() -> str:
    """⑩d 表里认的全部品牌名串成一句话（文案用 —— 别在别处再抄一遍品牌名）。"""
    return " / ".join(a.name for a in CAD_APPS)


# ── ⑩d2 用哪台「打印机」出 PDF：不写死设备名，问这台机器的 CAD 自己 ───────────
# ⭐ 同一个 bug 现场的第二处教训：原来把出图设备写死成 `DWG To PDF.pc3`。
#    各家 CAD 自带的 PDF 打印驱动名字都不一样（AutoCAD 叫 DWG To PDF.pc3，
#    浩辰未必叫这个），机器上还可能装着一堆第三方 PDF 打印机。**硬编码 = 换台机器就错**。
#    ⇒ 问布局自己：`GetPlotDeviceNames()` 把本机认得的设备全报上来，按关键词挑。
#    关键词**按想要的优先级排**：先找 CAD 自带的出图驱动，再退到任何带 PDF 字样的。
CAD_PDF_DEVICE_WORDS: tuple[str, ...] = ("dwg to pdf", "gstarcad", "cadpdf", "pdf")

# 出图前把这几个变量按掉（等价以前 .scr 里的 FILEDIA / CMDDIA），别让它弹窗等人点。
# ⚠️ BACKGROUNDPLOT 0 是**必须的**：后台打印开着时 PlotToFile 会**立刻返回**、
#    图纸还在后台慢慢写 —— 只判「文件在不在」就会把**写了一半的 PDF** 当成成功。
CAD_QUIET_VARS: tuple[tuple[str, int], ...] = (
    ("FILEDIA", 0), ("CMDDIA", 0), ("BACKGROUNDPLOT", 0))

# BarTender 的界面程序（安装目录名字各家不同，所以按「名字里带 BarTender」找）。
BARTENDER_ROOTS: tuple[str, ...] = (
    r"C:\Program Files",
    r"C:\Program Files (x86)",
)
BARTENDER_EXE = "bartend.exe"

# ⭐ BarTender 的 COM 名字（跟 CAD 的 ⑩d 表同一套思路：认名字、不猜安装路径）。
#    Automation 组件装着就会登记 `BarTender.Application`；有的版本只登记带版本号的。
#    逐个试，命中的那个赢。开 CMD 自查：`reg query HKCR /f "BarTender" /k`。
BARTENDER_COM_PROGIDS: tuple[str, ...] = (
    "BarTender.Application",
    "BarTender.Application.2021",
    "BarTender.Application.2022",
    "BarTender.Application.2019",
)


_cad_box: dict[str, object] = {}          # 进程级缓存，见 find_cad()


def cad_progid_registered(progid: str) -> bool:
    """注册表里登记着这个 COM 名字吗（**只读注册表、不起进程**，毫秒级）。

    ⭐ 探测**故意不 Dispatch**：Dispatch 会真把 CAD 拉起来（冷启动一二十秒），
       而 probe 每次刷界面都要问一遍。注册表里写着就＝装了；至于它启动时会不会
       卡在许可证向导上，那是**真转的时候**的事 —— 到那时报错会带着它自己的原话。
    """
    if os.name != "nt":
        return False
    try:
        import winreg                                         # noqa: PLC0415
    except Exception:                                         # noqa: BLE001
        return False                                          # 连 winreg 都没有（非 Windows）
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except OSError:
        return False


def _scan_cad() -> tuple[CadApp, str] | None:
    """按 ⑩d 表逐个问「登记了没」，第一个点头的就是它。纯查询，不启 CAD。"""
    for app in CAD_APPS:
        for pid in app.progids:
            if cad_progid_registered(pid):
                return app, pid
    return None


def find_cad(refresh: bool = False) -> tuple[CadApp, str] | None:
    """本机要用哪家 CAD —— 返回 (品牌, 它实际认的那个 ProgID)；一台没装则 None。

    ⭐ 为什么要缓存：注册表虽然便宜，但 GUI 每次刷新都要问一遍，
       而「装了还是没装」在一个进程的生命里通常不会变。
       代价是「工具开着的时候才装 CAD」要点一次「重新检测引擎」（refresh=True），
       跟 Office 那条是同一个规矩。
    """
    if refresh or "hit" not in _cad_box:
        _cad_box["hit"] = _scan_cad()
    hit = _cad_box["hit"]
    return hit if isinstance(hit, tuple) else None


def reset_probe_caches() -> None:
    """「重新检测引擎」时调 —— **所有**「注册表里有没有这个 COM 名字」的缓存一起清。

    ⚠️ 不能只清 CAD 那一份：Visio（见 find_visio）也是同一套缓存。
       漏一个，表现就是「刚装完点重新检测，还是说没找到」。
    """
    _cad_box.pop("hit", None)
    _visio_box.pop("hit", None)


# 旧名保留：seq_app 的「重新检测引擎」与回归里都按这个名字调（行为已含 Visio）。
reset_cad_cache = reset_probe_caches


def find_bartender() -> str | None:
    """bartend.exe 在哪（安装目录名字各家不同，所以按名字里带 BarTender 找）。"""
    if os.name != "nt":
        return None
    for root in BARTENDER_ROOTS:
        base = Path(root)
        if not base.is_dir():
            continue
        for sub in sorted(base.glob("*")):
            if not sub.is_dir() or "bartender" not in sub.name.lower():
                continue
            for cand in sorted(sub.glob("*.exe")):
                if cand.name.lower() == BARTENDER_EXE:
                    return str(cand)
    return shutil.which(BARTENDER_EXE)


def find_bartender_com() -> str | None:
    """BarTender 的 COM 名字登记着哪个（**只读注册表、不起进程**，毫秒级）。

    ⭐ 对齐 CAD 的 `cad_progid_registered`：认「名字」不猜安装路径 ——
       Automation 组件装着就会在 HKCR 登记 `BarTender.Application`。
       返回命中的 ProgID；一个都没登记（＝ Automation 组件没勾）则 None。
    """
    if os.name != "nt":
        return None
    for pid in BARTENDER_COM_PROGIDS:
        if cad_progid_registered(pid):
            return pid
    return None


def looks_like_pdf(path: Path) -> bool:
    """产物到底是不是**真 PDF**（看开头那 5 个字节）。

    ⭐ 为什么非要有这道核对：BarTender 与 CAD 都不是「格式转换」，是**打印** ——
       · 模板里存的打印机不是 PDF 打印机时，BarTender 会把**标签机自己的指令流**
         写进我们给的那个 .pdf 路径：文件在、名字对、就是打不开；
       · CAD 那边若那台「打印机」不是 PDF 打印机，同样会写出个打不开的东西。
       只判「文件在不在」会把这类**假成功**放过去，用户拿到一堆打不开的「PDF」。
    """
    try:
        with open(path, "rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def pick_pdf_device(names) -> str | None:
    """从一串打印设备名里挑出能出 PDF 的那台（**纯函数** —— 回归在 Mac 上就能盯住）。

    ⭐ 按 ⑩d2 的关键词逐个试、前面的优先 —— 先用各家 CAD 自带的出图驱动，
       实在没有再退到任何名字里带 PDF 的那台。
    """
    for word in CAD_PDF_DEVICE_WORDS:
        for n in names or ():
            if word in str(n).lower():
                return str(n)
    return None


def _probe_cad() -> bool:
    return find_cad() is not None


def _probe_bartender() -> bool:
    # ⭐ 出图走 COM（Formats.Open + Export…），所以**先认 COM 名字**；
    #    只有 exe、没 Automation 组件（没登记 COM）时，命令行也出不了 JPG，判「没有」。
    return find_bartender_com() is not None


def _open_cad_session(progid: str):
    """连上一个 CAD 会话。

    ⚠️ **绝不在退出时 Quit**：Dispatch 会复用用户**正在跑的那个实例**，
       Quit 掉等于把人家正在画的图一起关了。这个进程留给用户自己关。
    """
    import win32com.client as win32                            # noqa: PLC0415
    return win32.Dispatch(progid)


def _cad_layouts(doc) -> list[tuple[int, str, int]]:
    """这份图纸有哪些**图纸空间布局**（不含模型空间），按界面标签顺序返回。

    返回 [(界面顺序, 布局名, 在 Layouts 集合里的索引), ...]，界面顺序已升序。

    ⭐ 为什么跳过模型空间（2026-10-10 用户实测拍板）：
       一份 dwg 里「模型空间 + 布局1 + 布局2」，用户要的交付物是**布局页**不是模型空间；
       浩辰「批处理打印」默认也是「不勾模型」，只出布局们 —— 这里跟它一个口径。
    ⚠️ 三个「别猜」：
        · `Layouts.Count` 从 1 还是 0 起、`Item(i)` 吃不吃下标，各家不一样 →
          这里用 `Item(i)` 逐个试，Name 为空就跳过，不因一家下标习惯就整体失败；
        · 集合枚举顺序**不等于**界面标签顺序 → 用 `TabOrder` 拿到界面顺序再排序；
        · 布局名可能是「布局1/布局2」也可能是英文「Layout1/Layout2」→ 排序只看
          TabOrder，不 parse 名字里的数字。
    """
    out: list[tuple[int, str, int]] = []
    try:
        layouts = doc.Layouts
        n = int(getattr(layouts, "Count", 0))
    except Exception:                                    # noqa: BLE001
        return []
    for i in range(n):
        try:
            lay = layouts.Item(i)
        except Exception:                                # noqa: BLE001
            continue
        name = str(getattr(lay, "Name", ""))
        if not name or name.lower() == "model":
            continue                                     # 模型空间不要
        order = i
        try:
            order = int(getattr(lay, "TabOrder", i))
        except Exception:                                # noqa: BLE001
            pass
        out.append((order, name, i))
    out.sort(key=lambda t: t[0])
    return out


def _cad_content_size(lay) -> tuple[float, float] | None:
    """读这个布局里**图纸空间内容**的包围盒尺寸（宽, 高）。读不到返回 None。

    ⭐ 为什么用 `lay.Block.GetBoundingBox`：布局对象的 `.Block` 就是「图纸空间」那个块，
       里面装的就是这张图在布局里实际画的几何（含视口、图框、标注）。
       它的包围盒长宽比，比「模型空间 extmin/extmax」更能反映「这张图该横还是竖」。
    ⚠️ 各家的包围盒都返回 (min, max) 两个三维点（WCS 坐标），min 是左下、max 是右上。
    """
    try:
        block = lay.Block
        mn, mx = block.GetBoundingBox()          # 各是一个 3 元素数组 [x, y, z]
        w = float(mx[0]) - float(mn[0])
        h = float(mx[1]) - float(mn[1])
    except Exception:                            # noqa: BLE001
        return None
    if w <= 0 or h <= 0:
        return None
    return w, h


def _cad_fix_orientation(lay) -> str:
    """读这份布局「纸张横竖 vs 当前方向横竖」是否一致，**只读不写**，返回状态码。

    ⭐ 2026-10-10 真机踩坑后彻底改口径：**不再主动掰方向**。之前这函数会写
       `PlotRotation` 想把「拧了」的图掰正，结果反而把对的掰错，还踩了两个雷：
        ① 浩辰 `Layout.PlotRotation` 是 **AcPlotRotation 枚举**（AutoCAD 口径），不是角度：
             0 = ac0degrees  = 0°   = 横
             1 = ac90degrees  = 90°  = 竖
             2 = ac180degrees = 180° = 横
             3 = ac270degrees = 270° = 竖
           上一版按「0/180 竖、90/270 横」判，语义完全反了。
        ② 浩辰本来就会按图纸里自己存的 PlotRotation 正确出横竖，根本不用我们掰。
       ⇒ 结论：出图时**不碰方向**；这里只负责「读出来给人看」，诊断才调用。

    返回状态码（诊断口据此给文案）：
        · "ok"      纸张横竖与当前方向一致（这是绝大多数正常图）
        · "mismatch" 纸张与方向不一致（可能正是「转出来横竖不对」的病根，供人工判断）
        · "paper"   GetPaperSize 读不到，没法比
        · "rot"     PlotRotation 读不到，没法比
    """
    try:
        pw, ph = lay.GetPaperSize()              # (宽, 高)，纸张本体的名义尺寸
    except Exception:                            # noqa: BLE001
        return "paper"
    try:
        rot = float(lay.PlotRotation)
    except Exception:                            # noqa: BLE001
        return "rot"
    if pw <= 0 or ph <= 0:
        return "paper"
    paper_landscape = pw >= ph                   # 纸张本体横不横（宽 ≥ 高）
    # AcPlotRotation：浩辰返回**枚举** 0/1/2/3（=0°/90°/180°/270°），AutoCAD 返回角度。
    # 归一成角度再判：枚举 <4 就 ×90；角度本就 ≥4。角度 %180==0（0°/180°）→ 横，否则竖。
    rot_val = float(rot)
    angle = rot_val * 90.0 if abs(rot_val) < 4 else rot_val
    cur_landscape = abs(angle) % 180.0 < 1e-6    # 0°/180° 横；90°/270° 竖
    if paper_landscape == cur_landscape:
        return "ok"
    return "mismatch"


def _cad_match_media(lay, w: float, h: float) -> tuple[str, bool] | None:
    """在这台打印机的纸张清单里，找一张**跟 (宽,高) 同尺寸**的纸，设上并返回。

    返回 `(纸张规范名, 是否反向)`：反向 = 只有「高×宽」的纸没有「宽×高」的，
    设上它之后内容还得靠 PlotRotation 转回来。一张都没匹配上返回 None。

    ⭐ 为什么需要它（2026-10-10 真机定案）：`ConfigName = "DWG To PDF.pc3"` 会把布局
       的纸张**塌成 PDF 打印机的默认 A4**（横 A3 图 → 竖 A4、竖 A4 图 → 横 A4，
       PDF MediaBox 实测铁证）。方向 PlotRotation 根本没变，变的是纸张尺寸。
    ⭐ 做法：`GetCanonicalMediaNames()` 列出这台打印机认的所有纸张，逐个设上、
       `GetPaperSize()` 量尺寸，谁跟图纸原纸张 (w,h) 同尺寸就用谁 ——
       **先试正着匹配（宽=宽 高=高），再试反着匹配**。
       直接照抄图纸原来的 CanonicalMediaName 往往不行：那是**原来那台绘图仪**的纸名，
       PDF 打印机的清单里多半没有这个名字。
    """
    if w <= 0 or h <= 0:
        return None
    try:
        names = tuple(lay.GetCanonicalMediaNames())
    except Exception:                                    # noqa: BLE001
        return None
    for want_swapped in (False, True):                   # 先正匹配，再反匹配
        for cn in names:
            try:
                lay.CanonicalMediaName = str(cn)
                lay.RefreshPlotDeviceInfo()
                cw, ch = (float(x) for x in lay.GetPaperSize())
            except Exception:                            # noqa: BLE001
                continue
            tw, th = (h, w) if want_swapped else (w, h)
            if abs(cw - tw) < 1.0 and abs(ch - th) < 1.0:
                return str(cn), want_swapped
    return None


# ══════════════════════════════════════════════════════════════════════════
# -PUBLISH 批处理出图通道（2026-10-10 第二十三波）
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ 为什么另开这条路：`lay.ConfigName = dev`（COM 换 PDF 打印机）会把布局的
#    纸张**塌成该打印机默认 A4**（真机 PDF MediaBox 铁证：横 A3 图→竖 A4、
#    竖 A4 图→横 A4），方向 PlotRotation 反而没变。连续多轮「换前读回、换后设回」
#    都治不干净 —— 因为浩辰对 `CanonicalMediaName` / `PlotRotation` 的写回不老实。
# ⭐ 正路：跟用户在 CAD 里手动点「批处理打印（发布）」完全同一条路 —— 写一份
#    DSD（Drawing Set Description，图纸列表描述）文件，让 CAD 自己的 `-PUBLISH`
#    去读它出图。DSD 里每条只写「哪份图 + 哪个布局」，**Setup 留空 = 用布局自己
#    存好的页面设置**（横 A3 / 竖 A4 都照图纸原装），工具一个打印参数都不猜。
#    ⭐ AutoDesk 官方：`-PUBLISH` 在脚本/前台跑时 BACKGROUNDPLOT 被忽略、同步执行，
#       不会像 PlotToFile 那样提前返回。
# ⚠️ 浩辰的 -PUBLISH 具体交互/DSD 支持度还没真机验过 —— 所以这条是**新主路 +
#    旧 COM 路作降级**：-PUBLISH 走不通就退回原来的 ConfigName 那套，绝不卡死。


def _plot_one(app, names_fn, src: Path, dst: Path, brand: str) -> str:
    """这一份图纸 → PDF。返回 "" = 成了，否则是给人看的原因。

    ⭐ 出图**不替图纸做主**（用户 2026-10-10 定的口径：「导出什么就是什么，
       不需要人为微调」）：
        · 布局、纸张、比例、打印区域**全用图纸里自己存的** —— 跟 Word / Excel
          「读文件里存的打印设置」完全同一个口径，工具不微调、不猜方向；
        · 只额外指定两件事：**用哪台打印机（必须是 PDF 打印机）**、**打到哪个文件**。
    ⚠️ 「用哪台打印机」这一项**必须由我们指定**：图纸里存的多半是实体绘图仪，
       照它打就出不了 PDF。设备名问本机（`GetPlotDeviceNames`），别写死。

    ⭐ **自动判方向**（2026-10-10 用户报「新建 dwg 输出是竖版、上下留大空」）——
       唯一一处「掰图纸」的例外：读布局**内容包围盒**与**纸张方向**，拧了就把
       `PlotRotation` 翻 90°；**读不到内容就默认掰成横向**（DWG 大多横版，这是
       用户给的关键先验）。**只动方向**，纸张尺寸/比例/打印区域照旧。见
       `_cad_fix_orientation`。

    ⭐⭐ 为什么是「打全部布局、拼成多页 PDF」（2026-10-10 用户真机发现）：
       第一版只打 `doc.ActiveLayout` 一页 —— 而「当前布局」取决于图纸上次保存时停在哪个，
       结果是「有参照那张布局」没被打出来，用户以为参照没解析。
       真机对照：那份 dwg 打开时模型 + 布局1 + 布局2 **全部内容都完整可见**（参照根本
       没缺），浩辰「批处理打印」勾「布局1+布局2」两页都能正确出。
       ⇒ 根因是「只打一页」，不是「参照缺」。
       ⇒ 修法对齐浩辰批处理打印：枚举所有布局、跳过模型、按编号逐个 PlotToFile 成
          单页 PDF，再用 pypdf 按顺序拼成一个多页 PDF。
    ⚠️ 每份图纸**先打开一次读布局清单**，逐个切 `ActiveLayout = layout` 再出图 ——
       切完记得别改图纸里别的（只读打开 + 不保存，原件一字节不动）。
    """
    from pypdf import PdfReader, PdfWriter                    # noqa: PLC0415
    try:
        doc = app.Documents.Open(str(src), ReadOnly=True)
        layouts = _cad_layouts(doc)
        if not layouts:
            doc.Close(False)
            return (f"这份图纸里读不到任何「布局」（只有模型空间？）。\n"
                    f"工具默认只打布局页、跳过模型空间（跟浩辰「批处理打印」不勾模型一个口径）。")
        # 打印机问第一个布局（设备清单每个布局都一样，问一次就够）
        dev = pick_pdf_device(names_fn(doc.ActiveLayout))
        if dev is None:
            doc.Close(False)
            return ("这台机器上找不到能出 PDF 的打印设备。\n"
                    f"问过 {brand}，它认得的打印机里没有一个名字带 PDF。\n"
                    "先在 Windows 里装一台 PDF 打印机，或在 CAD 里确认「DWG To PDF」还在。")
        doc.SetVariable("BACKGROUNDPLOT", 0)      # ⭐ 关后台打印：不然 PlotToFile 会提前返回，
                                                  #    图纸还没画完就被当成「转好了」
        # 临时页放系统临时目录（不污染原件目录；跑完必删）
        tmp_dir = Path(tempfile.mkdtemp(prefix="pdfseq_cad_"))
        tmp_pages: list[Path] = []
        try:
            for _order, name, idx in layouts:
                lay = doc.Layouts.Item(idx)      # 用索引拿布局（比 Item(name) 更稳，各家都认）
                doc.ActiveLayout = lay           # 切到这个布局再打
                # ⭐⭐⭐ 第三十波（根因翻转）：之前连续七波都在「换设备 → 读回设回」上打转，
                #    越改白边越大 —— 因为一换 PDF 打印机，浩辰就会把布局的
                #    PlotOrigin/StandardScale/纸张 全套重置，读回设回永远治不干净
                #    （真机铁证：PlotOrigin (13.27,0)→(0,54.18)，内容顶左上、下方留白 54mm）。
                #    ⭐ 正路 = 跟「方向」那次成功同一招：**别动布局一个字**。
                #    Autodesk 官方 PlotToFile 的第二个参数可以传 PC3 设备名，
                #    **在不改布局当前设备的前提下**用指定设备出图（官方论坛 Solved 写法）。
                #    图纸存什么页面设置就原样出，白边该多少就是图纸本意多少 ——
                #    跟 CAD 里「批处理打印」逐字节同路。
                page_pdf = tmp_dir / f"{len(tmp_pages):03d}_{name}.pdf"
                ok = False
                try:
                    ok = bool(doc.Plot.PlotToFile(str(page_pdf), dev))
                except Exception:                              # noqa: BLE001
                    ok = False
                if not ok or not looks_like_pdf(page_pdf):
                    raise RuntimeError(f"布局「{name}」打出来不是 PDF")
                tmp_pages.append(page_pdf)
        finally:
            try:
                doc.Close(False)                 # False = 不保存，原件一个字节不动
            except Exception:                    # noqa: BLE001
                pass
        # 拼成一个多页 PDF
        try:
            if not tmp_pages:
                raise RuntimeError("一个布局都没打成 PDF")
            if len(tmp_pages) == 1:
                tmp_pages[0].replace(dst)
                return ""
            writer = PdfWriter()
            for p in tmp_pages:
                for pg in PdfReader(str(p)).pages:
                    writer.add_page(pg)
            with open(dst, "wb") as fh:
                writer.write(fh)
        finally:
            for p in tmp_pages:
                try:
                    p.unlink()
                except OSError:
                    pass
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except OSError:
                pass
    except Exception as exc:                      # noqa: BLE001
        return f"叫不动 {brand} 出图：{exc}"
    if not looks_like_pdf(dst):
        return (f"{brand} 说打完了，但落出来的**不是 PDF** ——\n"
                "多半是那台「打印机」其实不是 PDF 打印机（打出来的是它的打印指令）。\n"
                "换一台 PDF 打印机再试（关键词表见 ⑩d2）。")
    return ""


def _run_cad_many(jobs: list[tuple[Path, Path]], fmt: ConvFormat) -> list[str]:
    """Windows 批量口：**一个 CAD 会话串着打完整批**，不是一份图纸拉起一次 CAD。

    跟 ⑪ 表 `office_win` 那条完全同构：CoInitialize → Dispatch → 逐份 Open/Plot/Close。
    ⚠️ 必须在**当前线程** CoInitialize（转换是扔到后台线程跑的），
       不初始化会直接报「尚未调用 CoInitialize」。
    ⚠️ 这里只报**本族**的原因（找不到 CAD / 叫不响）；「这一族整个没引擎」由上层
       列队闸门说，别在这儿越权替别的族发言。
    """
    n = len(jobs)
    if n == 0:
        return []
    hit = find_cad()
    if hit is None:
        return [f"这台电脑上没找到 CAD（{cad_brand_words()}），转不了图纸。"] * n
    spec, progid = hit
    try:
        import pythoncom                                         # noqa: PLC0415
    except Exception as exc:                                     # noqa: BLE001
        return [f"缺 pywin32（环境里装一下 pip install pywin32）—— {exc}"] * n
    pythoncom.CoInitialize()
    try:
        app = _open_cad_session(progid)
        try:
            app.Visible = False
        except Exception:                                        # noqa: BLE001
            pass                                                 # 有的版本不吃这一句
        for key, val in CAD_QUIET_VARS:
            try:
                app.SetVariable(key, val)
            except Exception:                                    # noqa: BLE001
                pass                                             # 这个变量它不认就算了

        def _names(layout):
            try:
                # ⭐ 官方：GetPlotDeviceNames 前必须先 Refresh，否则换机可能漏设备
                try:
                    layout.RefreshPlotDeviceInfo()
                except Exception:                             # noqa: BLE001
                    pass
                return tuple(layout.GetPlotDeviceNames())
            except Exception:                                 # noqa: BLE001
                return ()

        return [_plot_one(app, _names, src, dst, spec.name) for src, dst in jobs]
    except Exception as exc:                                     # noqa: BLE001
        return [f"叫不动 {spec.name}：{exc}"] * n
    finally:
        pythoncom.CoUninitialize()


def _run_cad(src: Path, dst: Path, fmt: ConvFormat) -> str:
    """单份版 —— 批量口的退化情形（命令行 `--check-convert` 走它）。

    ⚠️ **签名必须正好是 `(src, dst, fmt)`**：转换层统一 `eng.run(src, dst, fmt)` 三参调用。
       上一版写成 `(src, dst)`，而注入假引擎的回归**根本走不到这条** → 真机上必 TypeError。
       回归 [13b3] 现在逐条盯着「引擎的 run / run_many 签名对不对」。
    """
    return _run_cad_many([(src, dst)], fmt)[0]


# ⭐⭐ BarTender 枚举的整数值（win32com 只能传整数，不能用枚举名）。
#    这些值是 Seagull 官方 COM 文档里的**稳定值**（btSaveOptions / btColors 多年没变过）；
#    万一某版本对不上，真机跑 `--show-bartender` 会把「导出到底报什么错」打出来。
#    出处见 Seagull BarTender Automation 文档（BtSaveOptions / BtColors 枚举）。
BT_SAVE_DONOTSAVE = 1     # BtSaveOptions.btDoNotSaveChanges
BT_SAVE_CHANGES = 2       # BtSaveOptions.btSaveChanges
BT_COLORS_24BIT = 4       # BtColors.btColors24bit（真彩色）

# 导出用的分辨率（JPG 和 PDF 都吃这个参数）：**1 = 打印机分辨率**，让 BarTender 按
#    模板里存的打印设置自适应，不自己拍脑袋定 DPI（-1=1dpi、0=屏幕、2~3000=具体值）。
#    （2026-10-12 真机：手动定 200/300 DPI 都算「自己添冗余规则」，改回 1 交给工具自适应。）
BT_JPEG_DPI = 1


def _bartender_open(app, src: Path):
    """打开一份 .btw —— `app.Formats.Open(path, False, "")`，只读打开、不弹子串名。

    ⚠️ 第二参 False = 不「打印前预览」；第三参 "" = 不预先填具名字串。
       返回 Format 对象；失败抛异常（调用方报错）。
    """
    return app.Formats.Open(str(src), False, "")


def _bartender_export_jpg(fmt_obj, dst_jpg: Path) -> None:
    """Format → 一张 JPG：走 `ExportToFile`（BarTender 的「文件 → 导出图片」）。

    ⭐ 用户点破（2026-10-12）：BarTender 出图是**导出**（Export），不是「打印」、
       也不是「另存为」，更不是「打印预览」。`ExportToFile` 导的是**模板本身**，
       尺寸＝标签实际大小（接近正方形、白边极小），**不套纸张/页面设置**。
       而 `ExportPrintPreviewToImage` 是「打印预览」路，会把标签套到 A4 空白页上
       （Seagull 官方论坛有用户 120×60 标签被放 A4 的实锤）—— 不能用那条。
       签名：ExportToFile(完整路径, 文件类型字符串, 颜色, DPI, 保存选项)。
    """
    fmt_obj.ExportToFile(
        str(dst_jpg), "jpg", BT_COLORS_24BIT, BT_JPEG_DPI, BT_SAVE_CHANGES)


def _bartender_export_pdf(fmt_obj, dst_pdf: Path) -> None:
    """Format → PDF：走 `ExportToFile`，FileType 传字符串 "PDF"。

    ⭐ 同样用「导出」而不是「打印预览」：PDF 尺寸也由标签本身决定，
       不套 A4 纸张（跟 JPG 同一套口径，两份各导一份、同尺寸）。
       ExportToFile 是 BarTender 6.0+ 就有的老接口，FileType 传字符串，
       兼容性最好、也避开 ImageType 枚举整数值随版本变的问题。
    """
    fmt_obj.ExportToFile(
        str(dst_pdf), "PDF", BT_COLORS_24BIT, BT_JPEG_DPI, BT_SAVE_CHANGES)


def _run_bartender(src: Path, dst: Path, fmt: ConvFormat) -> str:
    """一份 .btw → **两个产物**：`<同名>.jpg`（标签图片）+ `<同名>.pdf`（标签 PDF）。

    ⭐ 走 **COM 自动化**（跟 CAD 那条完全同构，2026-10-10 重写、2026-10-12 改方法）：
       `BarTender.Application` → `Formats.Open` → 分别 `ExportToFile`(jpg)
       + `ExportToFile`(pdf) → `Close` 不保存。**不再走命令行打印**。
       理由（用户点名）：「BarTender 是**导出**（Export），不是打印、不是另存为」——
       调它原生的「文件 → 导出图片/PDF」接口即可，所以：
        · 不依赖模板里存的打印机（不会再导出成标签机指令）；
        · 不自己写代码算尺寸/方向 —— 尺寸由标签本身决定（接近正方形、白边极小）。
    ⚠️ `dst` 是 PDF 落点（主产物，跟其它格式一致）；JPG 落点 = `dst` 换 `.jpg`（同名副产物）。
        JPG 由这里自己管，不污染上层「一份原件 → 一个 PDF」的通用链路（Word/Excel/CAD 不动）。
    ⚠️ **不 Quit**：跟 CAD 一样，用户可能正开着 BarTender；这个进程留给用户自己关。
    ⚠️ 签名同为三参（理由同 `_run_cad`）。
    """
    progid = find_bartender_com()
    if not progid:
        return ("没找到 BarTender 的自动化接口（COM 名字没登记）。\n"
                "多半是装 BarTender 时没勾「Automation」组件 —— 重装一次、勾上它。\n"
                "（只在 Windows 上能转 .btw，Mac 上没有 BarTender。）")
    try:
        import pythoncom                                     # noqa: PLC0415
        import win32com.client                               # noqa: PLC0415
    except Exception as exc:                                 # noqa: BLE001
        return f"缺 pywin32（环境里装一下 pip install pywin32）—— {exc}"
    dst_jpg = dst.with_suffix(".jpg")
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch(progid)
        try:
            app.Visible = False
        except Exception:                                    # noqa: BLE001
            pass                                             # 有的版本不吃这一句
        try:
            f = _bartender_open(app, src)
        except Exception as exc:                             # noqa: BLE001
            return f"BarTender 打不开这份模板：{exc}"
        try:
            try:
                _bartender_export_jpg(f, dst_jpg)
            except Exception as exc_j:                       # noqa: BLE001
                return f"BarTender 导出图片（JPG）失败：{exc_j}"
            try:
                _bartender_export_pdf(f, dst)
            except Exception as exc_p:                       # noqa: BLE001
                return f"BarTender 导出 PDF 失败：{exc_p}"
            # 两个都得真的落盘才认成功
            if not dst.is_file():
                return "BarTender 没生成 PDF（导出没报错，但文件没出现）。"
            if not looks_like_pdf(dst):
                return ("BarTender 写出来的**不是 PDF** —— 导出的可能不是标签 PDF。\n"
                        "真机跑一下 `--show-bartender` 看它认不认 PDF 导出。")
            if not dst_jpg.is_file():
                return "BarTender 生成了 PDF，但没生成 JPG（图片导出那步没落地）。"
            return ""
        finally:
            try:
                f.Close(BT_SAVE_DONOTSAVE)
            except Exception:                                # noqa: BLE001
                pass
    except Exception as exc:                                 # noqa: BLE001
        return f"叫不动 BarTender：{exc}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:                                    # noqa: BLE001
            pass
# ══════════════════════════════════════════════════════════════════════════
# ⑩e Visio 流程图 → PDF（2026-10-23 用户点名要的第五种格式）
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ 怎么转 —— 跟 Word / Excel **同一个三步曲**，只是换成 Visio 的成员名：
#      连上它（COM `Visio.Application`）→ 只读打开图纸 → 导出为 PDF → 不保存关掉。
#    整批复用一个会话（一个 `DispatchEx`），跟 office_win 完全同构；
#    **不经过打印**，所以不依赖本机装没装 PDF 打印机（这点比 CAD 那条省事）。
#
# ⭐ 为什么它自己一族（⑩a）而不是并进 office：见 ⑩ 表那一行上方的注释。

VISIO_COM_PROGIDS: tuple[str, ...] = (
    "Visio.Application",      # 任何版本都登记这一个（跟 Word.Application 同理）
    "Visio.Application.16",   # 2016 / 2019 / 2021 / 365
    "Visio.Application.15",   # 2013
    "Visio.Application.14",   # 2010
)

# 打开图纸用的标志位 —— Visio 没有 Office 那种 `DisplayAlerts`，
# **弹窗是靠打开时把这些标志位一起传进去关掉的**（VisOpenSaveArgs，按位和，
# win32com 只能传整数）：
#   visOpenRO(2)              只读打开 —— 原件一个字节不动
#   visOpenDontList(8)        不进「最近使用」（整批几十份不会把列表刷满）
#   visOpenMinimized(16)      不占窗口
#   visOpenMacrosDisabled(64) 不跑宏 —— .vsdm/.vstm 是带宏的，
#                             跑宏既会弹安全提示，又可能改原件
VISIO_OPEN_FLAGS = 2 | 8 | 16 | 64        # = 90

# ExportAsFixedFormat 的格式常量：VisFixedFormatTypes.visFixedFormatPDF = 1
VISIO_FMT_PDF = 1

# ⚠️⚠️ 参数个数是这个方法唯一的坑（2026-10-23 真机三份 .vsd 全挂的原因）：
#     官方签名是 `(FixedFormat, OutputFileName, Intent, PrintRange, FromPage, …)`，
#     其中 **`Intent` 和 `PrintRange` 是 Required**（文档原话），只有 `OutputFileName`
#     以及后面的 FromPage/ToPage/… 是 optional。
#     少给参数 → COM 直接回 `0x8002000E / -2147352562「无效的参数数目」`，
#     压根不进 Visio 的导出逻辑。所以**至少 4 个**，详见 `_visio_export_pdf`。
VISIO_INTENT_PRINT = 1     # visDocExIntentPrint　按打印意图导（工艺流程图是要打出来看的）
VISIO_INTENT_SCREEN = 0    # visDocExIntentScreen　只在线看（万一没装任何打印机时用）
VISIO_PRINT_ALL = 0        # visPrintAll　全部前景页（背景页由 IncludeBackground 管，默认带上）

# ⭐ 这次实际用的是哪一套签名 —— 只给 `--check-convert` 打一行，不上界面。
#    真机版本千奇百怪，留个现场，下次谁再说转不了，一眼能看出走的是哪条。
VISIO_EXPORT_USED: list[str] = []


_visio_box: dict[str, object] = {}        # 进程级缓存，见 find_visio()


def find_visio(refresh: bool = False) -> str | None:
    """本机 Visio 的 COM 名字登记着哪个（**只读注册表、不起进程**，毫秒级）。

    ⭐ 与 CAD 同一个理由（见 `cad_progid_registered`）：不能靠 Dispatch 探测 ——
       那会把 Visio 真拉起来（冷启动十几秒），而探测每次刷界面都会被问一遍。
       注册表里写着就＝装了；启动时会不会卡在激活向导上，那是**真转的时候**的事。
    ⭐ 缓存：跟 `find_cad` 一样是进程级的 —— 「工具开着的时候才装 Visio」
       要点一次「重新检测引擎」（`reset_probe_caches` 会连这个一起清）。
    """
    if refresh or "hit" not in _visio_box:
        _visio_box["hit"] = next((pid for pid in VISIO_COM_PROGIDS
                                  if cad_progid_registered(pid)), None)
    hit = _visio_box["hit"]
    return hit if isinstance(hit, str) else None


def _probe_visio() -> bool:
    return find_visio() is not None


def _init_visio_app(app) -> None:
    """Visio 起步时几个「别出声」的属性 —— **认哪个设哪个**（同 _init_office_app）。

    ⚠️ Visio 不认 `DisplayAlerts`（那是 Word / Excel 的），所以别往那张共用表里塞；
       它关弹窗靠的是打开标志位（见 VISIO_OPEN_FLAGS）。
    """
    for name, value in (("Visible", False), ("ScreenUpdating", False)):
        try:
            setattr(app, name, value)
        except Exception:                     # noqa: BLE001
            pass


def _visio_open(app, src: Path):
    """只读打开一份流程图 —— `Documents.OpenEx(路径, 标志位)`。

    ⚠️ `OpenEx` 是 Visio 专有的「带标志位打开」；万一某个版本不认（很老的版本 /
       第三方兼容包），退回普通 `Documents.Open` —— 后者会受宏提示与「最近使用」
       影响，但总比整份转不了强。
    """
    docs = app.Documents
    try:
        return _com_retry(docs.OpenEx, str(src), VISIO_OPEN_FLAGS)
    except Exception:                         # noqa: BLE001
        return _com_retry(docs.Open, str(src))


def _visio_quiet_close(doc) -> None:
    """关图纸：**关不上不算失败**（同 `_quiet_close` 的道理）。

    ⚠️ Visio 的 `Document.Close` 是**无参**的，跟 Office 的 `Close(False)` 不一样 ——
       拿 `_quiet_close` 去关它会 TypeError（异常被吞掉、文档其实没关、越积越多）。
       只读打开的图纸不会有「要不要保存」的提示，所以不需要传参。
    """
    if doc is None:
        return
    try:
        # 先标一句「没改动」—— 万一 OpenEx 退回了可写的 Open，
        # 导出后 Visio 可能认为文档「脏了」，Close 就会弹出「要不要保存」把人挂住。
        doc.Saved = True
    except Exception:                         # noqa: BLE001
        pass
    try:
        doc.Close()
    except Exception:                         # noqa: BLE001
        pass


def _visio_export_pdf(doc, dst: Path) -> None:
    """图纸 → PDF：`ExportAsFixedFormat(格式, 路径, 意图, 页范围)`。

    ⭐ 走 Visio 自己的「导出为 PDF」（等于界面上的 文件 → 导出 → 创建 PDF 文档），
       **不做打印**：不依赖本机装没装 PDF 打印机，也不会被打印驱动搅进边距/尺寸。
    ⭐ 纸张、缩放、页数**全用图纸里自己存的** —— 跟 Word / Excel「读文件里存的
       打印设置」同一个口径，工具一个参数都不替图纸做主。

    ⚠️⚠️ **坑就在参数个数**（2026-10-23 真机现场）：三份 `.vsd` 全报
       `(-2147352562, '无效的参数数目。')` ＝ `DISP_E_BADPARAMCOUNT`。
       当时只试了「2 参」「3 参」两套 —— 而官方签名里 `Intent` 与 `PrintRange`
       是**必填**，至少得给 4 个。给少了 Visio 连门都不让进，跟路径 / 权限 / 图纸
       内容一点关系都没有。
    ⚠️ 所以这里按「最完整 → 最简」依次试，**哪一套先落出真 PDF 就用哪一套**：
       Visio 各版本（2007 / 2010 / 2013 / 365）对这个方法的可选参数容忍度不一样，
       写死一套总有一边踩空。用过的签名记进 `VISIO_EXPORT_USED`，自检会打出来。
    """
    before = _product_stamp(dst)
    tried: list[tuple[str, BaseException]] = []
    for name, args in (
        # ① 官方文档的标准写法（VBA 示例就是这 4 个）—— 正常情况第一次就过
        ("4 参·打印意图", (VISIO_FMT_PDF, str(dst),
                           VISIO_INTENT_PRINT, VISIO_PRINT_ALL)),
        # ② 同上但「屏幕意图」：某些机器一台打印机都没装时，打印意图会被拒
        ("4 参·屏幕意图", (VISIO_FMT_PDF, str(dst),
                           VISIO_INTENT_SCREEN, VISIO_PRINT_ALL)),
        # ③④ 老版本 / 兼容包的退化写法（真机上没见过，但留着不花钱）
        ("3 参", (VISIO_FMT_PDF, str(dst), VISIO_INTENT_PRINT)),
        ("2 参", (VISIO_FMT_PDF, str(dst))),
    ):
        try:
            _com_retry(doc.ExportAsFixedFormat, *args)
        except Exception as exc:              # noqa: BLE001
            tried.append((name, exc))
            continue
        now = _product_stamp(dst)
        # ⭐「没报错」不等于「导出了」：必须真多出一份像样的 PDF（还要跟动手前比，
        #    否则覆盖那一次原件旁本来就有同名文件，会把没写当成写好了）。
        if now is not None and now != before and looks_like_pdf(dst):
            VISIO_EXPORT_USED.append(name)
            del VISIO_EXPORT_USED[:-5]        # 只留最近 5 条
            return
        tried.append((name, RuntimeError(
            "Visio 没报错，但那个位置没出现 PDF（可能被安全策略拦住了）")))
    # 全部失败：分两种情形报，别把 COM 的原话直接甩给用户。
    # ① 四套都回「参数个数不对」→ 这台 Visio 谁的账都不买，得人工看一眼；
    # ② 有真错误（图纸坏了 / 没权限 / 被策略挡）→ 报那条，「参数个数不对」没有解释力。
    if tried and all(_exc_hresult(exc) == COM_BADPARAMCOUNT for _n, exc in tried):
        raise RuntimeError(
            "这台电脑的 Visio 不认任何一种「导出为 PDF」的参数写法"
            "（四种都被回「无效的参数数目」）。\n"
            "请先在 Visio 里手动「文件 → 导出 → 创建 PDF 文档」试这份图纸：\n"
            "  能导出 → 是版本差异，把这句提示发给我；\n"
            "  也导不出 → 图纸本身有问题，跟本工具无关。\n"
            f"（原始报错：{tried[0][1]}）")
    for _name, exc in tried:
        if _exc_hresult(exc) != COM_BADPARAMCOUNT:
            raise exc
    if tried:
        raise tried[0][1]
    raise RuntimeError("Visio 导出没成功，也没给出原因")


def _visio_error_note(raw: object) -> str:
    """Visio 那份的失败原因（给人看的那一句）。"""
    msg = str(raw).strip()[:300] or "（没有报错信息）"
    return "Microsoft Visio 转换失败：\n" + msg


def _run_visio_many(jobs: list[tuple[Path, Path]], fmt: ConvFormat) -> list[str]:
    """Windows 批量口：**一个 Visio 会话导完整批**，不是一份图纸拉起一次 Visio。

    跟 ⑪ 表 office_win / cad 两条完全同构：CoInitialize → DispatchEx → 逐份
    OpenEx/Export/Close。④「假失败」那三道防线照样沿用（Visio 收尾也可能被
    RPC_E_CALL_REJECTED 顶一下，而 PDF 其实早写完了）：
      · `_quiet_close` 那条道理 → `_visio_quiet_close`
      · 「忙」类错误重试 → `_com_retry`
      · 产物为准 → `_win_salvaged`（**必须**跟动手前的 (mtime, size) 比）
    """
    n = len(jobs)
    if n == 0:
        return []
    progid = find_visio()
    if progid is None:
        return ["这台电脑上没找到 Microsoft Visio（COM 名字没登记），转不了流程图。"] * n
    try:
        import pythoncom                                     # noqa: PLC0415
        import win32com.client as win32                      # noqa: PLC0415
    except Exception as exc:                                 # noqa: BLE001
        return [f"缺 pywin32（在环境里执行 pip install pywin32）—— {exc}"] * n
    pythoncom.CoInitialize()
    app = None
    errs: list[str] = []
    try:
        app = win32.DispatchEx(progid)
        _init_visio_app(app)
        for src, dst in jobs:
            # 先记一笔产物现在什么样（覆盖那一次原件旁本来就有同名 PDF）
            before = _product_stamp(dst)
            doc = None
            try:
                doc = _visio_open(app, src)
                _visio_export_pdf(doc, dst)
                errs.append("" if dst.is_file()
                            else _visio_error_note("没有输出，但也没报错"))
            except Exception as exc:                         # noqa: BLE001
                # 有异常 ≠ 没转成（见 ④「假失败」）：产物确实新生成了 → 这份算成功
                if _win_salvaged(dst, before):
                    _note_salvaged(src, exc)
                    errs.append("")
                else:
                    errs.append(_visio_error_note(exc))       # 一份坏了不拖累整批
            finally:
                _visio_quiet_close(doc)
    except Exception as exc:                                 # noqa: BLE001
        # 连 Visio 都没起来（没装 / 被策略挡住 / 激活向导没走完）
        errs = [_visio_error_note(exc)] * n
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:                                # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:                                    # noqa: BLE001
            pass
    if len(errs) < n:                                        # 防御：按 jobs 对齐
        errs += ["转换结果没拿到"] * (n - len(errs))
    return errs[:n]


def _run_visio(src: Path, dst: Path, fmt: ConvFormat) -> str:
    """单份版 —— 批量口的退化情形（命令行 `--check-convert` 走它）。

    ⚠️ 签名必须正好是 `(src, dst, fmt)`（转换层统一三参调用，回归 [13b3] 盯着）。
    """
    return _run_visio_many([(src, dst)], fmt)[0]


# ══════════════════════════════════════════════════════════════════════════
# ⑪ 转换引擎表（续）—— 谁真的去执行转换
# ══════════════════════════════════════════════════════════════════════════
CONVERT_ENGINES: tuple[ConvEngine, ...] = (
    ConvEngine(
        "office_mac", "Microsoft " + family_kind_words("office", "mac") + "（本机）",
        "用本机装的 " + family_kind_words("office", "mac") + " 自己导出 PDF —— 保真度最高，"
        "文件里存的打印设置（打印区域、缩放）原样生效。\n\n"
        "**一次只拉起一次 Office**：整批文件交给同一个 " + family_kind_words("office", "mac") +
        " 会话，\n所以不会一份文件弹一次授权、也不会中途反复开关应用。\n\n"
        "首次使用 macOS 会为每个应用各问一次「允许控制 Microsoft Word / Excel」，\n"
        "点允许即可（**按应用记，不是按文件记**）。\n"
        "⚠️ 请**自己双击启动器**跑：从别的程序里代开的窗口拿不到这个授权。",
        _probe_office_mac, _run_office_mac, _run_office_mac_many,
        family="office", platform="mac",
    ),
    ConvEngine(
        "office_win", "Microsoft " + family_kind_words("office", "win") + "（本机）",
        "用本机装的 " + family_kind_words("office", "win") + " 自己导出 PDF —— 保真度最高，"
        "文件里存的打印设置（打印区域、缩放）原样生效。\n\n"
        "跟 macOS 那条是同一套三步曲（只读打开 → 导出 → 丢弃关闭），"
        "只是底层接口不同（COM vs AppleScript）。\n\n"
        "**整批复用一个 Office 实例**，不反复启停；Windows 上也没有"
        "「自动化授权」这一关，不会一份一份弹窗问你。",
        _probe_office_win, _run_office_win, _run_office_win_many,
        family="office", platform="win",
    ),
    # ⚠️ 这条是**自用兜底**，不是让用户去装的东西（2026-10-09 用户拍板）：
    #    机器上本来就有它 → 自动用得上；没有 → 一声不吭，界面与文案都不提它。
    #    所以 detail 写成中性描述，别在这里写「先装一个」。
    ConvEngine(
        "libreoffice", "LibreOffice（后台转换）",
        "**Mac 和 Windows 走的是同一套逻辑**（同一条命令行、同一套产物处理），"
        "不经过 Office。只有本机装着它、且 Office 那条没转成时才会轮到。\n\n"
        "⚠️ 分页口径是它自己那套，未必跟 Excel 里存的打印设置一模一样 ——\n"
        "　 所以有 Office 的机器一律优先走它。",
        _probe_libreoffice, _run_libreoffice,
        family="office", platform="any",
    ),
    # ── 下面两条**只在 Windows 上成立**（platform="win"）────────────────────
    # ⚠️ Mac 版**根本不会列出它们**（active_engines 按平台过滤），
    #    也不会把 CAD 的活交给上面那三条（family 不同，兜底只在同族内换）。
    # ⭐ **改走 COM 自动化**（2026-10-10 重写）：之前是 accoreconsole.exe + 一张我猜的
    #    -PLOT 应答表 —— 用户现场发现机器上是**浩辰 CAD 2025**（不是 AutoCAD），那套直接作废。
    ConvEngine(
        "cad", "CAD（本机）",
        "用本机装的 CAD 出 PDF：**跟 Office 那条完全同构** ——\n"
        "连上它（COM）→ 打开图纸 → 出图到 PDF → 不保存关闭。\n\n"
        "⭐ **不猜它装在哪**：靠 COM 名字（`GStarCAD.Application` 这种）找 ——\n"
        "　 跟 Word / Excel 一模一样，每台电脑装在哪都无所谓。\n"
        "⭐ **不替图纸做主**：布局、纸张、比例**都用图纸里自己存的** ——\n"
        "　 跟你说的「导出什么就是什么、不用人微调」是同一个口径。\n"
        "⭐ **整批只连一次**：一个 CAD 会话串着打完，不是一份图纸拉起一次。\n\n"
        "⚠️ 认的是 " + cad_brand_words() + "（⑩d 品牌表；加一家只加一行）。\n"
        "⚠️ 只在 Windows 上能用；Mac 上会自动隐去这条、也不提它。",
        _probe_cad, _run_cad, _run_cad_many,
        family="cad", platform="win",
    ),
    ConvEngine(
        "bartender", "BarTender（本机）",
        "用本机装的 BarTender 导出**两个文件**：一张标签图片（JPG）+ 一份标签 PDF。\n"
        "走的是 BarTender 自己的自动化接口（跟 CAD 那条完全同构）：\n"
        "　 打开模板 → 导出图片 → 导出 PDF → 不保存关掉，模板里的尺寸/内容原样导出、\n"
        "　 工具不旋转、不调整（你导什么就是什么）。\n\n"
        "⚠️ 要靠本机装的 BarTender（装的时候要勾「Automation」组件），Mac 上没有它。\n"
        "⚠️ 只有 .btw 会导出 JPG + PDF 两个；其它格式照旧只出 PDF。",
        _probe_bartender, _run_bartender,
        family="btw", platform="win",
    ),
    # ── ⭐ Visio 流程图（2026-10-23 加；同样只在 Windows 上成立）─────────────
    ConvEngine(
        "visio", "Visio（本机）",
        "用本机装的 Microsoft Visio 导出 PDF：**跟 CAD / BarTender 那条完全同构** ——\n"
        "连上它（COM `Visio.Application`）→ 只读打开图纸 → 导出为 PDF → 不保存关掉。\n\n"
        "⭐ **不经过打印**：走 Visio 自己的「导出为 PDF」，所以不依赖本机装没装\n"
        "　 PDF 打印机，也不会被打印驱动改掉尺寸或边距。\n"
        "⭐ **不替图纸做主**：纸张、缩放、页数都用图纸里自己存的（跟 Word / Excel 一个口径）。\n"
        "⭐ **整批只连一次**：一个 Visio 会话串着导完，不是一份图纸拉起一次。\n"
        "⭐ 打开时就只读、且**禁用宏**（.vsdm / .vstm 带宏）—— 原件一个字节不动。\n\n"
        "⚠️ 需要这台 Windows 电脑装了 Microsoft Visio（它不在 Office 套件里，得单独装）。\n"
        "⚠️ 只在 Windows 上能用；Mac 上没有 Visio，这条会自动隐去。",
        _probe_visio, _run_visio, _run_visio_many,
        family="visio", platform="win",
    ),
)


# ══════════════════════════════════════════════════════════════════════════
# ⑩b 转前预处理 —— 原件动不得，但有时候得「换个方式喂」给它
# ══════════════════════════════════════════════════════════════════════════
#
# ⭐ 为什么要有这一层（2026-10-09 用户实测）：
#    macOS 会给「来自网络」的文件打一个 `com.apple.quarantine` 标记。Word / Excel
#    见到它就以**受保护的视图**打开 —— 每份都要人工点一次「启用编辑」才肯导出，
#    用户体感就是「一份文件点一次，几十份点几十次」。
#    当天的对照实测：
#      · 微信来的 6 份（`0082;…;WeChat;`）      → 一口气转完，一次都不用点；
#      · 简历那 3 份（`0086;…;Microsoft Word;`）→ 每份都要点一下。
#    差的**不是「授权」**（那个按应用记，见 ⑪ 表），就是这层**文件级**标记。
#
# ⭐ 处理办法：**转副本**。实测（macOS 27.2）`cp` 到别处**不带** quarantine，
#    而 `ditto` 会带 —— 所以把原件复制进临时目录再喂给 Office，受保护视图就不出现。
#    三条底线，缺一不可：
#      · **原件一个字节都不动**：不改内容，也**不改它的扩展属性**（比 `xattr -d`
#        干净，不污染用户文件；回归 [13j] 盯着这条）；
#      · **产物路径不变**：照旧落在**原件旁边**；副本只拿去读，跑完就删；
#      · **不许帮倒忙**：查不到标记 / 复制失败 / 任何异常 → 一律老实退回原件。
#        宁可让它去撞受保护的视图（报错至少看得懂），也不能因为这一层凭空失败。
#
# ⭐ 第二条规则「被占用」（2026-11-06 用户实测，6.0 / 6.1 都有这毛病）：
#    原件正在 Word / Excel / WPS 里开着（或微信里刚点开过没关）时，喂原件会让
#    Word 弹「文件被锁定」三选一（打开只读副本 / 创建本地副本 / 接收通知）——
#    点哪个都关不掉，Mac 上 osascript 卡在对话框上，整批转不动。
#    病根是 Office 打开文档时在旁边写的「主人文件」（`~$同名`，LibreOffice 是
#    `.~lock.同名#`）—— 喂**副本**就没有锁，弹窗根本不出现。同样只动副本。
#
# 加规则 = 加一行：`detect` 说「这份要不要预处理」，`stage` 说「怎么弄出一份能用的」。
QUARANTINE_XATTR = "com.apple.quarantine"


class PrepRule(NamedTuple):
    key: str
    reason: str                       # 一句话：为什么要动它（给人看）
    detect: Callable[[Path], bool]    # 这份原件命中吗
    stage: Callable[[Path, Path], Path | None]   # (原件, 临时目录) → 可用的路径


def has_quarantine(path: Path) -> bool:
    """这份文件带 macOS「来自网络」标记吗？

    ⚠️ Python 在 macOS 上**没有** `os.getxattr`（那是 Linux 才有的），所以走系统
       自带的 `/usr/bin/xattr`。只是读属性，不起 Office，几十份也就几十毫秒。
    ⚠️ 非 macOS 一律 False —— 别的系统没有这套东西，别白跑子进程。
    """
    if sys.platform != "darwin":
        return False
    try:
        r = subprocess.run(["/usr/bin/xattr", "-p", QUARANTINE_XATTR, str(path)],
                           capture_output=True, text=True, timeout=5)
    except (OSError, subprocess.TimeoutExpired):
        return False                  # 查不出来 → 当没有 → 退回原件（保守）
    return r.returncode == 0 and bool((r.stdout or "").strip())


def has_owner_lock(path: Path) -> bool:
    """这份文件是不是**正被别的程序开着**（旁边躺着它的「主人文件」）。

    ⭐ Office 打开一份文档时会在同目录写一个 `~$同名` 的小文件（owner file）；
       LibreOffice 写 `.~lock.同名#`。看见它们 ＝ 这份正被 Word / Excel / WPS 开着，
       直接喂原件就会弹「文件被锁定」三选一（2026-11-06 用户实测：点哪个都关不掉，
       Mac 上整批卡死）。喂副本则根本没有锁 —— 弹窗不出现。
    ⚠️ 只做**只读的目录扫描**，不碰原件；目录读不出来一律当「没锁」（保守退回原件）。
    """
    name = path.name
    try:
        return ((path.parent / f"~${name}").exists()
                or (path.parent / f".~lock.{name}#").exists())
    except OSError:
        return False


def stage_clean_copy(src: Path, tmpdir: Path) -> Path | None:
    """复制出一份**干净副本**（无 quarantine、也天然没有 `~$` 锁），返回它的路径。

    「来自网络」和「被占用」两条规则共用这一个处理函数 —— 副本对两个病都免疫。

    ⭐ macOS 上用 `/bin/cp` 而不是 `shutil.copy2` / `ditto`：**实测**过 —— `cp` 不搬运
       扩展属性，副本天然干净；`ditto` 会把标记一起带过去，那就白忙了。
    ⚠️ Windows 上没有 quarantine 这套东西，走 `shutil.copy2` 即可（`/bin/cp` 在那边
       根本不存在 —— 「被占用」恰是 Windows 上更常见的场景，别让副本在那边失效）。
    ⚠️ macOS 结尾再显式 `xattr -d` 一次，纯属双保险（防哪天 `cp` 行为变了）。
       动的永远是**副本**：原件那一行在任何地方都不会出现 `xattr -d`。
    """
    dst = _free_name(tmpdir, src.name)
    try:
        if sys.platform == "darwin":
            r = subprocess.run(["/bin/cp", str(src), str(dst)],
                               capture_output=True, text=True, timeout=120)
            if r.returncode != 0:
                return None
        else:
            shutil.copy2(src, dst)
    except (OSError, subprocess.TimeoutExpired):
        return None
    if not dst.is_file():
        return None
    try:
        subprocess.run(["/usr/bin/xattr", "-d", QUARANTINE_XATTR, str(dst)],
                       capture_output=True, text=True, timeout=5)
    except (OSError, subprocess.TimeoutExpired):
        pass
    return None if has_quarantine(dst) else dst


def _free_name(tmpdir: Path, name: str) -> Path:
    """在临时目录里找个不重名的落点（同批可能有两份同名原件、只是不在一个文件夹）。"""
    dst = tmpdir / name
    n = 2
    while dst.exists():
        dst = tmpdir / f"{n}_{name}"
        n += 1
    return dst


PREP_RULES: tuple[PrepRule, ...] = (
    PrepRule("quarantine",
             "带 macOS 的「来自网络」标记（Word 会以「受保护的视图」打开，"
             "每份都要人工点一次「启用编辑」）",
             has_quarantine, stage_clean_copy),
    PrepRule("locked",
             "正被别的程序开着（旁边有它的 ~$ 锁文件，直接转会弹「文件被锁定」）",
             has_owner_lock, stage_clean_copy),
)


def first_prep_rule(path: Path) -> PrepRule | None:
    """这份原件命中哪条预处理规则？都没有 = None（绝大多数文件都是这个）。"""
    for rule in PREP_RULES:
        try:
            if rule.detect(path):
                return rule
        except Exception:             # noqa: BLE001  规则自己炸了不算命中
            continue
    return None


def safe_stage(rule: PrepRule, src: Path, tmpdir: Path) -> Path | None:
    """执行预处理。**任何异常都当失败**（返回 None），由调用方退回原件。"""
    try:
        out = rule.stage(src, tmpdir)
    except Exception:                 # noqa: BLE001
        return None
    return out if out is not None and out.is_file() else None


def prep_source(src: Path, tmpdir: Path) -> tuple[Path, str]:
    """这份原件**真正该拿去喂给 Office** 的是谁？

    返回 (路径, 原因)；原因 "" 表示原样用原件。**只读引用，绝不改动原件。**
    """
    rule = first_prep_rule(src)
    if rule is None:
        return src, ""
    out = safe_stage(rule, src, tmpdir)
    return (out, rule.reason) if out is not None else (src, "")


def _prep_sources(todo) -> tuple[dict[int, Path], Path | None]:
    """给整批算出「要换的份」→ {位置: 副本路径}，外加临时目录。

    ⚠️ 临时目录**按需创建**：一份都不用换（绝大多数情况）就一个目录都不建；
       一个都没换成也把空目录收掉。调用方负责在 finally 里删除。
    """
    work: dict[int, Path] = {}
    tmpdir: Path | None = None
    for pos, src, _target, _fmt in todo:
        rule = first_prep_rule(src)
        if rule is None:
            continue
        if tmpdir is None:
            tmpdir = Path(tempfile.mkdtemp(prefix="pdfseq_src_"))
        staged = safe_stage(rule, src, tmpdir)
        if staged is not None:
            work[pos] = staged
    if not work and tmpdir is not None:
        shutil.rmtree(tmpdir, ignore_errors=True)
        tmpdir = None
    return work, tmpdir


def active_engines(platform: str | None = None) -> list[ConvEngine]:
    """⑪ 表里**当前平台**上有可能用得上的引擎。

    ⚠️ 平台不符的行**在这里就被滤掉**，不是「进了梯队再靠 probe 返回 False」——
       以前全是靠探测函数里藏一句 `if sys.platform != "darwin": return False`，
       读表的人会以为「Mac 版里装着 Windows 引擎」。现在一眼就能看出来。
    """
    plat = platform or running_platform()
    return [e for e in CONVERT_ENGINES if e.platform in (plat, "any")]


def family_engines(family: str, platform: str | None = None) -> list[ConvEngine]:
    """某一族在当前平台上的引擎，**保持 ⑪ 表的顺序** —— 那个顺序就是优先级。"""
    return [e for e in active_engines(platform) if e.family == family]


def families_of(fmts: Sequence[ConvFormat]) -> list[str]:
    """这一批格式用到哪几族（保序去重）—— 列队闸门与引擎提示都读它。"""
    out: list[str] = []
    for f in fmts:
        if f.family not in out:
            out.append(f.family)
    return out


def pick_engine(family: str = "office", platform: str | None = None) -> ConvEngine | None:
    """⑪ 按表序探测**这一族**的引擎，返回第一个可用的；都没有返回 None。"""
    for e in family_engines(family, platform):
        if probe_engine(e):
            return e
    return None


def engine_by_key(key: str) -> ConvEngine | None:
    for e in CONVERT_ENGINES:
        if e.key == key:
            return e
    return None


def probe_engine(e: ConvEngine) -> bool:
    """一条引擎现在能不能用。探测自己炸了算「不能用」，不许拖垮整张表。"""
    try:
        return bool(e.probe())
    except Exception:  # noqa: BLE001
        return False


# ⭐ 引擎兜底：自动选出来的引擎**转失败了**，就按 ⑪ 表的顺序换下一个再试一遍。
#
#   为什么值得做：Mac 上「Office 自动化」天生脆（授权 / 沙箱 / 弹窗 / 没就绪），
#   Windows 上 Office 也可能卡在宏提示上。**一个引擎失手不该让整批白跑** ——
#   只要机器上还装了 LibreOffice，就应该让它接着把活干完。
#
#   ⚠️ 只在**自动挑引擎**时启用。回归脚本注入的假引擎是「钦定」的，
#      不参与兜底 —— 否则一个故意造失败的用例会去调真实的 LibreOffice。
#   ⚠️ 兜底是「前一个彻底报错」才触发（不含「没报错但没产物」之外的成功路径）；
#      已经转出 PDF 的不会再被第二个引擎覆盖。
#   ⭐ 2026-10-09：兜底**只在同族内**换（见 engine_chain）—— 一份 .dwg 失手，
#      绝不掉到 LibreOffice / Word 上去试；那既转不了，还会吐出莫名其妙的报错。
ENGINE_AUTO_FALLBACK = True


def engine_chain(first: ConvEngine, allow_fallback: bool = True) -> list[ConvEngine]:
    """第一个是要用的引擎，后面是**同族**的兜底梯队（按 ⑪ 表序，跳过自己、跳过不能用的）。

    ⭐ 两条硬约束：**同族**（`e.family == first.family`）与**同平台**
       （`e.platform in (当前平台, "any")`）。少了任何一条，加 CAD 那天就会出事。
    """
    if not allow_fallback or not ENGINE_AUTO_FALLBACK:
        return [first]
    plat = running_platform()
    return [first] + [e for e in CONVERT_ENGINES
                      if e.key != first.key
                      and e.family == first.family
                      and e.platform in (plat, "any")
                      and probe_engine(e)]


def engine_no_engine_msg(family: str = "office") -> str:
    """这一族**一个能用的引擎都没有**时的提示。

    ⚠️ 只说「装这一族要的那个程序」，**不要在这里推荐装别的转换软件**。
       2026-10-09 用户拍板：使用者的电脑不让随便装东西，指路去装第三方
       只会让人犯迷糊。⑪ 表里那条备用引擎依旧生效（机器上本来就有的话
       会自动用上），但不做宣传。
    ⚠️ 文案要**按族分**：CAD / BarTender 在 macOS 上根本装不了，
       笼统地写「去装一个」等于把人支使去找一个不存在的东西。
    """
    tmpl = FAMILY_NO_ENGINE.get(family)
    if tmpl is None:                       # 表里漏了这一族（只是文案，不许抛异常）
        return f"没找到能转{family_label(family)}的程序。"
    return tmpl.format(kinds=family_kind_words(family), what=family_label(family),
                       maker=family_maker(family))


# ══════════════════════════════════════════════════════════════════════════
# ⑫ 转换闸门表 —— 与 ② 同一种闸门，只是判据不同
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class ConvCtx:
    """转换闸门要看的东西。"""

    sources: list[Path]
    files: list[Path]                    # 收上来的、认得出格式的那些
    # ⭐ 2026-10-09 起不再是「一个全局引擎」—— 一个队列里可以同时有 Word 与 CAD，
    #    它们本来就该走两条不同的路（⑩a 的「族」）。
    families: list[str] = field(default_factory=list)          # 队列里出现的族（保序）
    engines: dict[str, ConvEngine] = field(default_factory=dict)   # 族 → 选中的引擎
    missing: list[str] = field(default_factory=list)           # 一族都没有引擎的那些族


def _cg_no_src(ctx: ConvCtx) -> str:
    if not ctx.sources:
        return ("还没有添加任何原件。\n"
                "点「＋ 添加文件…」把要转的 " + convert_kind_words() +
                " 挑进来 —— 支持一次多选，从不同文件夹里分几次挑也行。")
    missing = [p for p in ctx.sources if not p.exists()]
    if missing:
        names = "、".join(str(p) for p in missing[:3])
        more = f" 等 {len(missing)} 个" if len(missing) > 3 else ""
        return (f"这些来源已经不存在了：{names}{more}\n"
                "（被挪走了？在清单里把它移除、再重新添加一下。）")
    return ""


def _cg_empty(ctx: ConvCtx) -> str:
    if not ctx.files:
        return ("这些来源里没找到能转的 " + convert_kind_words() + " 文件。\n\n"
                "认这些扩展名：" + "、".join(CONVERT_EXTS) + "\n"
                "（文件夹来源只收它**这一层**里的文件 —— 子文件夹里的要单独添加；"
                "加了 PDF 也会被跳过，PDF 不需要转。）")
    return ""


def _cg_engine(ctx: ConvCtx) -> str:
    """⚠️ 只有**一族都轮不到**时才整批阻断。

    为什么不是「有一族没有引擎就全拦下」：用户手上常有混合队列，夹了一份 .dwg
    就让他那几十份 Word 也转不了，那是工具在给用户添堵。这类**局部**缺口由
    队列本身去说（那些行在表格里直接标出原因，见 plan_convert），其余照转。
    """
    if not ctx.families or not ctx.missing:
        return ""
    if len(ctx.missing) < len(ctx.families):
        return ""                          # 还有族转得了 → 放行
    return "\n\n".join(engine_no_engine_msg(f) for f in ctx.missing)


CONVERT_GATES: tuple[Gate, ...] = (
    Gate("有没有来源", _cg_no_src),
    Gate("有没有能转的文件", _cg_empty),
    Gate("有没有转换引擎", _cg_engine),
)


# ══════════════════════════════════════════════════════════════════════════
# 转换的数据模型
# ══════════════════════════════════════════════════════════════════════════

class ConvStatus(str, Enum):
    PENDING = "待转换"
    OK = "转换成功"
    FAILED = "转换失败"
    SKIPPED = "已跳过"


@dataclass
class ConvItem:
    """一份待转文件。"""

    index: int
    src: Path
    fmt_key: str = ""
    fmt_label: str = ""
    status: str = ConvStatus.PENDING.value
    note: str = ""


@dataclass
class ConvertPlan:
    """一次「列队」的结果。blocked=True 时一个文件都不转。"""

    items: list[ConvItem] = field(default_factory=list)
    blocked: bool = False
    block_reason: str = ""
    # ⭐ 2026-10-09：不再是「一个引擎」，而是**一族一条**（队列里可以同时有 Word 与 CAD）。
    #    · engines         —— 这一批用得到的引擎，保序（一族的首选那个）。
    #    · missing_families—— 队列里有、但本机没有程序能转的那几族。
    #    · engine_key / engine_label —— 「第一族那个」，留给界面与命令行（多数时候只有一族）。
    engines: list[ConvEngine] = field(default_factory=list)
    missing_families: list[str] = field(default_factory=list)
    engine_key: str = ""
    engine_label: str = ""
    # 清单里那些**认不出格式**的文件（只统计不进队，界面可以提一句）
    unknown: list[Path] = field(default_factory=list)


@dataclass
class ConvertResult:
    src: Path
    out: Path | None = None
    ok: bool = False
    status: str = ConvStatus.PENDING.value
    note: str = ""
    pages: int = 0                       # 转出来的 PDF 有几页（0 = 没读到）
    # 失败归类（见 ⑪c）："" ＝ 这行没出过问题；否则是 "no_permission" / "no_doc" /
    # "no_app" / "busy" / "syntax" / "other"。
    # ⚠️ 两种含义：① 这行失败了，它就是失败原因；② 这行后来被兜底梯队救活了，
    #    它记的是**首选引擎当初失手的原因**（这样界面照样能给「打开设置」的补救入口，
    #    让用户有机会根治，而不是每次都靠兜底）。
    # 界面靠它决定给不给补救入口 —— 别去解析 note 的文本。
    err_code: str = ""
    # 实际把这份转出来的引擎（⑪ 表的 key）。"" ＝ 没转成。
    # ⭐ 它跟 plan.engine_key 可能不一样 —— 首选引擎失手、兜底梯队接手时就会不同，
    #    所以「谁干的活」要按份记，不能拿计划里那个当答案。
    engine_key: str = ""
    # 这一份是**拿副本转的**（见 ⑩b：原件带「来自网络」标记时会换副本）。
    # ⚠️ 只是记账用的：`src` / `out` 仍然是原件与原件旁边的产物，副本跑完就删。
    via_copy: bool = False


# ══════════════════════════════════════════════════════════════════════════
# 收集待转文件（与「收集原件」同一套口径，只是认的扩展名不同）
# ══════════════════════════════════════════════════════════════════════════

def collect_convertibles(sources: Sequence[str | Path],
                         order_rule: str = DEFAULT_ORDER
                         ) -> tuple[list[Path], list[Path]]:
    """把来源清单里的 Word / Excel 收拢成一队。

    返回 (认得出格式的, 认不出的)。

    每个来源按类型处理（与 PDF 侧一致）：
        · 文件夹 → 只收它**这一层**里的 Word / Excel（不往子目录钻）
        · 单个文件 → 直接用（认不出格式的进 `unknown`，只统计不进队）

    排队顺序沿用 ① 的表（默认全部按文件名）—— 转换出来的 PDF 顺序跟替换那步
    对得上，后面摊开去打印、去扫都省事。
    """
    hit: list[Path] = []
    unknown: list[Path] = []
    seen: set[Path] = set()
    for s in sources:
        entry = Path(s).expanduser()
        if entry.is_dir():
            cands = [c for c in sorted(entry.glob("*"))
                     if c.is_file() and c.suffix.lower() in CONVERT_EXTS]
        elif entry.is_file():
            cands = [entry]
        else:
            continue                      # 不存在 → 交给闸门去报
        for p in cands:
            try:
                pr = p.resolve()
            except OSError:
                continue
            if pr in seen:
                continue                  # 同一个文件被指到两次
            seen.add(pr)
            if format_of(p) is None:
                unknown.append(p)         # 单文件来源里挑了个 PDF 之类 → 只统计
                continue
            hit.append(p)
    return sort_originals(hit, order_rule), unknown


# ══════════════════════════════════════════════════════════════════════════
# ⑨ 的转换版：产物名、产物识别、落点
# ══════════════════════════════════════════════════════════════════════════

def out_pdf_name_for(src: Path, suffix: str) -> str:
    """转换产物的文件名 = 原件名（可加后缀）+ .pdf。

    后缀留空也安全：原件是 .xlsx / .docx，产物是 .pdf，扩展名不同名必不同。
    """
    return f"{src.stem}{clean_output_suffix(suffix)}.pdf"


def convert_sibling_of(p: Path, suffix: str) -> Path | None:
    """p 认不认得出是「同目录里某份 Word / Excel 的转换产物」？认出就返回那份原件。

    判据（与 ⑧ 同一个思路：宁可漏判，也不误伤）：
        ① p 是 .pdf；
        ② 去掉后缀（后缀为空就直接用文件名）得到「原件名」；
        ③ 同目录**确实存在**那个名字 + 某个 Word / Excel 扩展名的文件。

    ⭐ 为什么要认：转换产物也写在原件旁边，第二次跑时会撞上上一轮的产物名。
       认得出 → 默认就原地更新（重跑＝刷新），而不是攒出一堆 xxx_2.pdf。
    ⚠️ 后缀为空时**不剥 _2 尾巴**（剥了会去猜，猜错就可能把别人的文件当自家的）：
       后果只是「点过否留下的 _2」不会被原地更新，它们本来也没人再动。
    """
    if p.suffix.lower() != ".pdf":
        return None
    sfx = clean_output_suffix(suffix)
    stem = p.stem
    if sfx:
        if stem.endswith(sfx):
            stem = stem[: -len(sfx)]
        else:
            m = _NUMBERED_TAIL.match(stem)
            if not m or not m.group("base").endswith(sfx):
                return None
            stem = m.group("base")[: -len(sfx)]
        if not stem:
            return None
    for ext in CONVERT_EXTS:
        cand = p.with_name(stem + ext)
        if cand.is_file():
            return cand
    return None


def is_convert_product(p: Path, suffix: str) -> bool:
    return convert_sibling_of(p, suffix) is not None


def convert_base_dir(src: Path, options: Options) -> Path:
    """转换产物该落在哪个目录（⑥ 之外的唯一口径，两处都调它，别各写一遍）。

    默认 = 原件所在目录（5.0 及以前的行为，一个字节都不变）；
    合并打印那一轮把 `options.out_dir` 填成临时目录，中间 PDF 就不落进用户的文件夹。
    """
    raw = (options.out_dir or "").strip()
    if not raw:
        return src.parent
    return Path(raw).expanduser()


def resolve_convert_target(src: Path, options: Options, used: set[Path]) -> Path:
    """转换产物的落点 —— 复用 ⑨ 那张表，只换 base 与「认不认得出是自家产物」。"""
    base = convert_base_dir(src, options) / out_pdf_name_for(src, options.suffix)
    # ⚠️ 换到临时目录时，`base` 旁边根本不会有对应的 Word/Excel 原件，
    #    `is_convert_product` 自然为 False —— 「覆盖自家产物」那条分支不会误触发。
    return _resolve_target(base, options, used, is_convert_product(base, options.suffix))


def convert_overwrite_list(plan: ConvertPlan, options: Options) -> list[tuple[Path, Path]]:
    """哪些**现有的同名 PDF**会被盖掉 → [(已有文件, 对应原件)]。给覆盖确认弹窗用。

    ⚠️ 与替换那边同一个口径：不看 `options.overwrite`，只回答「选覆盖会动到谁」。
    """
    out: list[tuple[Path, Path]] = []
    seen: set[Path] = set()
    for it in plan.items:
        if it.status == ConvStatus.SKIPPED.value:
            continue
        t = convert_base_dir(it.src, options) / out_pdf_name_for(it.src, options.suffix)
        if t.exists() and t not in seen:
            seen.add(t)
            out.append((t, it.src))
    return out


# ══════════════════════════════════════════════════════════════════════════
# 转换：列队 → 执行
# ══════════════════════════════════════════════════════════════════════════

def plan_convert(options: Options) -> ConvertPlan:
    """把来源清单算成一份「待转队列」。**不读文件内容**，只认扩展名 ——
    读页数得等真转出来（转换前那还是个 Excel，页数无从谈起）。"""
    src_paths = [Path(s).expanduser() for s in options.sources if str(s).strip()]
    files, unknown = collect_convertibles(options.sources, options.order_rule)

    # ⭐ 一族一条引擎：队列里可以同时有 Word 与 CAD，它们本来就走两条不同的路。
    #   探不到的那几族**不在这里抛异常、也不一定阻断** —— 交给 ⑫ 的闸门与队列去说。
    families = families_of([f for f in (format_of(p) for p in files) if f is not None])
    engines: dict[str, ConvEngine] = {}
    missing: list[str] = []
    for fam in families:
        eng = pick_engine(fam)
        if eng is None:
            missing.append(fam)
        else:
            engines[fam] = eng
    picked = [engines[f] for f in families if f in engines]

    plan = ConvertPlan(
        engines=picked,
        missing_families=list(missing),
        engine_key=(picked[0].key if picked else ""),
        engine_label=(picked[0].label if picked else ""),
        unknown=list(unknown))
    ctx = ConvCtx(sources=src_paths, files=files, families=families,
                  engines=engines, missing=missing)
    for gate in CONVERT_GATES:
        why = gate.check(ctx)
        if why:
            plan.blocked = True
            plan.block_reason = why
            return plan

    for i, p in enumerate(files, start=1):
        f = format_of(p)
        # ⭐ 这一族本机没有程序能转 → 队列里**当场**写清原因（不是等点完才报）。
        #    其余族照常转（这是「夹了一份 .dwg，别把 Word 也拖下水」的另一半）。
        note = engine_no_engine_msg(f.family) if (f and f.family in missing) else ""
        plan.items.append(ConvItem(index=i, src=p, note=note,
                                   fmt_key=f.key if f else "",
                                   fmt_label=f.label if f else "?"))
    return plan


def _engine_run_group(eng: ConvEngine, jobs: list[tuple[Path, Path]],
                      fmt: ConvFormat) -> list[str]:
    """一个引擎把**一批同格式**的文件转掉。返回与 jobs 等长的失败原因（"" = 成了）。

    ⭐ 有批量口的引擎（本机 Office 那两条）走批量：整批在一个应用会话里过一遍。
    ⭐ 没有批量口的（LibreOffice 命令行、回归里注入的假引擎）就老老实实一份一份来 ——
       流程只有一个入口，两条路都从这里走，免得以后又分叉。
    """
    if not jobs:
        return []
    if eng.run_many is not None:
        out = list(eng.run_many(jobs, fmt))
        # 防御：批量实现少给/多给了结果，按 jobs 对齐（少了算这份没转成，多了丢掉）
        if len(out) < len(jobs):
            out += ["批量接口没给出这份的结果"] * (len(jobs) - len(out))
        return out[:len(jobs)]
    return [eng.run(src, dst, fmt) for src, dst in jobs]


def _chain_for(fmt: ConvFormat, injected: ConvEngine | None) -> list[ConvEngine]:
    """这一族格式该用哪条梯队（「族」见 ⑩a）。

    ⭐ 注入的（回归脚本用）= **钦定**，一条到底、不参与兜底（见 ENGINE_AUTO_FALLBACK）。
    ⭐ 没注入就按 ⑪ 表挑这一族的首选，再沿**同族**梯队往下排。
    ⭐ 一族一个都挑不出 → 空列表；调用方据此把这一组直接判失败（别拿别的族去顶）。
    """
    if injected is not None:
        return [injected]
    first = pick_engine(fmt.family)
    return engine_chain(first, allow_fallback=True) if first else []


def _run_groups(injected: ConvEngine | None, fmt_order: list[str],
                groups: dict[str, list[tuple[int, Path, Path, ConvFormat]]],
                slot: dict[int, ConvertResult], work: dict[int, Path]) -> None:
    """逐组转：一组 = **一族**格式的一次应用会话；组内沿**同族**兜底梯队走。

    ⭐ 「一组一次会话」是 2026-10-09 第二轮定下的（治 macOS 的「每份都点授权」）。
    ⭐ 「组内只在同族里换引擎」是 2026-10-09 第四轮加的（加 CAD 那天必须的：
       一份 .dwg 失手时，梯队里只该有 CAD 那条路，绝不能把活交给 LibreOffice）。

    `work` = {位置: 真正喂给引擎的路径} —— 被 ⑩b 换成副本的那几份在这里，其余取原件。
    ⚠️ `slot[pos].src` / `.out` 永远是**原件与原件旁边的产物**：副本只出现在 `jobs`
       里，绝不外泄成结果（否则界面会显示临时目录里的路径）。
    """
    for fkey in fmt_order:
        pending = groups[fkey]
        fmt = pending[0][3]
        chain = _chain_for(fmt, injected)
        if not chain:                     # 这一族本机一个程序都没有（如 Mac 上放 .dwg）
            why = engine_no_engine_msg(fmt.family)
            for pos, _src, _tgt, _f in pending:
                res = slot[pos]
                res.note = why
                res.err_code = classify_office_error(why)
                res.status = ConvStatus.FAILED.value
            continue
        first_note: dict[int, str] = {}
        first_code: dict[int, str] = {}
        for k, cand in enumerate(chain):
            jobs = [(work.get(pos, src), tgt) for pos, src, tgt, _f in pending]
            errs = _engine_run_group(cand, jobs, fmt)
            nxt: list[tuple[int, Path, Path, ConvFormat]] = []
            for (pos, src, target, f), err in zip(pending, errs):
                if not err and not target.is_file():
                    err = "引擎没报错，但目标位置没有出现 PDF"
                if err:
                    if k == 0:                     # 首选那次的原因才是用户最该看的
                        first_note[pos] = err
                        first_code[pos] = classify_office_error(err)
                    nxt.append((pos, src, target, f))
                    continue
                res = slot[pos]                # 这一份成了
                res.ok = True
                res.status = ConvStatus.OK.value
                res.engine_key = cand.key
                if k > 0:
                    # 换过引擎 → 明说一句，免得用户以为一直是首选那个在干活
                    why = (first_note.get(pos, "") or "").strip().splitlines()
                    why = why[0][:70] if why else "原因不明"
                    res.note = (f"{chain[0].label} 没转成（{why}），"
                                f"已自动改用 {cand.label} 转好")
                    # 兜底救了场，但首选当初为什么失手还得留着 ——
                    # 界面靠它照样能给「打开授权设置」这类补救入口（根治，而不是每次都兜底）。
                    res.err_code = first_code.get(pos, "")
            pending = nxt
            if not pending:
                break
        for pos, _src, _target, _f in pending:   # 整条梯队都没救回来
            res = slot[pos]
            res.note = first_note.get(pos, "") or "转换失败"
            res.err_code = first_code.get(pos, "") or classify_office_error(res.note)
            res.status = ConvStatus.FAILED.value


def run_convert(plan: ConvertPlan, options: Options,
                engine: ConvEngine | None = None
                ) -> tuple[list[ConvertResult], list[str]]:
    """按队列转换。返回 (结果, 提示语)；结果的顺序与 plan.items 一一对应。

    ⚠️ `engine` 可注入：不给就按 ⑪ 表自己探测。回归脚本靠这个注入假引擎，
       「整条流程」才能在没装 Office 的机器上照样验证。
    ⚠️ 原件只读：整段代码没有任何一处往原件写东西；转出来的 PDF 落在它旁边。

    ⭐⭐ 2026-10-09（第二轮）：**先按格式分组，一组交给同一个应用会话**。
       旧写法是「每份文件都调一次引擎」，于是每份都重新拉起一次 Word / Excel ——
       这正是 macOS 上「每份文件都要点一次授权」和「3 份只成功 1 份」的根源。
       分组之后：Word 组只拉一次 Word、Excel 组只拉一次 Excel，
       换格式才换应用；而**结果仍按 plan.items 的原顺序交回**，界面一行对一行。

    ⭐ 2026-10-09（第三轮）：前面再加一道 ⑩b **转前预处理** —— 带 macOS「来自网络」
       标记的、或正被 Word / Excel 开着（旁边有 `~$` 锁）的原件，换成**临时副本**
       再喂给 Office（原件一个字节不动），绕开受保护视图和「文件被锁定」三选一弹窗。
    """
    if plan.blocked:
        return [], [plan.block_reason or "被阻断，未执行"]
    # ⚠️ 引擎**按组现挑**（一族一条，见 _chain_for）—— 一个队列里可以同时有 Word 与 CAD。
    #    回归脚本注入的假引擎是「钦定」的，不参与兜底：见 _chain_for。

    used: set[Path] = set()
    slot: dict[int, ConvertResult] = {}       # 位置 → 结果（最后按原顺序交回）
    todo: list[tuple[int, Path, Path, ConvFormat]] = []

    # ① 先给每一行定好产物落点，把不用真转的（跳过 / 认不出 / 演练）就地结掉
    for pos, it in enumerate(plan.items):
        if it.status == ConvStatus.SKIPPED.value:
            slot[pos] = ConvertResult(src=it.src, out=None, ok=True,
                                      status=ConvStatus.SKIPPED.value,
                                      note="你手动忽略了这一行")
            continue
        target = resolve_convert_target(it.src, options, used)
        res = ConvertResult(src=it.src, out=target)
        slot[pos] = res
        fmt = format_by_key(it.fmt_key) or format_of(it.src)
        if fmt is None:
            res.note = "认不出这是什么格式"
            res.status = ConvStatus.FAILED.value
            continue
        if options.dry_run:
            res.ok = True
            res.status = ConvStatus.OK.value
            res.note = "演练：没有真的转"
            continue
        todo.append((pos, it.src, target, fmt))

    # ② 按格式分组（保序）—— 一组 = 同一个 Office 应用的一次会话
    fmt_order: list[str] = []
    groups: dict[str, list[tuple[int, Path, Path, ConvFormat]]] = {}
    for row in todo:
        key = row[3].key
        if key not in groups:
            groups[key] = []
            fmt_order.append(key)
        groups[key].append(row)

    # ③ 转前预处理（⑩b）：要「换个方式喂」的原件，先复制一份干净副本出来
    work, tmpdir = _prep_sources(todo)
    for pos in work:
        slot[pos].via_copy = True
    try:
        # ④ 逐组转；组内沿**同族**兜底梯队走（首选失手就换同族的下一个接着来）
        _run_groups(engine, fmt_order, groups, slot, work)
    finally:
        if tmpdir is not None:                     # 副本用完就删，不留痕
            shutil.rmtree(tmpdir, ignore_errors=True)

    # ⑤ 读页数（真转出来的那些）—— 页数有值就说明 PDF 真的生成好了
    #    顺带按 ⑩c 的规则核对「应该出几页」（Excel ＝ 工作表数）；**只在不符时**出声。
    for pos, it in enumerate(plan.items):
        res = slot.get(pos)
        if res is None or not res.ok or not res.engine_key \
                or res.out is None or res.status != ConvStatus.OK.value:
            continue
        n, _sz, rerr = read_pdf_info(res.out)
        res.pages = n
        if n == 0:
            extra = f"转出来了，但读不出页数（{rerr}）"
            res.note = f"{res.note}；{extra}" if res.note else extra
            continue
        fmt = format_by_key(it.fmt_key)
        warn = page_rule_report(fmt, res.src, n) if fmt else ""
        if warn:                       # 规则对得上就一个字都不说
            res.note = f"{res.note}；{warn}" if res.note else warn

    results = [slot[i] for i in range(len(plan.items)) if i in slot]
    notes: list[str] = []
    if work:
        # 换过副本就交代一句（命令行会打出来；界面侧走 via_copy 记账，不弹窗）。
        notes.append(f"有 {len(work)} 份改用临时副本转换（避开「来自网络」标记或文件被占用），"
                     f"原件一个字节都没动。")
    return results, notes


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="seq_core.py",
        description=f"{TOOL_NAME} v{TOOL_VERSION}（命令行）",
    )
    p.add_argument("--mode", choices=("swap", "convert"), default="convert",
                   help="convert（默认）= 把 "
                        + convert_kind_words() + " 批量转成 PDF；swap = PDF 首页替换")
    p.add_argument("--src", action="append", default=[], metavar="路径",
                   help="原件来源：一个文件夹（收它这一层里的 PDF）或一个 PDF 文件；"
                        "可重复给多个，例如 --src A文件夹 --src B/C.pdf")
    p.add_argument("--scan", default="",
                   help="一次扫完的多页扫描件 PDF。**不给就只列队**（打印排队结果，不执行）")
    p.add_argument("--reverse", action="store_true", help="扫描件页序倒序")
    p.add_argument("--no-scale", action="store_true", help="不缩放，扫描页原样作为第 1 页")
    p.add_argument("--overwrite", dest="overwrite", action="store_const", const=OVERWRITE_YES,
                   default=OVERWRITE_AUTO,
                   help="一律覆盖同名产物（不加则：认得出是上一轮产物的原地更新，"
                        "认不出的加编号 _2、_3 让开）")
    p.add_argument("--no-overwrite", dest="overwrite", action="store_const", const=OVERWRITE_NO,
                   help="一律不覆盖：撞名就加编号 _2、_3 让开")
    p.add_argument("--suffix", default="", help="输出文件名后缀（**必填**，产物写在原件旁边）")
    p.add_argument("--order", default=DEFAULT_ORDER,
                   choices=[r.key for r in ORDER_RULES],
                   help="原件排队规则（默认 name = 文件夹里按名称看到的顺序）")
    p.add_argument("--rotate", default="",
                   help="人工判的方向：「页码:角度,页码:角度」，例 3:90,5:270（正数向右转）")
    p.add_argument("--no-auto-rotate", action="store_true",
                   help="关掉自动方向（默认：原件首页是横向的，那一页自动左转 90°）")
    p.add_argument("--dry-run", action="store_true", help="只演练不写文件")
    p.add_argument("--check-convert", nargs="?", const="", default=None, metavar="文件",
                   help="诊断用：**真的试转一份**（不给文件名就用 --src 里的第一份），"
                        "把选中的引擎、归类结果、原始报错原样打出来。"
                        "产物写在临时目录、读完就删，原件一个字节都不动")
    p.add_argument("--show-cad", action="store_true",
                   help="诊断用：CAD 环境体检 —— 这台机器找不找得到 CAD、找的是谁、"
                        "它认不认得 PDF 打印机（只读注册表 / 只问 CAD 自己，不碰任何文件）。"
                        "CAD 转不出 PDF 时先看它。配 --cad-file 一份 .dwg 会额外诊断"
                        "「这份图为什么转出来是横/竖」：把纸张尺寸、内容包围盒、当前"
                        "方向原样打出来")
    p.add_argument("--cad-file", default="", metavar="路径",
                   help="配 --show-cad 用：给一份 .dwg，诊断它的布局方向判定"
                        "（为什么转出来是横版还是竖版）。只读打开、不保存、不碰原件")
    p.add_argument("--show-rules", action="store_true",
                   help="诊断用：把本机认哪些格式、每族有哪些引擎、在哪个平台，一次列清楚")
    p.add_argument("--show-bartender", action="store_true",
                   help="诊断用：BarTender 环境体检 —— 自动化接口（COM）有没有登记、"
                        "连不连得上、导出方法在不在（只读注册表 / 只问 BarTender 自己，"
                        "不碰任何文件）。配 --bartender-file 一份 .btw 会实际试导出一次")
    p.add_argument("--bartender-file", default="", metavar="路径",
                   help="配 --show-bartender 用：给一份 .btw，实际试导出 JPG+PDF"
                        "（产物写临时目录、读完就删，原件一个字节不动）")
    p.add_argument("--brief", action="store_true",
                   help="精简输出（配 --show-rules / --show-cad 用）：体检.bat 里用，"
                        "只报「认哪些格式、每族谁干、找到谁、挑哪台打印机」，不展开明细表")
    p.add_argument("--no-head", action="store_true", dest="no_head",
                   help="配 --show-* 用：不打自己那两行抬头 —— 给启动器"
                        "统一版式用，抬头由 setup_env.py 的 bootui 出")
    return p


def _main_show_rules(a) -> int:
    """把「这台机器上到底谁管谁」一次列清楚 —— 排查「怎么不认我的文件」时先看它。

    用户 2026-10-09 问「引擎梯队是什么意思」的时候，缺的就是这么一个入口。
    """
    if getattr(a, "brief", False):
        # 精简版：体检.bat 里用，只报结果、不铺明细表
        # ⭐ 行式统一走 bootui 的模板（`[ OK ] / [ -- ]` 那套），别在这里自造符号。
        _cad_hit = find_cad()
        if not getattr(a, "no_head", False):
            print(f"{TOOL_NAME} v{TOOL_VERSION} —— 认哪些、谁转")
            print(f"本机平台：{running_platform()}")
        # ⚠️ 原来这里还打一行 18 个扩展名，体检时纯属噪音（而且折三行）：
        #    「能转什么」看下面每一族的结论就够了；真要逐条扩展名，
        #    跑不带 --brief 的 `--show-rules`。
        for fam in ENGINE_FAMILIES:
            eng = pick_engine(fam)
            B.item(B.OK if eng else B.SKIP, B.pad(FAMILY_SHORT.get(fam, fam), 16),
                   eng.label if eng else "没找到能转它的程序")
        if _cad_hit:
            B.item(B.OK, B.pad("CAD 程序", 16), _cad_hit[0].name)
        return 0
    print(f"{TOOL_NAME} v{TOOL_VERSION} —— 格式与引擎清单")
    print(f"本机平台：{running_platform()}（sys.platform={sys.platform}）")
    print(f"\n【⑩ 认这些扩展名】{'、'.join(CONVERT_EXTS)}")
    print("\n【⑩ 格式表】")
    for f in CONVERT_FORMATS:
        here = "本机有" if running_platform() in f.platforms else "本机没有"
        print(f"  {f.key:<6} {f.label:<10} 族={f.family:<7} 平台={'/'.join(f.platforms):<7} "
              f"{here}  分页规则={f.page_rule}")
    print("\n【⑪ 引擎表】同族内按下面的顺序排队；兜底只在同族内换。")
    for e in CONVERT_ENGINES:
        plat_ok = e.platform in (running_platform(), "any")
        mark = "✔ 本机可用" if (plat_ok and probe_engine(e)) else (
            "— 本机没有这个平台" if not plat_ok else "✘ 装了才可用")
        print(f"  {e.key:<11} 族={e.family:<7} 平台={e.platform:<4} {mark:<16} {e.label}")
    print("\n【本机选中的引擎（一族的第一个可用引擎）】")
    for fam in ENGINE_FAMILIES:
        eng = pick_engine(fam)
        print(f"  {family_label(fam):<22} {eng.label if eng else '（没有可用引擎）'}")
    _cad_hit = find_cad()
    print("\n【⑩d CAD 品牌表】去注册表问「登记了没」（认 COM 名字、不看安装路径）；"
          f"本机认到的：{_cad_hit[0].name if _cad_hit else '（没有 —— 非 Windows 上属正常）'}")
    for app in CAD_APPS:
        print(f"  {app.name:<10} {app.key:<9} {'、'.join(app.progids)}")
    return 0


def _main_show_cad(a) -> int:
    """CAD 环境体检 —— 找不找得到 CAD、找的是谁、它认不认得 PDF 打印机。

    ⭐ 为什么要这么个入口（2026-10-10 重写）：CAD / BarTender 这半套**没法在
       开发这台 Mac 上验**（用户公司那台装的是浩辰 CAD 2025，Mac 上根本装不了）。
       与其让我在 Mac 上猜，不如给一个**只读**的口子，在真机上跑一眼就知道：
         · 找不找得到 → 查注册表（认的是 **COM 名字**，不是安装路径）；
         · 找的是谁   → ⑩d 品牌表里哪个名字登记着；
         · 出不出得来 → 让它把本机认得的打印机列出来，看挑中的那台是不是 PDF 打印机。
    ⚠️ 第三步会**真把 CAD 拉起来**问一句（要问它设备清单就得有个文档对象），
       会顺手开一张**空图纸**、问完就关、**不保存**；用户正在画的东西不受影响，
       问完也**不 Quit**（那是用户自己的实例）。
    ⚠️ 上一版这个口子叫 `--show-cad-script`，打印的是我**猜**的 17 行 -PLOT 应答表 ——
       用户一句「其实是浩辰 CAD 2025」它就全废了。现在**不再猜任何东西，只问机器**。
    """
    if not getattr(a, "no_head", False):
        print(f"{TOOL_NAME} v{TOOL_VERSION} —— CAD 环境体检（只读，不碰任何文件）")
        print(f"本机平台：{running_platform()}（sys.platform={sys.platform}）")
    if getattr(a, "brief", False):
        hit = find_cad(refresh=True)
        if hit is None:
            # ⚠️ Mac 上「没有 CAD」是**正常**的（--），别用「失败」吓人
            B.item(B.SKIP if os.name != "nt" else B.FAIL,
                   B.pad("CAD 程序", 16),
                   "没找到（CAD 只有 Windows 版，本机是 %s，正常）" % running_platform()
                   if os.name != "nt" else "注册表里没登记任何 CAD 的 COM 名字")
            return 0
        spec, progid = hit
        B.item(B.OK, B.pad("CAD 程序", 16), f"{spec.name}（COM 名 = {progid}）")
        try:
            app = _open_cad_session(progid)
        except Exception as exc:                              # noqa: BLE001
            B.item(B.FAIL, B.pad("CAD 程序", 16), f"名字登记着，但叫不应它：{exc}")
            return 1
        try:
            doc = app.Documents.Add()
            try:
                lay = doc.ActiveLayout
                # ⭐ 官方：GetPlotDeviceNames 前必须先 RefreshPlotDeviceInfo，
                #    否则可能只返回部分设备（换台机器就漏掉 DWG To PDF.pc3）。
                try:
                    lay.RefreshPlotDeviceInfo()
                except Exception:                             # noqa: BLE001
                    pass
                names = tuple(lay.GetPlotDeviceNames())
            finally:
                doc.Close(False)
        except Exception as exc:                              # noqa: BLE001
            B.item(B.FAIL, B.pad("CAD 程序", 16), f"连上了，但问不出打印机：{exc}")
            return 1
        picked = pick_pdf_device(names)
        if picked is None:
            B.item(B.FAIL, B.pad("PDF 打印机", 16), "一台都没找到（名字里没有 PDF）")
            B.note("它实际认得的打印设备（用来判断是「没装 PDF 打印机」还是「名字对不上」）：")
            for nm in names:
                B.note("  " + nm)
            # ⭐ HideSystemPrinters=1 会把系统打印机藏起来，导致列表不完整
            try:
                _hsp = app.GetVariable("HideSystemPrinters")
            except Exception:                                 # noqa: BLE001
                _hsp = "读不到"
            B.note(f"系统变量 HideSystemPrinters = {_hsp}（=1 表示隐藏了系统打印机）")
            return 1
        B.item(B.OK, B.pad("PDF 打印机", 16), picked)
        return 0
    print("\n【⑩d 品牌表】照这张表去问注册表（认 COM 名字，不猜安装路径）：")
    for i, app in enumerate(CAD_APPS, 1):
        print(f"  {i}. {app.name:<10} {app.key:<9} {'、'.join(app.progids)}")
    hit = find_cad(refresh=True)
    if hit is None:
        print("\n❌ 一个都没登记（注册表里找不到上面任何一个 COM 名字）。")
        if os.name != "nt":
            print(f"   —— 这是正常的：CAD 只有 Windows 版，本机是 {running_platform()}。")
        else:
            print("   确认它装好了没有。装了还是这样，开 CMD 自查一句：")
            print('     reg query HKCR /f "GStarCAD" /k')
        return 0
    spec, progid = hit
    print(f"\n✅ 找到：{spec.name}（COM 名 = {progid}）")
    print("   ⭐ 靠注册表里的**名字**找到的 —— 它装在哪、装了几个版本，都不用管。")
    print("\n（下一步会把 CAD 拉起来问一句「你认得哪些打印机」，冷启动可能要一二十秒；"
          "\n  会顺手开一张空图纸，问完就关、不保存。）")
    try:
        app = _open_cad_session(progid)
    except Exception as exc:                                     # noqa: BLE001
        print(f"\n⚠️ 名字登记着，但叫不应它：{exc}")
        print("   （很多 CAD 第一次被外部程序调用会弹许可证 / 向导窗口，"
              "把它点掉再试一次。）")
        return 1
    try:
        doc = app.Documents.Add()       # 空图纸只为问一句设备清单，不读用户任何文件
        try:
            names = tuple(doc.ActiveLayout.GetPlotDeviceNames())
        finally:
            doc.Close(False)
    except Exception as exc:                                     # noqa: BLE001
        print(f"\n⚠️ 连上了，但问不出打印机清单：{exc}")
        return 1
    print(f"\n它认得的打印设备（共 {len(names)} 台）：")
    for nm in names:
        print(f"  {nm}")
    picked = pick_pdf_device(names)
    if picked is None:
        print("\n❌ 没有一台名字带 PDF —— 这样出不了 PDF。")
        print("   给它装一台 PDF 打印机再来（CAD 自带的「DWG To PDF」也算）。")
        return 1
    print(f"\n✅ 会用它出 PDF：{picked}")
    print("\n⭐ 出图**不替图纸做主**：布局 / 纸张 / 比例 / 打印区域都用图纸里自己存的，"
          "只把出图设备换成上面这台。")

    # ── 方向诊断：给了 --cad-file 就打开它，把「判方向」用到的三个值原样打出来 ──
    if (getattr(a, "cad_file", "") or "").strip():
        _show_cad_orientation(app, Path(a.cad_file).expanduser())
    return 0


def _main_show_bartender(a) -> int:
    """BarTender 环境体检 —— 自动化接口（COM）在不在、连不连得上、导出方法认不认。

    ⭐ 为什么要有（2026-10-10 加）：BarTender 出图**改走 COM**（Formats.Open + Export…），
       而这套跟 CAD 一样**没法在开发这台 Mac 上验**。给一个只读口子，真机一跑就知道：
         · COM 名字登记着没（装的时候勾没勾「Automation」组件）；
         · Dispatch 连不连得上；
         · Formats.Open 打不打得开 .btw、Export 方法在不在、导出报什么错。
    ⚠️ 会真把 BarTender 拉起来（Dispatch 才问得到对象模型）；问完**不 Quit**
       （那是用户自己的实例）。
    """
    if not getattr(a, "no_head", False):
        print(f"{TOOL_NAME} v{TOOL_VERSION} —— BarTender 环境体检（只读，不碰任何文件）")
        print(f"本机平台：{running_platform()}（sys.platform={sys.platform}）")
    progid = find_bartender_com()
    if progid is None:
        if os.name != "nt":
            B.item(B.SKIP, B.pad("BarTender", 16),
                   "注册表里没登记（BarTender 只有 Windows 版，本机是 %s，正常）"
                   % running_platform())
        else:
            B.item(B.FAIL, B.pad("BarTender", 16), "注册表里没登记它的 COM 名字")
            B.note("排查：装 BarTender 时有没有勾「Automation」组件？")
            B.note('自查一句（CMD）：reg query HKCR /f "BarTender" /k')
        return 0
    B.item(B.OK, B.pad("BarTender", 16), f"COM 名 = {progid}")
    try:
        import pythoncom                                     # noqa: PLC0415
        import win32com.client                               # noqa: PLC0415
    except Exception as exc:                                 # noqa: BLE001
        B.item(B.FAIL, B.pad("BarTender", 16), f"缺 pywin32：{exc}")
        return 1
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch(progid)
        B.item(B.OK, B.pad("BarTender", 16),
               f"连上了（版本 {getattr(app, 'Version', '读不到')}）")
        if getattr(a, "bartender_file", "").strip():
            src = Path(a.bartender_file).expanduser()
            if not src.is_file():
                B.item(B.FAIL, B.pad("试导出", 16), f"给的这份找不到：{src}")
                return 1
            B.item(B.OK, B.pad("试导出", 16), src.name)
            tmp = Path(tempfile.mkdtemp(prefix="bt_check_"))
            try:
                f = app.Formats.Open(str(src), False, "")
                try:
                    dst_jpg = tmp / f"{src.stem}.jpg"
                    dst_pdf = tmp / f"{src.stem}.pdf"
                    try:
                        _bartender_export_jpg(f, dst_jpg)
                        B.item(B.OK, B.pad("JPG", 16),
                               f"{dst_jpg.stat().st_size} 字节")
                    except Exception as exc_j:               # noqa: BLE001
                        B.item(B.FAIL, B.pad("JPG", 16), str(exc_j))
                    try:
                        _bartender_export_pdf(f, dst_pdf)
                        B.item(B.OK, B.pad("PDF", 16),
                               f"{dst_pdf.stat().st_size} 字节"
                               + ("" if looks_like_pdf(dst_pdf) else "（但不是 PDF 魔数）"))
                    except Exception as exc_p:               # noqa: BLE001
                        B.item(B.FAIL, B.pad("PDF", 16), str(exc_p))
                finally:
                    f.Close(BT_SAVE_DONOTSAVE)
            except Exception as exc:                         # noqa: BLE001
                B.item(B.FAIL, B.pad("试导出", 16), f"打不开 / 导不出：{exc}")
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
        return 0
    except Exception as exc:                                 # noqa: BLE001
        B.item(B.FAIL, B.pad("BarTender", 16), f"名字登记着，但叫不应它：{exc}")
        return 1
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:                                    # noqa: BLE001
            pass


def _show_cad_orientation(app, src: Path) -> None:
    """诊断「这份图为什么转出来是横/竖」—— 把判方向用到的三个值原样打出来。

    ⭐ 只读打开、读完就关、不保存，原件一个字节不动。
    ⭐ 打印三样东西（都按「自动判方向」里实际读的口径）：
        · GetPaperSize()  —— 纸张名义宽高（没考虑旋转）；
        · Block.GetBoundingBox() —— 图纸空间内容的包围盒（判「内容横竖」的依据）；
        · PlotRotation —— 当前方向（0/180 竖、90/270 横）。
      有了这三行，就能一眼看出「判方向没生效」到底卡在哪一环：
        GetBoundingBox 读到 0 / 异常 / 范围跟整张纸一样大，都是各自的病。
    """
    print(f"\n\n════ 方向诊断：{src.name} ════")
    if not src.is_file():
        print(f"❌ 找不到这个文件：{src}")
        return
    try:
        doc = app.Documents.Open(str(src), ReadOnly=True)
    except Exception as exc:                                  # noqa: BLE001
        print(f"❌ 打不开：{exc}")
        return
    try:
        layouts = _cad_layouts(doc)
        if not layouts:
            print("  这份图里读不到任何布局（只有模型空间？）。")
            return
        print(f"  布局清单（按界面顺序）："
              f"{'、'.join(f'{name}(#{idx})' for _o, name, idx in layouts)}")
        for _o, name, idx in layouts:
            try:
                lay = doc.Layouts.Item(idx)
            except Exception as exc:                          # noqa: BLE001
                print(f"\n  · 布局「{name}」：拿不到对象 —— {exc}")
                continue
            print(f"\n  · 布局「{name}」（索引 {idx}）：")
            # ① 纸张名义尺寸
            try:
                pw, ph = lay.GetPaperSize()
                print(f"      GetPaperSize() = ({pw:.2f}, {ph:.2f})  "
                      f"→ 纸张名义 {'横' if pw >= ph else '竖'}")
            except Exception as exc:                          # noqa: BLE001
                print(f"      GetPaperSize() 读不到 —— {exc}")
            # ② 内容包围盒（仅供参考 —— 浩辰常读不到，判方向不靠它了）
            csize = _cad_content_size(lay)
            if csize is None:
                print("      Block.GetBoundingBox() 读不到（返回 None / 异常 / 0）"
                      "  ← 浩辰常这样，判方向已改用 GetPaperSize")
            else:
                cw, ch = csize
                print(f"      Block.GetBoundingBox() 内容 = ({cw:.2f}, {ch:.2f})  "
                      f"→ 内容 {'横' if cw >= ch else '竖'}")
            # ③ 当前方向（AcPlotRotation：浩辰返回枚举 0/1/2/3 = 0°/90°/180°/270°）
            try:
                rot = float(lay.PlotRotation)
                _ang = rot * 90.0 if abs(rot) < 4 else rot
                rot_word = "横" if abs(_ang) % 180.0 < 1e-6 else "竖"
                print(f"      PlotRotation = {rot:g}  → 当前 {rot_word}版"
                      f"（浩辰枚举 0/180°=横、90/270°=竖）")
            except Exception as exc:                          # noqa: BLE001
                print(f"      PlotRotation 读不到 —— {exc}")
            # ④ 方向结论（只读不写：出图不碰方向，浩辰自己按存的设置出）
            try:
                status = _cad_fix_orientation(lay)
                if status == "ok":
                    print("      → 方向：纸张横竖与当前方向一致（没问题）")
                elif status == "mismatch":
                    print("      → 方向：纸张与当前方向**不一致**（可能正是转出来横竖不对的病根）")
                elif status == "paper":
                    print("      → 方向：GetPaperSize 读不到，没法比")
                else:  # "rot"
                    print("      → 方向：PlotRotation 读不到，没法比")
            except Exception as exc:                          # noqa: BLE001
                print(f"      → 方向判定报错 —— {exc}")
            # ⑤ 换设备实测（2026-10-10 真机定案）：`ConfigName = "DWG To PDF.pc3"` 会把
            #    纸张**塌成 PDF 打印机默认 A4**（MediaBox 铁证：横 A3 图→竖 A4），
            #    方向 PlotRotation 反而不变。这里模拟换设备，并实测两种「找回纸张」的办法：
            #      办法A：把图纸原来的纸张规范名（CanonicalMediaName）设回去；
            #      办法B：按尺寸在 PDF 打印机纸张清单里匹配（出图正式用的就是它）。
            try:
                dev_name = pick_pdf_device(tuple(lay.GetPlotDeviceNames()))
            except Exception:                                  # noqa: BLE001
                dev_name = None
            if dev_name:
                try:
                    b_pw, b_ph = (float(v) for v in lay.GetPaperSize())
                except Exception:                              # noqa: BLE001
                    b_pw = b_ph = 0.0
                try:
                    b_rot = float(lay.PlotRotation)
                except Exception:                              # noqa: BLE001
                    b_rot = None
                try:
                    b_canon = str(lay.CanonicalMediaName)
                except Exception:                              # noqa: BLE001
                    b_canon = None
                # ⭐⭐ 第二十五波：白边偏大根因——换设备把「比例/区域/居中」重置了。
                #    但只报「变了没变」不够：浩辰可能对这几个属性返回**一致但错误**的
                #    假值（比如 StandardScale 恒返回一个占位值），前后 repr 一样就误判
                #    成「没动」，白边病根就被藏起来了。第二十七波升级：换设备前后各把
                #    缩放相关属性**原始值**打一遍，没变也打出来，人眼直接看值对不对。
                _plot_props = ("StandardScale", "UseStandardScale",
                               "PlotType", "CenterPlot", "PlotOrigin")
                b_plot = {}
                for _k in _plot_props:
                    try:
                        b_plot[_k] = getattr(lay, _k)
                    except Exception as _e:                    # noqa: BLE001
                        b_plot[_k] = f"读不到({_e.__class__.__name__})"
                try:
                    lay.ConfigName = dev_name
                    lay.RefreshPlotDeviceInfo()
                except Exception as exc:                       # noqa: BLE001
                    print(f"      [换设备实测] 换 ConfigName 失败：{exc}")
                else:
                    # 换设备后，把缩放相关属性也读一遍，逐个打印**原始值**（含没变的）
                    _plot_changes = []
                    _a_plot = {}
                    for _k in _plot_props:
                        try:
                            _av = getattr(lay, _k)
                        except Exception as _e:                # noqa: BLE001
                            _av = f"读不到({_e.__class__.__name__})"
                        _a_plot[_k] = _av
                        if repr(_av) != repr(b_plot[_k]):
                            _plot_changes.append(_k)
                    print("      [缩放/区域原始值] 换设备前后（StandardScale 0=缩放到适合；"
                          "PlotType 0屏幕/1范围/2界限/3视图/4窗口/5布局；UseStandardScale 是否用标准比例）：")
                    for _k in _plot_props:
                        _mark = "  ← 变了" if _k in _plot_changes else ""
                        print(f"         · {_k}: {b_plot[_k]!r} → {_a_plot[_k]!r}{_mark}")
                    if _plot_changes:
                        print("      [比例/区域实测] 换设备改了上面标「←变了」的项。")
                    else:
                        print("      [比例/区域实测] 换设备没动这些值 —— 若 StandardScale 不是 0、"
                              "或 PlotOrigin 不是 0，白边病根就在这里的原始值，不在「换设备」这一层。")
                    try:
                        a_pw, a_ph = (float(v) for v in lay.GetPaperSize())
                        a_paper = f"({a_pw:.2f},{a_ph:.2f})"
                    except Exception:                          # noqa: BLE001
                        a_pw = a_ph = 0.0
                        a_paper = "读不到"
                    try:
                        a_rot = float(lay.PlotRotation)
                    except Exception:                          # noqa: BLE001
                        a_rot = None
                    paper_tag = ("、纸张没变" if b_pw > 0 and abs(a_pw - b_pw) < 1
                                 else f"、纸张 {b_pw:.0f}×{b_ph:.0f}→{a_pw:.0f}×{a_ph:.0f}（被换成默认）")
                    rot_tag = "" if b_rot is None or b_rot == a_rot else f"、方向 {b_rot:g}→{a_rot:g}"
                    print(f"      [换设备实测] 换成「{dev_name}」后{paper_tag}{rot_tag}")
                    # 办法A：照抄图纸原纸张规范名设回去
                    if b_canon is not None:
                        try:
                            lay.CanonicalMediaName = b_canon
                            lay.RefreshPlotDeviceInfo()
                            r_pw, r_ph = (float(v) for v in lay.GetPaperSize())
                            ok = abs(r_pw - b_pw) < 1 and abs(r_ph - b_ph) < 1
                            print(f"      [找回纸张·办法A] 设回「{b_canon}」→ ({r_pw:.2f},{r_ph:.2f})"
                                  f"{'（成功）' if ok else '（不行 —— 那是原绘图仪的纸名，PDF 打印机不认）'}")
                        except Exception as exc:               # noqa: BLE001
                            print(f"      [找回纸张·办法A] 设回「{b_canon}」失败：{exc}")
                    # 办法B：按尺寸匹配（跟出图代码同一条路）
                    hit = _cad_match_media(lay, b_pw, b_ph)
                    if hit is None:
                        print(f"      [找回纸张·办法B] 按尺寸 ({b_pw:.2f},{b_ph:.2f}) 没匹配到 —— 纸张只能保持默认")
                    else:
                        m_pw, m_ph = (float(v) for v in lay.GetPaperSize())
                        print(f"      [找回纸张·办法B] 按尺寸匹配到「{hit[0]}」"
                              f"{'（反向）' if hit[1] else ''} → ({m_pw:.2f},{m_ph:.2f})（成功，出图就用这条路）")
                    # 方向还原，别影响后续布局（原文件只读打开、Close 不保存，其实无所谓）
                    if b_rot is not None and a_rot is not None and a_rot != b_rot:
                        try:
                            lay.PlotRotation = b_rot
                        except Exception:                      # noqa: BLE001
                            pass
            else:
                print("      [换设备实测] 没找到名字带 PDF 的打印机，跳过")
    finally:
        try:
            doc.Close(False)                                  # 不保存，原件不动
        except Exception:                                     # noqa: BLE001
            pass
    # ⑥ -PUBLISH 批处理实测（2026-10-10 第二十三波）：这才是「复用图纸页面设置、
    #    不塌纸张」的正路 —— 试写一份 DSD、发 -PUBLISH，看浩辰认不认、出不出 PDF。
    _diag_publish(app, src)


def _diag_publish(app, src: Path) -> None:
    """-PUBLISH 路结论（2026-10-10 真机定案：浩辰不认命令行 DSD，只发结论不再实测）。

    ⭐ 真机踩坑：`SendCommand('-PUBLISH "xxx.dsd"')` 在浩辰上会**弹「发布」对话框
       等人选 DSD 文件**（它不认命令行里跟的路径参数），卡住体检流程。所以这里
       只打印一句结论，不再真的发命令 —— 别让体检每次卡在选文件那步。
    """
    print("\n  ── [ -PUBLISH 结论 ] ──")
    print("      ⚠️ 浩辰不认命令行 DSD 路径，-PUBLISH 会弹窗等人点 → 这条路已弃用，出图走 COM 路。")


def _run_cad_orientation_diag(src: Path) -> None:
    """独立跑一次「方向诊断」—— 自己负责 COM 初始化与收尾。

    ⭐ 为什么要独立包一层（2026-10-10 真机踩到）：`--check-convert` 转 dwg 走的是
       `_run_cad_many`，它在 `finally` 里 `CoUninitialize()` 了；转完再 `Dispatch`
       会报「尚未调用 CoInitialize」。这里**自己 CoInitialize → 开会话 → 诊断 →
       CoUninitialize**，不依赖调用方的 COM 状态。
    """
    try:
        import pythoncom                                     # noqa: PLC0415
    except Exception as exc:                                 # noqa: BLE001
        print(f"\n（方向诊断缺 pywin32：{exc}）")
        return
    hit = find_cad()
    if hit is None:
        print("\n（方向诊断：这台机器上没找到 CAD）")
        return
    pythoncom.CoInitialize()
    try:
        app = _open_cad_session(hit[1])
        _show_cad_orientation(app, src)
    except Exception as exc:                                 # noqa: BLE001
        print(f"\n（方向诊断开不了 CAD：{exc}）")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:                                    # noqa: BLE001
            pass


def _main_check_convert(a) -> int:
    """一键自检 —— 本机到底能不能转、不能的话是为什么。

    什么时候用：转换失败了但看不出原因（尤其是 macOS 的 -10004）。
    它跟正常流程走**同一条路**（同一个 pick_engine、同一个 engine.run），
    所以它报什么，界面上跑出来就是什么。
    """
    print(f"{TOOL_NAME} v{TOOL_VERSION} —— 转换自检")
    print(f"本机平台：{sys.platform}    Python：{sys.version.split()[0]}")

    cand: Path | None = None
    if (a.check_convert or "").strip():
        cand = Path(a.check_convert).expanduser()
    elif a.src:
        files, _unknown = collect_convertibles(a.src)
        cand = files[0] if files else None
    if cand is None or not cand.is_file():
        print("给我一份要试的 " + convert_kind_words() + "，二选一：")
        print("  python seq_core.py --check-convert ~/Desktop/某份.docx")
        print("  python seq_core.py --check-convert --src ~/Desktop/某文件夹")
        return 2

    fmt = format_of(cand)
    if fmt is None:
        print(f"❌ 认不出这个后缀：{cand.suffix!r}")
        print(f"  只认这些：{'、'.join(CONVERT_EXTS)}")
        return 2

    # ⭐ 引擎按**这一份属于哪一族**挑（见 ⑩a）—— 拿一份 .dwg 来自检，
    #    就该去挑 CAD 那条路，而不是把 Office 那个首选端出来（它根本转不了）。
    #    ⚠️ 也不能再挑「全平台首选」：Mac 上那条 office_win 压根不该被端出来。
    engine = pick_engine(fmt.family)
    if engine is None:
        print(f"❌ 没有可用的转换引擎（{family_label(fmt.family)} 这一族）。")
        print("  " + engine_no_engine_msg(fmt.family).replace("\n", "\n  "))
        return 2
    print(f"选中引擎：{engine.label}（{engine.key}）")

    # ⭐ 行式统一走 bootui（跟启动器、体检表一个样），别在这里自造 ✅/❌。
    B.item(B.OK, B.pad("试转", 16), cand.name)
    B.item(B.OK, B.pad("格式", 16), f"{fmt.label}（{fmt.key}）")
    tmp = Path(tempfile.mkdtemp(prefix="pdfseq_check_"))
    dst = tmp / f"{cand.stem}.pdf"
    # ⚠️ 自检也走 ⑩b：不然「带来自网络标记的文件」在这儿会照样失败，
    #    而真实流程已经被预处理救活了 —— 自检就成了误报源。
    feed, why = prep_source(cand, tmp)
    if why:
        B.item(B.SKIP, B.pad("预处理", 16), f"这份{why}")
        B.note("→ 先复制一份干净副本再转（原件没动）：%s" % feed.name)
    try:
        err = engine.run(feed, dst, fmt)
        if not err and not dst.is_file():
            err = "引擎没报错，但目标位置没有出现 PDF"
        if not err:
            n, _sz, rerr = read_pdf_info(dst)
            B.item(B.OK, B.pad("结果", 16),
                   f"转出 {dst.stat().st_size} 字节、{n} 页")
            if rerr:
                B.note("页数读取提示：%s" % rerr)
            # ⭐ 假失败现场（2026-10-23）：转是转成了，只是收尾时 Office 顶了一句。
            #    界面上不会显示它（用户不用看），但排查时这句原文最值钱 —— 打出来。
            if WIN_SALVAGED_NOTES:
                B.note("收尾时 Office 报了一句（不影响 PDF，已自动忽略）：")
                for _n in WIN_SALVAGED_NOTES[-5:]:
                    B.note("  " + _n)
            # ⭐ Visio 用的是哪一套参数（见 _visio_export_pdf：参数个数是它唯一的坑）。
            #    真机上换个版本又出问题时，先看这行 —— 走的可能是退化写法。
            if fmt.family == "visio" and VISIO_EXPORT_USED:
                B.note("Visio 导出走的是：%s" % VISIO_EXPORT_USED[-1])
            # ⭐ 转的是 CAD 图纸 → 顺带做一次「方向诊断」，把每个布局的纸张/内容/方向
            #    打出来，一眼看出「横竖」对不对（拖文件到体检.bat 时最有用）。
            if fmt.family == "cad":
                _run_cad_orientation_diag(cand)
            return 0
        code = classify_office_error(err)
        B.item(B.FAIL, B.pad("结果", 16), f"失败（归类：{code}）")
        B.note(err)
        return 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _main_convert(a) -> int:
    """命令行：格式转换模式。**不要 --scan、不要 --suffix**（后缀可空）。"""
    opts = Options(sources=tuple(a.src), suffix=a.suffix,
                   order_rule=a.order, dry_run=a.dry_run,
                   overwrite=a.overwrite)
    plan = plan_convert(opts)
    if plan.blocked:
        print("整批不执行：")
        print("  " + plan.block_reason.replace("\n", "\n  "))
        return 2
    if plan.engines:
        print("转换引擎：" + "、".join(f"{e.label}（{e.key}）" for e in plan.engines))
    _fb: list[str] = []
    for _e in plan.engines:                       # 兜底梯队**按族**算（见 ⑩a）
        _fb += [x.label for x in engine_chain(_e)[1:]]
    if _fb:
        print("（首选失手时会自动接着用：" + "、".join(dict.fromkeys(_fb)) + "）")
    for fam in plan.missing_families:
        print(f"⚠️ {family_label(fam)} 这一族本机没有程序能转，这部分会单独标出来。")
    print(f"转换规则：{CONVERT_RULE_NOTE}{page_rule_label()}")
    if plan.unknown:
        print(f"（清单里有 {len(plan.unknown)} 份不是 {convert_kind_words()}，已跳过）")
    n_group = len({it.fmt_key for it in plan.items})
    print(f"待转换：{len(plan.items)} 份（按文件名排）—— 同格式合成 {n_group} 批，"
          f"每批只拉起一次对应的 Office 应用（不会一份一份开）")
    for it in plan.items:
        print(f"  {it.index:>3}. {it.src.name}  →  {out_pdf_name_for(it.src, opts.suffix)}")
    will = convert_overwrite_list(plan, opts)
    if will:
        print(f"有 {len(will)} 份同名 PDF 已经存在 —— 按覆盖口径处理"
              f"（{'覆盖' if opts.overwrite is OVERWRITE_YES else '加编号让开' if opts.overwrite is OVERWRITE_NO else '认得出是上一轮产物的原地更新，其余加编号'}）：")
    results, notes = run_convert(plan, opts)
    for n in notes:
        print(n)
    if not results:
        return 2
    ok = sum(1 for r in results if r.status == ConvStatus.OK.value)
    bad = [r for r in results if r.status == ConvStatus.FAILED.value]
    print(f"完成 {ok} 份，失败 {len(bad)} 份")
    used_engines = {r.engine_key for r in results if r.engine_key}
    if used_engines - {e.key for e in plan.engines}:
        names = "、".join((engine_by_key(k).label if engine_by_key(k) else k)
                          for k in sorted(used_engines))
        print(f"实际干活的引擎：{names}（首选失手的那几份，兜底梯队接的手）")
    for r in bad:
        print(f"  [失败] {r.src.name} —— {r.note}")
    return 0 if not bad else 1


def main(argv: Sequence[str] | None = None) -> int:
    a = _build_parser().parse_args(argv)
    if a.show_rules:                       # 纯说明，优先于一切
        return _main_show_rules(a)
    if a.show_cad:                         # 纯说明，不碰任何文件
        return _main_show_cad(a)
    if a.show_bartender:                   # 纯说明，不碰任何文件
        return _main_show_bartender(a)
    if a.check_convert is not None:        # 自检优先于一切（它就是来诊断的）
        return _main_check_convert(a)
    if a.mode == "convert":
        return _main_convert(a)
    try:
        rots = parse_rotations(a.rotate)
    except ValueError as exc:
        print(f"旋转设置写错了：{exc}")
        return 2
    opts = Options(sources=tuple(a.src), scan_pdf=a.scan,
                   reverse_order=a.reverse, scale_to_original=not a.no_scale,
                   overwrite=a.overwrite,
                   suffix=a.suffix, order_rule=a.order, dry_run=a.dry_run,
                   rotations=rots, auto_rotate=not a.no_auto_rotate)
    why = suffix_error(opts)
    if why:
        print(why)
        return 2
    plan = plan_batch(opts)
    if plan.blocked:
        print("整批不执行：")
        print("  " + plan.block_reason.replace("\n", "\n  "))
        return 2
    if plan.products:
        print(f"（自动跳过 {len(plan.products)} 份上一轮留下的产物）")

    if not plan.paired:
        print(f"待替换队列：{len(plan.items)} 份（按文件名排）")
        for it in plan.items:
            print(f"  {it.index:>3}. {it.src}")
        print("（没给 --scan，只列队不执行。照着这个顺序摊开纸质件去扫，"
              "扫完把那份 PDF 用 --scan 传进来。）")
        return 0

    print(f"原件 {len(plan.items)} 份、扫描件 {plan.scan_pages} 页；"
          f"产物写在各自原件旁边（{out_name_for(plan.items[0].src, opts.suffix)} 这种名字）")
    will = overwrite_list(plan, opts)
    if will:
        if opts.overwrite is OVERWRITE_YES:
            print(f"有 {len(will)} 份产物会撞上已有文件 —— 已给 --overwrite，直接覆盖：")
        elif opts.overwrite is OVERWRITE_NO:
            print(f"有 {len(will)} 份产物会撞上已有文件 —— 已给 --no-overwrite，全部加编号避开：")
        else:
            print(f"有 {len(will)} 份产物会撞上已有文件 —— 认得出是上一轮产物的原地更新，"
                  f"认不出的加编号避开：")
        for t, s in will[:10]:
            print(f"  {s.name} → {t.name}")
    results, _notes = run_batch(plan, opts)
    ok = sum(1 for r in results if r.status == Status.OK.value)
    bad = [r for r in results if r.status == Status.FAILED.value]
    print(f"完成 {ok} 份，失败 {len(bad)} 份")
    for r in bad:
        print(f"  [失败] {r.src.name} —— {r.note}")
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
