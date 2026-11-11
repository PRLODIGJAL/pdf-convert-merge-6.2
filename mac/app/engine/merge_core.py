# -*- coding: utf-8 -*-
"""merge_core.py —— 合并打印 · 引擎层（6.0 新增）
=================================================
用户场景（承自项目 i4「PDF 奇偶拼接打印」）：

    打印机没有自动翻页 → 一批作业 / 报告（Word / Excel，也有现成 PDF）
    要拼成**一份**大 PDF，拿去做手动双面打印。
    双面模式下，奇数页的那份末尾补一张空白页凑偶数 ——
    这样每一份都从纸的**正面**开始，不会串到上一份的背面。

它在 6.0 里的位置
------------------
跟「导出 PDF」是**同一条转换链路**，区别只在下游：

    导出 PDF ：  转换 → 每份 PDF 落在原件旁边（用户自己拿）
    合并打印 ：  转换 → 补白 → 合成一份 → 写到你选的输出路径

所以这里**一行转换代码都不写**：转换全部交给 `seq_core.run_convert`
（引擎梯队、Office 批量会话、网络标记副本、Excel 一表一页规则全都白捡）。
本文件只补三件事：**补白**、**合一**、**把中间产物关进临时目录**。

⚠️ 三条硬约定
--------------
1. **中间 PDF 不落在用户文件夹里**：`Options.out_dir` 指向临时目录，用完即删。
   用户在「导出 PDF」里想要的「产物落在原件旁边」，在这里恰恰是噪音。
2. **整批失败就不出产物**：任何一份转不出来 / 读不出页数，直接放弃、不写半份文件。
   理由和 ① 表的「顺序配对是脆弱链」一样 —— 少一份的合集，打印出来才发现，比报错糟。
3. **`engine` 参数一路透传**：Mac 上 Office 转换不稳定，回归只能靠注入假引擎跑。
   不留这个口子，整条链路在开发机上一条用例都跑不了。

依赖：只用 `pypdf`（与 seq_core 同一套），**不引入 pymupdf**。
     少一个依赖，也避开 AGPL 授权在内部发包时的解释成本。
"""

from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass, field, replace
from enum import Enum
from pathlib import Path
from typing import Callable, Sequence

import seq_core as C


# ══════════════════════════════════════════════════════════════════════════
# ① 打印模式
# ══════════════════════════════════════════════════════════════════════════

class PrintMode(str, Enum):
    """打印方式 —— 用户按自家打印机有没有自动翻页来选。"""

    DUPLEX = "duplex"      # 双面：奇数页文档末尾补 1 张空白，凑成偶数
    SIMPLEX = "simplex"    # 单面：直拼，一页都不补
    # ⭐ 6.2 新增：首页单面、其余双面。
    #    首页常要单独签字 / 盖章 / 写分发号，背面必须是空的；后面几十页仍要省纸。
    #    所以第 1 页之后**强制**插一张空白 —— 不插的话正文会落到首页的背面去。
    HOME_SIMPLEX = "home_simplex"


PRINT_MODE_LABELS: dict[str, str] = {
    PrintMode.DUPLEX.value: "双面打印（奇数页补空白）",
    PrintMode.SIMPLEX.value: "单面打印（不补白）",
    PrintMode.HOME_SIMPLEX.value: "首页单面（其余双面）",
}
DEFAULT_PRINT_MODE = PrintMode.DUPLEX.value


class PageOrient(str, Enum):
    """页面方向 —— 一批里横竖混着的时候，统一成同一种再拼。

    ⚠️ 为什么要有这个开关：横向页夹在一摞竖版里，打印出来是一张躺着的纸，
       装订完得把整本转 90° 才看得见那几页。统一方向之后纸张一致，
       代价是横版那几页的内容会侧躺（内容没法重排，只能转纸）。

    ⚠️ 默认是 `KEEP`（不动），不默认开 —— 一开就会改变老口径的产物，
       6.1 那批历史合集会跟新拼的对不上。
    """

    KEEP = "keep"            # 保持原样（横竖混着就混着）
    PORTRAIT = "portrait"    # 全部转成竖向
    LANDSCAPE = "landscape"  # 全部转成横向


ORIENT_LABELS: dict[str, str] = {
    PageOrient.KEEP.value: "保持原样",
    PageOrient.PORTRAIT.value: "全部转竖向",
    PageOrient.LANDSCAPE.value: "全部转横向",
}
DEFAULT_ORIENT = PageOrient.KEEP.value

# 默认输出名 —— 与 i4 一致，用户一眼知道这是什么。
DEFAULT_OUT_NAME = "合集打印.pdf"

# PDF 行的状态词。
# ⚠️ 不用「转换成功」：那一行**根本没转**，写「转换成功」是在对用户说假话；
#    界面配色表里给它单独一档（见 seq_app.STATUS_COLORS）。
STATUS_READY = "可直接拼"


# ══════════════════════════════════════════════════════════════════════════
# ② 认哪些输入
# ══════════════════════════════════════════════════════════════════════════
# ⭐ 合并 = 「导出 PDF」的全部输入 **再加 PDF 本身**。
#    18 种可转格式原样继承（加格式只动 seq_core 的 ⑩ 表，这里自动跟），
#    PDF 单独放进来：它不用转，直接进拼接队列。
MERGE_EXTS: frozenset[str] = frozenset(C.CONVERT_EXTS) | frozenset(C.PDF_EXTS)


def is_merge_input(p: Path) -> bool:
    """这一份能不能进合并队列（认扩展名，不读内容）。"""
    return p.suffix.lower() in MERGE_EXTS


def is_pdf(p: Path) -> bool:
    return p.suffix.lower() in C.PDF_EXTS


# ══════════════════════════════════════════════════════════════════════════
# ③ 数据模型
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class MergeItem:
    """合并队列里的一份。一个来源文件 = 一行。"""

    index: int
    src: Path
    # 这一份走哪条路："" = 本来就是 PDF（直拼）；否则是 ⑩ 表的格式 key（要转）
    fmt_key: str = ""
    fmt_label: str = ""
    # 状态词直接沿用转换那套（转换成功 / 转换失败 / 已跳过 / 待转换），
    # 界面上的配色表也只认这一套 —— 别再发明第二组词。
    status: str = C.ConvStatus.PENDING.value
    note: str = ""
    # 列队那一刻的说明（「PDF 直接拼」「本机没有能转它的程序」…）。
    # ⭐ 每次开跑都要把 `note` 复位成它 —— 否则上一轮写进去的「补白失败」
    #    会赖在表格里，用户改了设置再跑，看到的还是上一轮的死因。
    plan_note: str = ""
    pages: int = 0          # 补白**之后**的页数（打印出来占几页）
    padded: bool = False    # 双面口径下这份末尾是不是补了一张白页
    # ⭐ 6.2：首页单面口径下，第 1 页之后插的那张空白（跟 `padded` 分开记 ——
    #    表格里要说清是「首页后插的」还是「末尾补的」，用户对账时看的不是同一个位置）
    pad_front: bool = False
    rotated: int = 0        # 这份转了几页（0 = 一页没动）
    pdf: Path | None = None  # 拿来拼的 PDF（临时目录里的转换产物，或原件本身）


@dataclass
class MergePlan:
    """一次「列队」的结果。blocked=True 时一份都不拼。"""

    items: list[MergeItem] = field(default_factory=list)
    unknown: list[Path] = field(default_factory=list)
    blocked: bool = False
    block_reason: str = ""
    engines: list[C.ConvEngine] = field(default_factory=list)
    missing_families: list[str] = field(default_factory=list)

    @property
    def blank_added(self) -> int:
        """一共补了几张白纸（首页后插的 + 末尾补的，都算）。"""
        return sum((1 if it.pad_front else 0) + (1 if it.padded else 0)
                   for it in self.items)

    @property
    def rotated(self) -> int:
        """一共转了几页（方向开关没开时恒为 0）。"""
        return sum(it.rotated for it in self.items)

    @property
    def total_pages(self) -> int:
        return sum(it.pages for it in self.items)


@dataclass
class MergeResult:
    """跑完一轮的结果。`ok=False` 时 `out` 一定不存在（不留半份文件）。"""

    ok: bool = False
    out: Path | None = None
    message: str = ""
    total_pages: int = 0
    blank_added: int = 0
    converted: int = 0
    rotated: int = 0        # 6.2：方向开关转了几页
    plan: MergePlan | None = None


# ══════════════════════════════════════════════════════════════════════════
# ④ 收集队列（口径与 seq_core.collect_convertibles 一致，只是多认 PDF）
# ══════════════════════════════════════════════════════════════════════════

def collect_merge_sources(sources: Sequence[str | Path],
                          order_rule: str = C.DEFAULT_ORDER
                          ) -> tuple[list[Path], list[Path]]:
    """把来源清单收成一队。返回 (认得的, 认不得的)。

    - 文件夹 → 只收它**这一层**里认得的文件，不往子目录钻（与 core 同一口径）；
    - 单个文件 → 直接用，认不出格式的进 `unknown` 只统计不进队；
    - 排序沿用 ① 表（默认按文件名），跟「导出 PDF」出来的顺序对得上。

    Args:
        sources: 界面上的来源清单（文件或文件夹路径，可混合）。
        order_rule: 排序规则 key，见 seq_core.ORDER_RULES。

    Returns:
        (可进队的路径, 认不出格式的路径)
    """
    hit: list[Path] = []
    unknown: list[Path] = []
    seen: set[Path] = set()
    for s in sources:
        entry = Path(s).expanduser()
        if entry.is_dir():
            cands = [c for c in sorted(entry.glob("*"))
                     if c.is_file() and c.suffix.lower() in MERGE_EXTS]
        elif entry.is_file():
            cands = [entry]
        else:
            continue                       # 不存在 → 交给闸门去报
        for p in cands:
            try:
                pr = p.resolve()
            except OSError:
                continue
            if pr in seen:
                continue                   # 同一份被指到两次
            seen.add(pr)
            if not is_merge_input(p):
                unknown.append(p)          # 单文件来源里挑了个 .txt 之类
                continue
            hit.append(p)
    return C.sort_originals(hit, order_rule), unknown


def _convert_subset(items: Sequence[MergeItem],
                    options: C.Options) -> C.ConvertPlan:
    """把「要转的那几份」拼成一份 seq_core 的 ConvertPlan。

    ⚠️ 只喂可转格式：PDF 若混进去，`run_convert` 会判「认不出这是什么格式」而报错。
    """
    sub = replace(options,
                  sources=tuple(str(it.src) for it in items),
                  suffix="",                       # 中间产物不留后缀，纯原件名
                  out_dir=str(options.out_dir or ""))
    return C.plan_convert(sub)


def plan_merge(options: C.Options) -> MergePlan:
    """把来源清单算成一份待拼队列。**不读文件内容**，只认扩展名。

    ⭐ 转换那一半完全交给 `seq_core.plan_convert` —— 闸门（⑫）、引擎梯队、
       「这台机器上没有能转它的程序」的逐行说明，全都复用，不另写一套。
    """
    hits, unknown = collect_merge_sources(options.sources, options.order_rule)
    plan = MergePlan(unknown=list(unknown))

    if not hits:
        plan.blocked = True
        plan.block_reason = "没有可合并的文件。"
        return plan

    pdfs = [p for p in hits if is_pdf(p)]
    conv = [p for p in hits if not is_pdf(p)]

    # ── 转换那一半：先列队，拿到每份的格式 / 引擎 / 逐行说明 ──
    cplan = _convert_subset([MergeItem(index=0, src=p) for p in conv], options) \
        if conv else C.ConvertPlan()
    if cplan.blocked:
        plan.blocked = True
        plan.block_reason = cplan.block_reason or "还不能合并。"
        return plan
    plan.engines = list(cplan.engines)
    plan.missing_families = list(cplan.missing_families)

    by_src = {it.src: it for it in cplan.items}

    # ── 按 hits 的原顺序铺成一张表（PDF 与可转格式混排也要保持顺序）──
    for i, p in enumerate(hits, start=1):
        if is_pdf(p):
            plan.items.append(MergeItem(
                index=i, src=p, fmt_label="PDF",
                status=STATUS_READY,
                note="PDF 直接拼，不用转", plan_note="PDF 直接拼，不用转"))
            continue
        ci = by_src.get(p)
        plan.items.append(MergeItem(
            index=i, src=p,
            fmt_key=(ci.fmt_key if ci else ""),
            fmt_label=(ci.fmt_label if ci else "?"),
            note=(ci.note if ci else ""),
            plan_note=(ci.note if ci else "")))

    if all(it.status == C.ConvStatus.SKIPPED.value for it in plan.items):
        plan.blocked = True
        plan.block_reason = "队列里每一行都被忽略了，没有可拼的东西。"
    return plan


# ══════════════════════════════════════════════════════════════════════════
# ⑤ PDF 操作：补白、合一
# ══════════════════════════════════════════════════════════════════════════

def _display_size(page: "object") -> tuple[float, float]:
    """页面**显示**尺寸：/Rotate 90/270 的页面按转过来之后的宽高返回。

    ⚠️ 这就是「补白页跟末页一样大」的关键：横向文档（/Rotate 90）用 mediabox
       原样补，补出来会是一张**竖的**白纸，打印时跟前面几张对不齐。
       i4 当年用 pymupdf 的 rect 天然规避了这个问题，换到 pypdf 必须自己算。
    """
    box = page.mediabox                       # type: ignore[attr-defined]
    w, h = float(box.width), float(box.height)
    rot = int(getattr(page, "rotation", 0) or 0) % 360
    return (h, w) if rot in (90, 270) else (w, h)


def pad_to_even(src_pdf: Path, dst_pdf: Path) -> bool:
    """奇数页 → 末尾补一张同尺寸空白页凑成偶数。返回是否补了。

    6.0 的口径（全双面）。6.2 起真正干活的是 `prepare_pdf`，这个函数留着是因为
    回归脚本直接调它，而且「只补末尾」这个动作本身还是独立可用的。

    Args:
        src_pdf: 源 PDF（不会被修改）。
        dst_pdf: 产物路径（父目录必须已存在）。

    Returns:
        True = 这一份补了一张白页；False = 本来就是偶数页，原样复制。

    Raises:
        Exception: 读不出 / 写不了时向上抛，由 run_merge 统一兜成「整批不出」。
    """
    from pypdf import PdfWriter

    reader = C._reader(src_pdf)               # 复用 core 的解密口径，别另写一份
    pages = list(reader.pages)
    writer = PdfWriter()
    for p in pages:
        writer.add_page(p)
    odd = len(pages) % 2 == 1
    if odd:
        w, h = _display_size(pages[-1])
        writer.add_blank_page(width=w, height=h)
    with open(dst_pdf, "wb") as fh:
        writer.write(fh)
    return odd


# ── 6.2：方向归一 + 按口径补白，合成一次读写 ────────────────────────────

def _is_landscape(page: "object") -> bool:
    """这一页**显示**出来是横的吗（已算上 /Rotate）。"""
    w, h = _display_size(page)
    return w > h


def _is_square(page: "object") -> bool:
    """正方形页 —— 既不算横也不算竖，别去转它。"""
    w, h = _display_size(page)
    return abs(w - h) < 1.0


@dataclass
class PrepInfo:
    """`prepare_pdf` 干了什么 —— 表格和结果提示都从这里读。"""

    front: bool = False      # 首页之后插了一张空白
    tail: bool = False       # 末尾补了一张空白（凑偶数）
    rotated: int = 0         # 转了几页

    @property
    def added(self) -> int:
        return (1 if self.front else 0) + (1 if self.tail else 0)


def prepare_pdf(src_pdf: Path, dst_pdf: Path, mode: str,
                orient: str = DEFAULT_ORIENT) -> PrepInfo:
    """转向 → 补白 → 写出。**顺序不能反**。

    ⚠️ 为什么必须「先转再补」：空白页的尺寸是照着**相邻页转向之后**的显示尺寸做的。
       反过来先补再转，那张白纸会按原始方向生成，转完就成了竖白纸夹在横版里，
       打印时跟前后几张对不齐 —— 6.0 在「横向文档补白」上踩过一模一样的坑。

    Args:
        src_pdf: 源 PDF（不会被修改）。
        dst_pdf: 产物路径（父目录必须已存在）。
        mode: `PrintMode` 的 value。
        orient: `PageOrient` 的 value。

    Returns:
        PrepInfo（插了首页空白没 / 补了末尾空白没 / 转了几页）。

    Raises:
        Exception: 读不出 / 写不了时向上抛，由 run_merge 统一兜成「整批不出」。
    """
    from pypdf import PdfWriter

    reader = C._reader(src_pdf)
    pages = list(reader.pages)
    if not pages:
        raise ValueError("这份 PDF 一页都没有")
    info = PrepInfo()

    # ── ① 方向归一 ───────────────────────────────────────────────────
    # ⚠️ 用 `rotate(90)`（/Rotate 加 90，即顺时针），不是重排内容：
    #    横→竖、竖→横都靠它一次到位，原来那些 /Rotate 90/270 的扫描件也不会被转乱。
    if orient != PageOrient.KEEP.value:
        want_landscape = orient == PageOrient.LANDSCAPE.value
        for p in pages:
            if _is_square(p) or _is_landscape(p) == want_landscape:
                continue
            p.rotate(90)                      # type: ignore[attr-defined]
            info.rotated += 1

    writer = PdfWriter()
    writer.add_page(pages[0])

    # ── ② 首页单面：第 1 页之后**强插**一张空白 ──────────────────────
    #    首页单独走单面之后，正文必须落在下一张纸的正面；
    #    中间不空这一页，正文就印到首页背面去了 —— 那正是用户要避免的。
    if mode == PrintMode.HOME_SIMPLEX.value:
        w, h = _display_size(pages[0])
        writer.add_blank_page(width=w, height=h)
        info.front = True

    for p in pages[1:]:
        writer.add_page(p)

    # ── ③ 末尾凑偶数 ─────────────────────────────────────────────────
    #    双面口径和首页单面口径都要凑（首页单面插完那一页之后，剩下的正文仍然双面印）。
    #    单面口径一页都不补。
    if mode in (PrintMode.DUPLEX.value, PrintMode.HOME_SIMPLEX.value):
        if len(writer.pages) % 2 == 1:
            w, h = _display_size(pages[-1])
            writer.add_blank_page(width=w, height=h)
            info.tail = True

    with open(dst_pdf, "wb") as fh:
        writer.write(fh)
    return info


def merge_pdfs(pdf_paths: Sequence[Path], out_path: Path) -> int:
    """把多份 PDF 按给定顺序合成一份。返回总页数。

    ⚠️ 先写 `.part` 再原子替换：中途断电 / 磁盘满，不会把用户上一轮的同名产物
       毁掉一半。`os.replace` 在同一个目录内是原子的。

    Args:
        pdf_paths: 要拼的 PDF，顺序即最终页序。
        out_path: 输出路径（父目录不存在会自动建）。

    Returns:
        合并后的总页数。
    """
    from pypdf import PdfWriter

    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(str(p))
    total = len(writer.pages)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = out_path.with_name(out_path.name + ".part")
    try:
        with open(tmp, "wb") as fh:
            writer.write(fh)
        os.replace(tmp, out_path)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
    return total


# ══════════════════════════════════════════════════════════════════════════
# ⑥ 主流程：转换 → 补白 → 合一
# ══════════════════════════════════════════════════════════════════════════

ProgressFn = Callable[[int, int, str], None]
StopFn = Callable[[], bool]


class _Stopped(Exception):
    """用户点了「停止」。内部信号，不出这个模块。"""


def _noop_progress(done: int, total: int, msg: str) -> None:  # pragma: no cover
    return None


def run_merge(plan: MergePlan, options: C.Options, out_path: Path,
              mode: str = DEFAULT_PRINT_MODE,
              orient: str = DEFAULT_ORIENT,
              engine: C.ConvEngine | None = None,
              should_stop: StopFn | None = None,
              progress: ProgressFn | None = None) -> MergeResult:
    """转换 → 转向 → 补白 → 合成一份，写到 `out_path`。

    Args:
        plan: `plan_merge()` 的结果。
        options: 与界面共用的一套选项（`out_dir` 会被临时目录覆盖）。
        out_path: 用户选的输出 PDF 路径。
        mode: `PrintMode` 的 value。
        orient: `PageOrient` 的 value（保持原样 / 全部竖向 / 全部横向）。
        engine: 可注入的假引擎 —— **Mac 上跑回归靠它**，产线不要传。
        should_stop: 每份之间问一次「要不要停」。
        progress: (已完成, 总数, 说明) 回调。

    Returns:
        MergeResult。失败时 `out` 一定为 None，且不会留下半份文件。

    Raises:
        None: 所有异常都兜成 MergeResult 的失败信息 —— 调用方是界面线程，
              抛出去只会让工作线程静默死掉、界面永久卡在「忙」。
    """
    report = progress or _noop_progress
    stop = should_stop or (lambda: False)
    duplex = mode == PrintMode.DUPLEX.value
    home = mode == PrintMode.HOME_SIMPLEX.value
    # 只要有一件事要做（补白 / 转向），就得生成一份中间 PDF。
    # 单面 + 不转向是最常见的组合，这时一份中间文件都不写 —— 跟 6.0 一样快。
    need_prepare = duplex or home or orient != PageOrient.KEEP.value

    if plan.blocked:
        return MergeResult(message=plan.block_reason or "被阻断，未执行", plan=plan)

    todo = [it for it in plan.items if it.status != C.ConvStatus.SKIPPED.value]
    if not todo:
        return MergeResult(message="队列里没有要拼的文件。", plan=plan)

    total_steps = len(todo) + 1
    # ⭐ 每一轮都从干净状态起跑：`pdf` 指向上一次那个**已经被删掉的临时目录**，
    #    不复位的话，第二次跑（比如从双面改成单面再跑一遍）会拿一个不存在的路径
    #    去补白，报一个完全看不懂的错。状态与说明同理复位。
    for it in plan.items:
        if it.status == C.ConvStatus.SKIPPED.value:
            continue
        it.pdf = None
        it.pages = 0
        it.padded = False
        it.pad_front = False
        it.rotated = 0
        it.note = it.plan_note
        it.status = STATUS_READY if is_pdf(it.src) else C.ConvStatus.PENDING.value
    workdir = Path(tempfile.mkdtemp(prefix="pdf_merge_"))
    conv_count = 0
    try:
        # ── ① 转换（只转可转格式；PDF 跳过这一步）───────────────
        conv_items = [it for it in plan.items
                      if not is_pdf(it.src)
                      and it.status != C.ConvStatus.SKIPPED.value]
        if conv_items:
            if stop():
                raise _Stopped
            report(0, total_steps, f"正在转换 {len(conv_items)} 份文件…")
            cplan = _convert_subset(conv_items, options)
            cplan.engines = list(plan.engines)
            sub_options = replace(options, out_dir=str(workdir), suffix="")
            results, _notes = C.run_convert(cplan, sub_options, engine)

            by_src = {r.src: r for r in results}
            first_bad: MergeItem | None = None
            for it in conv_items:
                r = by_src.get(it.src)
                if r is None:
                    it.status = C.ConvStatus.FAILED.value
                    it.note = "没有拿到转换结果"
                    first_bad = first_bad or it
                    continue
                it.status = r.status
                if r.note:
                    it.note = r.note
                if r.ok and r.status == C.ConvStatus.OK.value and r.out:
                    it.pdf = r.out
                    it.pages = r.pages
                    conv_count += 1
                elif r.status != C.ConvStatus.SKIPPED.value:
                    it.status = C.ConvStatus.FAILED.value
                    first_bad = first_bad or it
            if first_bad is not None:
                # 整批不出：少一份的合集，打印时才发现，比当场报错糟得多。
                return MergeResult(
                    message=f"转换失败：{first_bad.src.name}\n"
                            f"{first_bad.note or '原因见表格「说明」列'}",
                    plan=plan, converted=conv_count)

        # ── ② 补白（双面模式才补）+ 读出每份的页数 ─────────────
        step = 0
        for it in todo:
            if stop():
                raise _Stopped
            step += 1
            if it.pdf is None:                       # 原本就是 PDF
                it.pdf = it.src
                n, _sz, err = C.read_pdf_info(it.pdf)
                if n == 0:
                    it.status = C.ConvStatus.FAILED.value
                    it.note = f"读不出这一份（{err}）"
                    return MergeResult(message=f"读不了：{it.src.name}\n{err}",
                                       plan=plan, converted=conv_count)
                it.pages = n
                it.status = STATUS_READY

            if need_prepare:
                report(step, total_steps, f"规整 {it.src.name} …")
                prepped = workdir / f"{it.index:03d}_{it.src.stem}_prep.pdf"
                try:
                    info = prepare_pdf(it.pdf, prepped, mode, orient)
                except Exception as exc:  # noqa: BLE001 —— 坏 PDF 不该炸掉整个线程
                    it.status = C.ConvStatus.FAILED.value
                    it.note = f"规整失败：{exc}"
                    return MergeResult(message=f"补白 / 转向失败：{it.src.name}\n{exc}",
                                       plan=plan, converted=conv_count)
                it.pad_front = info.front
                it.padded = info.tail
                it.rotated = info.rotated
                it.pages += info.added
                it.pdf = prepped
            else:
                report(step, total_steps, f"检查页数 {it.src.name} …")

        # ── ③ 合一 ──────────────────────────────────────────────
        if stop():
            raise _Stopped
        report(total_steps, total_steps, "正在合成一份…")
        try:
            pages = merge_pdfs([it.pdf for it in todo if it.pdf], out_path)
        except Exception as exc:  # noqa: BLE001
            return MergeResult(message=f"拼接失败：{exc}", plan=plan,
                               converted=conv_count)

        report(total_steps, total_steps, "完成")
        msg = f"已生成：{out_path.name}（{pages} 页"
        extras: list[str] = []
        if plan.blank_added:
            extras.append(f"补了 {plan.blank_added} 张空白")
        if plan.rotated:
            extras.append(f"转了 {plan.rotated} 页方向")
        if extras:
            msg += "，" + "、".join(extras)
        msg += "）"
        return MergeResult(ok=True, out=out_path, message=msg,
                           total_pages=pages, blank_added=plan.blank_added,
                           converted=conv_count, rotated=plan.rotated,
                           plan=plan)

    except _Stopped:
        return MergeResult(message="已停止，没有生成合并文件。",
                           plan=plan, converted=conv_count)
    finally:
        # 中间 PDF 一份都不留（包括失败时）—— 用户的原件文件夹从此是干净的。
        shutil.rmtree(workdir, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════
# ⑦ 一句话口径（界面与说明书共用，别再手写）
# ══════════════════════════════════════════════════════════════════════════

def merge_kind_words(platform: str | None = None) -> str:
    """「合并认哪些文件」的说法 —— 从表 + 当前平台派生，别手写。"""
    return C.convert_kind_words(platform) + " / PDF"


def print_mode_words(mode: str = DEFAULT_PRINT_MODE) -> str:
    """把模式讲成一句人话（界面提示 / 说明书共用）。"""
    if mode == PrintMode.SIMPLEX.value:
        return "单面：直拼，一页都不补"
    if mode == PrintMode.HOME_SIMPLEX.value:
        return "首页单面：第 1 页之后插一张空白，其余照双面走，每份都从正面开始"
    return "双面：奇数页文档末尾补 1 张空白，每份都从正面开始、不串页"


def orient_words(orient: str = DEFAULT_ORIENT) -> str:
    """把方向开关讲成一句人话。

    ⚠️ 别忘了说「内容会侧躺」—— 用户看到横版那几页文字歪着，第一反应是程序坏了。
    """
    if orient == PageOrient.PORTRAIT.value:
        return "统一竖向：横版页转 90° 变竖版，整批纸张同向（横版那几页的内容会侧躺）"
    if orient == PageOrient.LANDSCAPE.value:
        return "统一横向：竖版页转 90° 变横版，整批纸张同向（竖版那几页的内容会侧躺）"
    return "保持原样：横的还横、竖的还竖，一张都不转"
