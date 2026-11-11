#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF转换拼接工具 v6.2 —— 环境自检与启动器（Windows / macOS 共用这一份）。

双击 .bat / .command 会调到这里，由它统一负责三件事：
  1. 检查工具自带的运行时（runtime/）还能不能用；
  2. 不能用就现场修复（重新识别：修权限、去隔离属性、补依赖）；
  3. 就绪后用 runtime 里的解释器拉起 seq_app.py。

6.0 起界面有**三个模式**（替换首页 / 批量转 PDF / 合并打印），
环境这边一个字都不用变 —— 合并打印补白 / 合一走的是 pypdf。

**从 6.1 起用内嵌运行时（runtime/），不再依赖系统 Python。**
整个文件夹拷到任何电脑、任何位置，双击「启动PDF转换拼接工具」就能用。

**打印格式统一走 bootui.py**：标记只有 `[ OK ] [ !! ] [ -- ] [ .. ]` 四种，
分节只用 `-- 名字 ----` 一行，宽度固定 62 列。**别在这里直接 print**。

为什么内嵌运行时
----------------
以前用 .venv（虚拟环境），它记着创建时那台机器的 Python 绝对路径，
二进制也不通用。拷到另一台电脑就废了。现在整个 Python 都在 runtime/ 里，
是可重定位的（relocatable），拷到哪都能跑。

「重新识别」做什么
------------------
  1. 修执行权限（拷到 FAT/exFAT 盘会丢 +x）
  2. 去 macOS 隔离属性（com.apple.quarantine，从网上下载/邮件传的会带）
  3. 验证 Python 能跑、tkinter 能起、依赖全在
  4. 缺依赖就从 wheels/ 离线装，没网也能修

用法
----
    python setup_env.py                 自检 → 修复 → 启动（正常双击走这条）
    python setup_env.py --report        运行前体检（只读）
    python setup_env.py --report --check-one 某文件
    python setup_env.py --check         只自检并打印结论
    python setup_env.py --setup         重新识别（修权限、补依赖、验证），不启动界面
    python setup_env.py --reidentify    同 --setup（更明确的名字）
"""

import os
import shutil
import subprocess
import sys

import bootui as B

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
RUNTIME = os.path.join(ROOT, "runtime")
WHEELS = os.path.join(ROOT, "wheels")
APP = os.path.join(ROOT, "app", "ui", "seq_app.py")
ENGINE_MAIN = os.path.join(ROOT, "app", "engine", "seq_core.py")

TOOL_NAME = "PDF转换拼接工具"
TOOL_VERSION = "6.2"

# ── 依赖清单 ──────────────────────────────────────────────────────────────
BASE_NEED = ("pypdf>=3.16", "tkinterdnd2>=0.3")
WIN_NEED = ("pywin32>=306",)
STD_IMPORTS = ("tkinter",)
PLATFORM_IMPORTS = {
    "nt": ("pythoncom", "win32com.client"),
    "posix": (),
}


def need_list():
    """这台机器上要装的依赖 —— Windows 多一份 pywin32。"""
    return list(BASE_NEED) + (list(WIN_NEED) if os.name == "nt" else [])


def check_imports():
    """自检要 import 的东西。"""
    key = "nt" if os.name == "nt" else "posix"
    mods = (["pypdf", "tkinterdnd2"] + list(PLATFORM_IMPORTS.get(key, ()))
            + list(STD_IMPORTS))
    return "import " + ", ".join(mods)


# ── 程序文件齐不齐 ─────────────────────────────────────────────────────────
PROGRAM_MODULES = ("paths", "bootui", "seq_core", "merge_core")


def module_probe_code():
    dirs = [os.path.join(ROOT, "app")]
    dirs += [os.path.join(ROOT, "app", d) for d in ("ui", "judge", "engine", "shell")]
    return ("import sys\n"
            "sys.path[:0] = %r\n"
            "import %s\n"
            "print('MODS=ok')\n" % (dirs, ", ".join(PROGRAM_MODULES)))


def probe_modules(py):
    """程序文件齐不齐。返回 (齐?, 说明)。"""
    ok, out = run([py, "-c", module_probe_code()], cwd=ROOT)
    if ok and "MODS=ok" in out:
        return True, "齐（%s）" % "、".join(m + ".py" for m in PROGRAM_MODULES)
    miss = ""
    for ln in out.strip().splitlines():
        ln = ln.strip()
        if "ModuleNotFoundError" in ln or "ImportError" in ln or "SyntaxError" in ln:
            miss = ln
            break
    if not miss:
        miss = next((ln.strip() for ln in out.strip().splitlines() if ln.strip()), "读不出来")
    return False, "缺东西 —— " + miss[:110]


# ── 镜像源（优先本地 wheels，其次清华，最后官方）────────────────────────────
MIRRORS = (
    ("本地 wheels", None),  # None 表示从 WHEELS 目录装
    ("清华镜像", "https://pypi.tuna.tsinghua.edu.cn/simple"),
    ("官方源", None),
)


def runtime_python():
    """runtime 里的解释器路径（Windows 在根目录，其余在 bin 下）。"""
    if os.name == "nt":
        return os.path.join(RUNTIME, "python.exe")
    return os.path.join(RUNTIME, "bin", "python3")


def run(args, **kw):
    """跑一个子进程并把输出接回来。返回 (成功?, 输出文本)。"""
    try:
        p = subprocess.run(args, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", **kw)
    except OSError as e:
        return False, str(e)
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode == 0, out


def probe(py):
    """这个解释器能不能跑起来、装的依赖全不全、它是不是本包的运行时。

    返回 (可用?, 说明)。
    """
    if not py or not os.path.lexists(py):
        return False, "解释器不存在"
    code = (check_imports() + ";import sys;print('PREFIX='+sys.prefix)")
    ok, out = run([py, "-c", code])
    if not ok:
        first = ""
        for ln in out.strip().splitlines():
            if ln.strip() and "PREFIX=" not in ln:
                first = ln.strip()
        return False, ("依赖缺失或无法运行" + ("（" + first[:110] + "）" if first else ""))
    prefix = ""
    for ln in out.splitlines():
        if ln.startswith("PREFIX="):
            prefix = ln[7:].strip()
    if os.path.normcase(os.path.realpath(prefix)) != os.path.normcase(os.path.realpath(RUNTIME)):
        return False, "运行时路径不对（%s）" % (prefix or "?")
    return True, "就绪"


# ── Tk 探测 ────────────────────────────────────────────────────────────────
TK_PROBE = ("import tkinter\n"
            "r = tkinter.Tk()\n"
            "try:\n"
            "    r.attributes('-alpha', 0.0)\n"
            "except Exception:\n"
            "    pass\n"
            "r.withdraw()\n"
            "r.destroy()\n"
            "print('TKPROBE=ok')\n")


def probe_tk(py):
    """这个解释器能不能把 Tk 起起来。返回 (行不行, 说明)。"""
    ok, out = run([py, "-c", TK_PROBE])
    if ok and "TKPROBE=ok" in out:
        return True, "能起 Tk"
    for ln in out.strip().splitlines():
        ln = ln.strip()
        if "init.tcl" in ln or "TclError" in ln or "No module named" in ln:
            return False, ln[:130]
    return False, "Tk 起不来"


# ── 修复运行时（重新识别）───────────────────────────────────────────────────
def fix_runtime():
    """修复运行时的权限和隔离属性。返回 (修了什么, 是否成功)。

    这就是「重新识别」的核心：拷到新位置后，权限和属性可能不对。
    """
    changed = []

    # 1. 修执行权限（FAT/exFAT 拷贝会丢 +x）
    if os.name != "nt":
        bin_dir = os.path.join(RUNTIME, "bin")
        if os.path.isdir(bin_dir):
            for name in os.listdir(bin_dir):
                p = os.path.join(bin_dir, name)
                if os.path.isfile(p) and not os.access(p, os.X_OK):
                    try:
                        os.chmod(p, os.stat(p).st_mode | 0o755)
                        changed.append("+" + name)
                    except OSError:
                        pass
        # Python 主程序
        py = runtime_python()
        if os.path.isfile(py) and not os.access(py, os.X_OK):
            try:
                os.chmod(py, os.stat(py).st_mode | 0o755)
                changed.append("+python3")
            except OSError:
                pass

    # 2. 去 macOS 隔离属性（从网上下载的会带 com.apple.quarantine）
    if sys.platform == "darwin":
        ok, _ = run(["xattr", "-rd", "com.apple.quarantine", RUNTIME])
        if ok:
            changed.append("去隔离")

    # 3. 清 __pycache__（旧字节码可能跟新 Python 不兼容）
    pyc_count = 0
    for dirpath, dirnames, filenames in os.walk(os.path.join(ROOT, "app")):
        if "__pycache__" in dirnames:
            shutil.rmtree(os.path.join(dirpath, "__pycache__"), ignore_errors=True)
            pyc_count += 1
    if pyc_count:
        changed.append("清缓存")

    return changed, True


def install_deps(py):
    """装依赖。先试本地 wheels（离线），再逐个镜像源。

    ⚠️ 装什么由 `need_list()` 按平台决定。
    """
    pkgs = need_list()

    # 先试本地 wheels
    if os.path.isdir(WHEELS):
        whl_files = [f for f in os.listdir(WHEELS) if f.endswith(".whl")]
        if whl_files:
            B.step("正在安装依赖（本地 wheels）：%s" % "、".join(pkgs))
            ok, out = run([py, "-m", "pip", "install", "--disable-pip-version-check",
                           "--no-input", "--no-index",
                           "--find-links", WHEELS] + pkgs)
            if ok:
                B.item(B.OK, "依赖", "安装完成（本地 wheels，无需网络）")
                return True
            # 本地装不全，继续走网络

    # 逐个镜像源试
    for name, url in MIRRORS[1:]:  # 跳过本地 wheels
        args = [py, "-m", "pip", "install", "--disable-pip-version-check",
                "--no-input"] + pkgs
        if url:
            args += ["-i", url]
        B.step("正在安装依赖（%s）：%s" % (name, "、".join(pkgs)))
        B.note("首次约 1～3 分钟，请勿关窗口")
        ok, out = run(args)
        if ok:
            B.item(B.OK, "依赖", "安装完成（%s）" % name)
            return True
        tail = [ln for ln in out.strip().splitlines() if ln.strip()][-6:]
        B.item(B.FAIL, "依赖", "用「%s」没装上，末尾几行：" % name)
        for ln in tail:
            B.note(ln[:150])
        B.note("换一个源再试…")
    return False


def version_code():
    names = [p.replace(">=", "=").split("=")[0] for p in need_list()]
    return ("import importlib.metadata as md\n"
            "for n in %r:\n"
            "    try:\n"
            "        print(f'{n} {md.version(n)}')\n"
            "    except Exception:\n"
            "        print(f'{n} （读不出版本）')" % (names,))


def dep_versions(py):
    ok, out = run([py, "-c", version_code()])
    if not ok:
        return "、".join(p.split(">=")[0] for p in need_list())
    items = []
    for ln in out.strip().splitlines():
        ln = ln.strip()
        if ln:
            items.append(ln.replace("（读不出版本）", "").strip())
    return " · ".join(items) or "（读不出版本）"


def ready_rows(py, tk_ok=True, tk_why="能起 Tk"):
    """环境就绪时的固定四行。返回「程序文件在不在」。"""
    B.item(B.OK, "运行环境", os.path.relpath(RUNTIME, ROOT))
    B.item(B.OK, "依赖", dep_versions(py))
    B.item(B.OK if tk_ok else B.FAIL, "界面库", tk_why)
    mods_ok, mods_why = probe_modules(py)
    B.item(B.OK if mods_ok else B.FAIL, "程序文件", mods_why)
    return mods_ok


def ensure_ready(force=False):
    """确保运行时可用。返回可用解释器路径，失败返回 None。

    「重新识别」的完整流程：探测 → 修复 → 补依赖 → 复核。
    """
    py = runtime_python()

    if not force:
        ok, why = probe(py)
        if ok:
            tk_ok, tk_why = probe_tk(py)
            if tk_ok:
                if ready_rows(py, tk_ok, tk_why):
                    return py
                B.note("这个文件夹可能没拷全 —— 重新拷一份完整的交付包再试。")
                return None
            B.item(B.FAIL, "界面库", tk_why)
            B.note("需要重新识别一次（修权限、补依赖）。")
        elif os.path.lexists(py):
            B.item(B.FAIL, "运行环境", why)
            B.note("需要重新识别一次（修权限、补依赖）。")
        else:
            B.item(B.SKIP, "运行环境", "还没就绪，这一次先把运行时备好")
        B.section("", "重新识别")

    # 修复权限/属性
    B.step("检查运行时权限与属性")
    changed, fix_ok = fix_runtime()
    if changed:
        B.item(B.OK, "修复", "已修：" + "、".join(changed))
    else:
        B.item(B.OK, "修复", "无需修复")

    # 验证运行时存在
    if not os.path.lexists(py):
        B.item(B.FAIL, "运行时", "找不到 runtime 里的 Python")
        B.note("确认这个文件夹是完整的（runtime/ 目录要在）")
        return None

    # 探测一次
    ok, why = probe(py)
    if not ok:
        # 缺依赖，装一次
        B.item(B.SKIP, "运行环境", why)
        if not install_deps(py):
            B.item(B.FAIL, "依赖", "没装上")
            B.note("常见原因：网络不通 —— 但本地 wheels/ 里有离线包，不该走到这里")
            return None

    # 复核
    ok, why = probe(py)
    if not ok:
        B.item(B.FAIL, "可用性", "修完仍然不可用 —— %s" % why)
        return None
    tk_ok, tk_why = probe_tk(py)
    mods_ok = ready_rows(py, tk_ok, tk_why)
    return py if (tk_ok and mods_ok) else None


def _embed(py, args):
    """把 seq_core 的体检输出原样嵌进来。"""
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    ok, out = run([py, ENGINE_MAIN] + args, cwd=ROOT, env=env)
    B.raw(out)
    return ok, out


def report(py, check_one=""):
    """运行前体检：环境 + 能转什么 + 外部程序（+ 可选的试转一份）。"""
    B.section("", "这台电脑能转什么")
    _embed(py, ["--show-rules", "--brief", "--no-head"])

    B.section("", "外部程序")
    _embed(py, ["--show-cad", "--brief", "--no-head"])
    _embed(py, ["--show-bartender", "--no-head"])

    if check_one:
        B.section("", "试转一份（产物写临时目录，原件不动）")
        _embed(py, ["--check-convert", check_one])

    B.end(B.OK, "体检结束：没有转换任何文件，原件一个字节都没动")


def launch(py):
    """用运行时的解释器拉起界面。"""
    if not os.path.exists(APP):
        B.end(B.FAIL, "找不到 seq_app.py",
              steps=["确认这个文件夹是完整的（app/ui/seq_app.py 要在）",
                     "重新拷一份完整的交付包"])
        return 1
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    B.end(B.OK, "环境就绪，正在打开界面…（关掉界面窗口即可退出）")
    try:
        return subprocess.run([py, APP], cwd=HERE, env=env).returncode
    except KeyboardInterrupt:
        return 0


def do_check():
    """`--check`：只看不动，把结论一次说完。"""
    py = runtime_python()
    exist = os.path.isdir(RUNTIME)
    B.section("", "运行环境")
    B.item(B.OK if exist else B.SKIP, "运行时",
           "%s（%s）" % (os.path.relpath(RUNTIME, ROOT), "存在" if exist else "不存在"))
    ok, why = probe(py)
    if not ok:
        B.item(B.FAIL, "可用性", "不可用 —— %s" % why)
        B.end(B.FAIL, "运行时还没准备好",
              steps=["重新拷一份完整的包（runtime/ 要和启动器在同一层）"])
        return 1
    B.item(B.OK, "依赖", dep_versions(py))
    tk_ok, tk_why = probe_tk(py)
    B.item(B.OK if tk_ok else B.FAIL, "界面库", tk_why)
    mods_ok, mods_why = probe_modules(py)
    B.item(B.OK if mods_ok else B.FAIL, "程序文件", mods_why)
    if not mods_ok:
        B.end(B.FAIL, "程序文件不全，先别启动",
              steps=["重新拷一份完整的交付包（少的是 app/ 里的 .py）",
                     "确认没被杀毒软件吃掉文件"])
        return 1
    B.item(B.OK, "可用性", "可用")
    B.end(B.OK, "可以正常启动")
    return 0


def main():
    args = sys.argv[1:]
    B.head(TOOL_NAME, "v" + TOOL_VERSION,
           "运行前体检" if "--report" in args else "",
           [("系统", B.os_line()),
            ("Python", "%d.%d.%d" % sys.version_info[:3]),
            ("位置", ROOT)])

    if "--check" in args:
        return do_check()

    B.section("", "运行环境")
    force = "--rebuild" in args or "--reidentify" in args
    py = ensure_ready(force=force)
    if not py:
        B.end(B.FAIL, "没能准备好环境",
              steps=["把这一屏截图发给我，我按情况处理",
                     "想重来一次就再双击一次启动器"])
        return 1

    if "--report" in args:
        one = ""
        if "--check-one" in args:
            i = args.index("--check-one")
            one = args[i + 1] if i + 1 < len(args) else ""
        report(py, one.strip())
        return 0

    if "--setup" in args or "--reidentify" in args:
        B.end(B.OK, "环境已就绪，以后双击「启动PDF转换拼接工具」就能直接用")
        return 0

    return launch(py)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(0)
