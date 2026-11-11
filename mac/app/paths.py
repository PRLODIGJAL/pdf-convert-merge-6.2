# -*- coding: utf-8 -*-
"""paths.py —— 全项目唯一的路径来源（四层目录分层，模块名保持扁平）。

参照 PDF 差异对比工具 6.1 的目录规范：
    app/ui/       界面层    只画界面、收事件
    app/judge/    判定层    业务规则（本项目体量小暂未启用，保留占位）
    app/engine/   引擎层    可复用底层能力
    app/shell/    启动壳    环境自检、装依赖、启动

⚠️ 模块名保持扁平（import seq_core）。四层只是目录分层，不是 Python 包分层。
   install() 把四层一起塞进 sys.path，扁平名字照旧能解析。
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP = os.path.join(ROOT, "app")
UI = os.path.join(APP, "ui")
JUDGE = os.path.join(APP, "judge")
ENGINE = os.path.join(APP, "engine")
SHELL = os.path.join(APP, "shell")
LAYERS = (UI, JUDGE, ENGINE, SHELL)
DATA = os.path.join(ROOT, "data")
DEV = os.path.join(ROOT, "dev")
VENV = os.path.join(ROOT, ".venv")
CONFIG_PATH = os.path.join(DATA, "上次选择.json")


def install():
    """把四层加进 sys.path（幂等），返回项目根目录。"""
    for p in reversed(LAYERS):
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)
    return ROOT
