#!/bin/bash
# Typora 导出钩子 — 调用 venv 中的 Python 执行公文排版
# Typora 配置：将"命令"设为此脚本的完整路径即可

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
"$SCRIPT_DIR/venv/bin/python3" "$SCRIPT_DIR/typora-gongwen.py" "$@"
