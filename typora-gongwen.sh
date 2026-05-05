#!/bin/bash
# Typora 导出包装脚本 — 调用 venv 中的 Python 执行公文排版
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/.venv/bin/python3" "$SCRIPT_DIR/markdown-to-gongwen.py" "$@"
