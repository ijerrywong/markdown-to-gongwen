#!/bin/bash
# Markdown → 党政机关公文格式 环境安装脚本
# 在新电脑上运行一次即可，自动创建虚拟环境并安装依赖

set -e
cd "$(dirname "$0")"

echo "📦 创建虚拟环境..."
python3 -m venv .venv

echo "📦 安装依赖..."
.venv/bin/pip install python-docx regex

PROJECT_DIR="$(pwd)"
echo ""
echo "✅ 安装完成！"
echo ""
echo "=== Typora 配置 ==="
echo "  1. 打开 Typora → 偏好设置 → 导出 → 添加导出"
echo "  2. 选择“自定义”，命令填入："
echo ""
echo "    $PROJECT_DIR/typora-gongwen.sh -d ~/Desktop"
echo ""
echo "  3. 保存后即可使用"
echo ""
echo "=== 终端命令行 ==="
echo "  $PROJECT_DIR/.venv/bin/python3 markdown-to-gongwen.py 输入.md 输出.docx"
echo ""
