#!/usr/bin/env python3
"""
Typora 自定义导出插件 — Markdown → 党政机关公文格式 (.docx)

macOS 配置方法（配置一次后，只需点击"导出"即可）：
  1. 打开 Typora → 偏好设置 → 导出 → 添加导出
  2. 类型选"自定义"，在"命令"字段填入完整命令：
     /Users/jerry/Projects/markdown-to-gongwen/typora-gongwen.sh -d ~/Desktop
  3. 保存后，文件 → 导出 → 公文导出 即可使用

  - 输出文件名自动与 .md 源文件同名（扩展名为 .docx）
  - -d 指定输出目录（不加则输出到 .md 文件所在目录）
  - 导出前请先在 Typora 中保存好文档
"""

import sys
import os
import importlib.util

# ── 导入核心排版模块 ──────────────────────────────────
# 原文件 markdown-to-gongwen.py 含连字符，不能用 import 语句直接导入，
# 因此用 importlib 动态加载。
_script_dir = os.path.dirname(os.path.abspath(__file__))
_original_path = os.path.join(_script_dir, "markdown-to-gongwen.py")
if not os.path.exists(_original_path):
    print("❌ 找不到核心排版文件 markdown-to-gongwen.py", file=sys.stderr)
    print(f"   预期路径: {_original_path}", file=sys.stderr)
    sys.exit(1)

_spec = importlib.util.spec_from_file_location("_gongwen_core", _original_path)
_gongwen = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_gongwen)
convert_markdown_to_gongwen = _gongwen.convert_markdown_to_gongwen


def _detect_typora_document() -> str | None:
    """macOS: 探测 Typora 当前文档路径（无需 Accessibility 权限）。"""

    import subprocess

    def _via_axdocument() -> str | None:
        """方案A：通过 AXDocument 属性（需 Accessibility 权限）。"""
        try:
            cmd = [
                "osascript", "-e",
                'tell application "System Events" to tell process "Typora" '
                'to get value of attribute "AXDocument" of front window'
            ]
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
            if r.returncode == 0 and r.stdout.strip():
                url = r.stdout.strip()
                if url.startswith("file://"):
                    from urllib.parse import unquote, urlparse
                    path = unquote(urlparse(url).path)
                    if os.path.exists(path):
                        return path
        except Exception:
            pass
        return None

    def _via_window_title() -> str | None:
        """方案B：通过窗口标题搜索文件（无需额外权限）。"""
        try:
            cmd = [
                "osascript", "-e",
                'tell application "Typora" to get name of window 1'
            ]
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
            if r.returncode != 0 or not r.stdout.strip():
                return None
            win_title = r.stdout.strip()
            if not win_title or win_title.startswith("未命名"):
                return None

            # 构造候选文件名
            base = win_title[:-3] if win_title.endswith(".md") else win_title
            fname = base + ".md"

            # 常见目录搜索
            search_dirs = [
                os.path.expanduser("~/Documents"),
                os.path.expanduser("~/Desktop"),
                os.getcwd(),
            ]
            for d in search_dirs:
                full = os.path.join(d, fname)
                if os.path.exists(full) and os.path.isfile(full):
                    return os.path.abspath(full)

            # Spotlight 兜底
            sp = subprocess.run(
                ["mdfind", "-onlyin", os.path.expanduser("~"),
                 f"kMDItemFSName == '{fname}'"],
                capture_output=True, text=True, timeout=10
            )
            for line in sp.stdout.splitlines():
                p = line.strip()
                if os.path.exists(p) and p.endswith(".md"):
                    return p
        except Exception:
            pass
        return None

    return _via_axdocument() or _via_window_title()


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Markdown 转党政机关公文格式（Typora 自定义导出用）"
    )
    parser.add_argument(
        "input", nargs="?",
        help="Markdown 文件路径（不传则自动探测 Typora 当前文档）"
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="完整输出文件路径（优先级高于 -d）"
    )
    parser.add_argument(
        "-d", "--dir",
        default=None,
        help="输出目录（默认与 .md 源文件同目录，配合 -d 可指定如 ~/Desktop）"
    )
    args = parser.parse_args()

    md = None
    input_filepath = None  # 记录输入文件路径，用于生成默认输出名

    # ── 模式1：传了文件路径 ──
    if args.input:
        input_filepath = args.input.strip("\"'")
        if not os.path.exists(input_filepath):
            print(f"❌ 找不到文件: {input_filepath}", file=sys.stderr)
            sys.exit(1)
        with open(input_filepath, "r", encoding="utf-8") as f:
            md = f.read()

    # ── 模式2：尝试探测 Typora 当前文档 ──
    if md is None:
        input_filepath = _detect_typora_document()
        if input_filepath:
            print(f"📄 检测到 Typora 文档: {input_filepath}", file=sys.stderr)
            with open(input_filepath, "r", encoding="utf-8") as f:
                md = f.read()

    # ── 模式3：从 stdin 读取 ──
    if md is None:
        md = sys.stdin.read()

    if not md or not md.strip():
        print("❌ 错误：未读取到 Markdown 内容", file=sys.stderr)
        print("   请确保已在 Typora 中保存文件，再点击导出。", file=sys.stderr)
        sys.exit(1)

    # ── 确定输出路径 ──
    if args.output:
        output_path = os.path.expanduser(args.output)
    elif args.dir or input_filepath:
        if args.dir:
            out_dir = os.path.expanduser(args.dir)
        else:
            out_dir = os.path.dirname(input_filepath) if input_filepath else "."
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(
            os.path.basename(input_filepath) if input_filepath else "公文输出"
        )[0]
        output_path = os.path.join(out_dir, base + ".docx")
    else:
        output_path = "公文输出.docx"

    result_path = convert_markdown_to_gongwen(md, output_path)
    print(f"✅ 公文已生成: {result_path}")


if __name__ == "__main__":
    main()
