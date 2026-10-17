#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown → Word 转换工具
========================
用法:
  python convert_to_docx.py                     # 扫描当前目录，交互式选择 MD 文件
  python convert_to_docx.py 报告.md              # 转换指定文件
  python convert_to_docx.py --dir ./某目录        # 转换指定目录下所有 MD 文件
  python convert_to_docx.py --all                # 转换当前目录及子目录下所有 MD 文件

也可将 .md 文件直接拖拽到本脚本上运行。
"""

import re
import sys
import os
import argparse
from pathlib import Path
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
except ImportError:
    print("错误: 缺少 python-docx 库，请运行: pip install python-docx")
    sys.exit(1)

# ── 样式配置 ──────────────────────────────────────
TITLE_COLOR   = RGBColor(0, 51, 102)    # 深蓝
HEADING_COLOR = RGBColor(0, 51, 102)
H3_COLOR      = RGBColor(51, 51, 51)
QUOTE_COLOR   = RGBColor(80, 80, 80)
SEP_COLOR     = RGBColor(200, 200, 200)
TABLE_HEAD_BG = '003366'
FONT_NAME     = '微软雅黑'

# ── 核心转换逻辑 ──────────────────────────────────

def add_formatted_run(paragraph, text, size=Pt(10)):
    """向段落添加文本，支持 **加粗** 标记"""
    remaining = text
    while remaining:
        bold_match = re.search(r'\*\*(.+?)\*\*', remaining)
        if bold_match:
            if bold_match.start() > 0:
                run = paragraph.add_run(remaining[:bold_match.start()])
                run.font.size = size
                run.font.name = FONT_NAME
            run = paragraph.add_run(bold_match.group(1))
            run.bold = True
            run.font.size = size
            run.font.name = FONT_NAME
            remaining = remaining[bold_match.end():]
        else:
            run = paragraph.add_run(remaining)
            run.font.size = size
            run.font.name = FONT_NAME
            remaining = ''


def build_table(doc, headers, data):
    """在文档中创建格式化表格"""
    table = doc.add_table(rows=len(data) + 1, cols=len(headers),
                          style='Light Grid Accent 1')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = ''
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(9)
        run.font.name = FONT_NAME
        run.font.color.rgb = RGBColor(255, 255, 255)
        shading = cell._element.get_or_add_tcPr()
        shd = shading.makeelement(qn('w:shd'), {
            qn('w:fill'): TABLE_HEAD_BG, qn('w:val'): 'clear'
        })
        shading.append(shd)

    for ri, row_data in enumerate(data):
        for ci, cell_text in enumerate(row_data):
            if ci < len(headers):
                cell = table.cell(ri + 1, ci)
                cell.text = ''
                run = cell.paragraphs[0].add_run(cell_text)
                run.font.size = Pt(9)
                run.font.name = FONT_NAME

    doc.add_paragraph()


def convert_md_to_docx(md_path, docx_path=None):
    """将单个 Markdown 文件转换为 Word 文档"""
    if docx_path is None:
        docx_path = md_path.with_suffix('.docx')

    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin   = Cm(3.18)
        section.right_margin  = Cm(3.18)

    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.5

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()

        if not line:
            i += 1
            continue

        # ── 一级标题 ──
        if line.startswith('# ') and not line.startswith('## '):
            p = doc.add_paragraph()
            run = p.add_run(line[2:])
            run.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = TITLE_COLOR
            run.font.name = FONT_NAME
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after = Pt(12)

        # ── 二级标题 ──
        elif line.startswith('## '):
            p = doc.add_paragraph()
            run = p.add_run(line[3:])
            run.bold = True
            run.font.size = Pt(15)
            run.font.color.rgb = HEADING_COLOR
            run.font.name = FONT_NAME
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(8)

        # ── 三级标题 ──
        elif line.startswith('### '):
            p = doc.add_paragraph()
            run = p.add_run(line[4:])
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = H3_COLOR
            run.font.name = FONT_NAME
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)

        # ── 引用块 ──
        elif line.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            add_formatted_run(p, line[2:], size=Pt(9.5))
            for r in p.runs:
                r.font.color.rgb = QUOTE_COLOR
                r.italic = True

        # ── 分隔线 ──
        elif line.strip() == '---':
            p = doc.add_paragraph()
            run = p.add_run('—' * 50)
            run.font.color.rgb = SEP_COLOR
            run.font.size = Pt(8)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # ── 表格 ──
        elif line.startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            i -= 1

            rows = []
            for tl in table_lines:
                cells = [c.strip() for c in tl.split('|')[1:-1]]
                rows.append(cells)

            data_rows = [r for r in rows
                         if not all(re.match(r'^-+$', c) for c in r)]

            if data_rows:
                headers = data_rows[0]
                data    = data_rows[1:] if len(data_rows) > 1 else []
                build_table(doc, headers, data)

        # ── 代码块（跳过） ──
        elif line.startswith('```'):
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                i += 1
            i += 1
            continue

        # ── 普通段落 ──
        else:
            p = doc.add_paragraph()
            if line.startswith('- ') or line.startswith('  - '):
                p.paragraph_format.left_indent = \
                    Cm(1.5) if line.startswith('  - ') else Cm(1)
                text = line[2:].strip() if line.startswith('  - ') else line[2:]
            elif re.match(r'^\d+\.\s', line):
                p.paragraph_format.left_indent = Cm(1)
                text = line
            elif re.match(r'^\s{2,4}', line):
                p.paragraph_format.left_indent = Cm(2)
                text = line.strip()
            else:
                text = line
            add_formatted_run(p, text)

        i += 1

    doc.save(str(docx_path))
    return docx_path


# ── 文件扫描 & 交互 ──────────────────────────────

def find_md_files(root_dir, recursive=False):
    """扫描目录中的 MD 文件（排除已有 DOCX 对应的临时文件）"""
    root = Path(root_dir)
    if recursive:
        md_files = sorted(root.rglob('*.md'))
    else:
        md_files = sorted(root.glob('*.md'))
    # 只保留有对应内容的文件，过滤空文件
    return [f for f in md_files if f.stat().st_size > 100]


def interactive_select(md_files):
    """交互式选择要转换的文件"""
    print("\n" + "=" * 60)
    print("  Markdown → Word 转换工具")
    print("=" * 60)
    print(f"\n  扫描到 {len(md_files)} 个 Markdown 文件:\n")

    for idx, f in enumerate(md_files, 1):
        size_kb = f.stat().st_size / 1024
        rel_path = f.relative_to(Path.cwd()) if f.is_relative_to(Path.cwd()) else f
        print(f"  [{idx}] {rel_path}  ({size_kb:.0f} KB)")

    print(f"  [A] 全部转换（共 {len(md_files)} 个）")
    print(f"  [Q] 退出")
    print()

    while True:
        try:
            choice = input("  请选择 (数字/A/Q): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  已取消。")
            return []

        if choice.upper() == 'Q':
            return []
        if choice.upper() == 'A':
            return md_files
        try:
            idx = int(choice)
            if 1 <= idx <= len(md_files):
                return [md_files[idx - 1]]
        except ValueError:
            pass
        print(f"  无效输入，请输入 1-{len(md_files)}、A 或 Q")


def convert_batch(md_files):
    """批量转换并输出进度"""
    success, fail = 0, 0
    for f in md_files:
        docx_path = f.with_suffix('.docx')
        try:
            result = convert_md_to_docx(f, docx_path)
            rel = result.relative_to(Path.cwd()) if result.is_relative_to(Path.cwd()) else result
            print(f"  ✓ {rel}")
            success += 1
        except PermissionError:
            alt = f.parent / f"{f.stem}_new.docx"
            try:
                result = convert_md_to_docx(f, alt)
                print(f"  ✓ {alt.name} (原文件被锁定，已另存)")
                success += 1
            except Exception as e:
                print(f"  ✗ {f.name}: {e}")
                fail += 1
        except Exception as e:
            print(f"  ✗ {f.name}: {e}")
            fail += 1

    return success, fail


# ── 主入口 ────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='将 Markdown 分析报告转换为格式优美的 Word 文档')
    parser.add_argument('files', nargs='*',
                        help='要转换的 .md 文件（可拖拽）')
    parser.add_argument('--dir', '-d', type=str, default=None,
                        help='扫描并转换指定目录下的所有 .md 文件')
    parser.add_argument('--all', '-a', action='store_true',
                        help='转换当前目录及子目录下所有 .md 文件')
    parser.add_argument('--recursive', '-r', action='store_true',
                        help='递归扫描子目录')

    args = parser.parse_args()

    # 修复 Windows 控制台编码
    if sys.platform == 'win32':
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass

    md_files = []

    # 情况1: 指定了具体文件（命令行参数或拖拽）
    if args.files:
        for f in args.files:
            p = Path(f).resolve()
            if p.exists() and p.suffix.lower() == '.md':
                md_files.append(p)
            else:
                print(f"  跳过: {f} (不存在或非 .md 文件)")

    # 情况2: --dir 指定目录
    elif args.dir:
        md_files = find_md_files(args.dir, recursive=args.recursive)

    # 情况3: --all 当前目录递归
    elif args.all:
        md_files = find_md_files(Path.cwd(), recursive=True)

    # 情况4: 无参数 → 扫描当前目录 + 子目录，交互选择
    else:
        md_files = find_md_files(Path.cwd(), recursive=True)

    if not md_files:
        print("  未找到任何 .md 文件。")
        print("  用法: python convert_to_docx.py [文件.md ...]")
        print("        python convert_to_docx.py --dir ./某目录")
        return

    # 单个文件直接转换，多个文件交互选择
    if len(md_files) == 1 and not args.files:
        selected = [md_files[0]]
        print(f"\n  发现 1 个文件，直接转换...")
    elif args.files or args.dir or args.all:
        # 命令行模式下直接全部转换
        selected = md_files
    else:
        selected = interactive_select(md_files)

    if not selected:
        return

    print(f"\n  开始转换 ({len(selected)} 个文件)...\n")
    success, fail = convert_batch(selected)

    print(f"\n  完成: {success} 成功, {fail} 失败")
    print("=" * 60)

    # Windows 下暂停以查看结果
    if sys.platform == 'win32':
        try:
            input("\n  按 Enter 关闭...")
        except (EOFError, KeyboardInterrupt):
            pass


if __name__ == '__main__':
    main()
