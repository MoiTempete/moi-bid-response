#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合并多个 docx 为一份（SKILL.md 长期引用但此前缺失的脚本）。

用途：
    按章分批生成 docx 后合并为完整文档。合并采用"把后续文档的 body 元素
    依序搬入首个文档"的方式，保留各自的样式与编号定义（以第一份为准）。

用法：
    python3 merge_docx.py -o 完整文档.docx 01.docx 02.docx 03.docx
    python3 merge_docx.py -o out.docx --glob "part-*.docx" --page-break
"""
import argparse
import glob
import copy
import sys

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn


def merge(output, inputs, page_break=True):
    """把后续文档的 body 元素依序搬入首个文档。

    要点：sectPr（节属性）必须始终留在 body 末尾，所有内容都插到它之前；
    python-docx 的 add_paragraph 依赖 'w:sectPr' 作为 successors，会自动
    插到 sectPr 之前，因此分页符也可安全使用。
    """
    if not inputs:
        raise SystemExit("没有输入文件")
    base = Document(inputs[0])
    body = base.element.body
    sect = body.find(qn("w:sectPr"))

    def append_element(el):
        el = copy.deepcopy(el)
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)

    for path in inputs[1:]:
        d = Document(path)
        src_body = d.element.body
        src_sect = src_body.find(qn("w:sectPr"))
        children = [c for c in list(src_body) if c is not src_sect]
        if not children:
            continue
        if page_break:
            p = base.add_paragraph()          # 自动插到 sectPr 之前
            p.add_run().add_break(WD_BREAK.PAGE)
        for child in children:
            append_element(child)

    base.save(output)
    return output


def main():
    ap = argparse.ArgumentParser(description="合并多个 docx")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("inputs", nargs="*")
    ap.add_argument("--glob", help="用通配符指定输入（按文件名排序）")
    ap.add_argument("--no-page-break", action="store_true")
    args = ap.parse_args()

    files = list(args.inputs)
    if args.glob:
        files = sorted(glob.glob(args.glob))
    if not files:
        ap.error("需要输入文件或 --glob")

    out = merge(args.output, files, page_break=not args.no_page_break)
    print(f"已合并 {len(files)} 个文档 -> {out}")
    print("⚠ 合并后请运行 verify_docx.py 或逐份核对，确认合并未丢内容")


if __name__ == "__main__":
    sys.exit(main())
