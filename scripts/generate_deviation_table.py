#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
技术响应偏离表生成器 —— 从技术规格书自动生成偏离表骨架。

背景（为什么需要）：
    技术响应偏离表几乎是每份招标的必备交付物，且招标文件通常规定"未在偏离表
    中列明的偏离项，视为完全响应"。但很多响应文件只在正文里写"完全响应"，
    没有独立的偏离表；或者偏离表覆盖不全，导致漏报偏离而承担兜底责任。
    本脚本把规格书条款结构抽出来，预填"条款位置/内容摘要"两列，编写者只需
    补"响应内容/响应状态"两列，确保覆盖无遗漏。

用法：
    # 生成骨架（5列制式，兼容常见招标格式）
    python3 generate_deviation_table.py 技术规格书.docx -o 偏离表骨架.md

    # 仅抽"强制性条款"（应/须/必须/不得/严禁），得到精简版
    python3 generate_deviation_table.py 技术规格书.docx --mandatory-only

    # 自定义表头（如招标规定了列名）
    python3 generate_deviation_table.py 技术规格书.docx \
        --headers "序号,招标文件条目号,招标文件的技术条款,投标文件的技术条款,响应状态"
"""
import argparse
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from parse_tender import parse_tender  # noqa: E402

MANDATORY = re.compile(r"应|须|必须|不得|严禁|不允许|禁止|至少|不低于|不超过")

# 结构性标题（本身不构成技术要求，作为"条款位置"归组用）
SKIP_TITLE = re.compile(r"^(目\s*录|编\s*制|审\s*核|审\s*批|第[一二三四五六七八九十]+章)")


def iter_clauses(docx_path, mandatory_only=False, max_len=150):
    """产出 (条款位置, 内容摘要)。条款位置取最近的上层标题路径。"""
    r = parse_tender(docx_path)
    path = {}
    for node in r.get("flat_headings", []):
        text = re.sub(r"\s", "", node.get("text", ""))
        lvl = node.get("level") or 0
        if not text or SKIP_TITLE.match(text) or len(text) < 4:
            continue
        # 纯标题（短且无标点、无强制措辞）作为路径节点
        is_title = len(text) <= 30 and not re.search(r"[。；：]", text)
        if is_title:
            path[lvl] = text
            for k in [k for k in path if k > lvl]:
                path.pop(k)
            continue
        loc = " - ".join(path[k] for k in sorted(path))
        if mandatory_only and not MANDATORY.search(text):
            continue
        yield loc, (text[:max_len] + ("…" if len(text) > max_len else ""))


def main():
    ap = argparse.ArgumentParser(description="技术响应偏离表骨架生成")
    ap.add_argument("spec", help="技术规格书（.docx）")
    ap.add_argument("-o", "--output", default="技术响应偏离表-骨架.md")
    ap.add_argument("--mandatory-only", action="store_true",
                    help="仅输出含强制性措辞的条款")
    ap.add_argument("--headers", default="序号,条款位置,条款内容摘要,投标文件的技术条款,响应状态")
    ap.add_argument("--max-len", type=int, default=150)
    args = ap.parse_args()

    headers = [h.strip() for h in args.headers.split(",")]
    rows = list(iter_clauses(args.spec, args.mandatory_only, args.max_len))
    if not rows:
        print("未抽取到条款，请检查规格书是否为 .docx 且含结构化标题。")
        return 1

    out = ["# 技术响应偏离表", "",
           "填写说明：投标人所有偏离项必须在此表中列出；本表未列明的偏离项，"
           "招标人将视为投标人完全响应招标要求，投标人应按不低于招标时的技术要求执行，"
           "并承担由此产生的一切责任和损失。", "",
           "| " + " | ".join(headers) + " |",
           "|" + "---|" * len(headers)]
    for i, (loc, txt) in enumerate(rows, 1):
        cells = []
        for h in headers:
            if "序号" in h:
                cells.append(str(i))
            elif "位置" in h or "条目" in h:
                cells.append(loc)
            elif "招标" in h and ("条款" in h or "内容" in h or "摘要" in h):
                cells.append(txt)
            else:
                cells.append("【待填写】")
        out.append("| " + " | ".join(cells) + " |")
    out += ["", "填写说明：", "",
            "投标人所有偏离项必须在此表中列出，投标人未在本偏离表中明确投标偏离内容，"
            "招标人将视为投标人完全响应招标要求。", "",
            "投标人公章：", "", "投标人代表签字：　　　　　　　　　　日期："]

    open(args.output, "w", encoding="utf-8").write("\n".join(out))
    print(f"已生成 {args.output}，共 {len(rows)} 条条款")
    print("\n后续要求：")
    print("  1) 逐条填写'投标文件的技术条款'列，写明响应位置与内容提要；")
    print("  2) '响应状态'列区分：完全响应 / 正偏离（增值） / 负偏离（须说明理由）；")
    print("  3) 凡招标未要求而自行增设的内容，一律标为正偏离，避免被误认为负偏离；")
    print("  4) 存在范围边界澄清（如某子系统仅监不控）的，须在表后'填写说明'中逐条说明。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
