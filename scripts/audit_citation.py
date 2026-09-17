#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引用依据占比审计 —— 量化"有多少内容有据可依"。

用途：
    1) 应对评审或内部合规审查对"是否套用他项目标书"的质疑；
    2) 反向自查：自主撰写比例过高的章节，往往是过度设计的高发区。

统计口径（可复核）：
    直接引用  —— 与任一权威来源存在 ≥N 字（默认15）连续相同文字的段落
    有据转述  —— 无连续相同文字，但含显式引用标记（"本章节满足""依据技术规格书"…）
    自主撰写  —— 其余

用法：
    python3 audit_citation.py --draft "0*.md" \
        --spec 技术规格书.docx --tender 招标文件.docx \
        --ref 参考A.docx 参考B.docx --sub 子系统资料/
"""
import argparse
import glob
import os
import re
import sys
from difflib import SequenceMatcher

from check_invented import read_any, collect_drafts

CITE = re.compile(
    r"本章节满足|依据技术规格书|依据招标文件|依据子系统资料|依据参考|"
    r"按技术规格书|招标文件要求|技术规格书要求|投标人须知|招标公告|"
    r"规格书明确|招标文件明确|子系统资料")


def norm(t):
    t = re.sub(r"\[TABLE\]|\[/TABLE\]", "", t)
    return re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", t)


def cn(s):
    return len(re.findall(r"[\u4e00-\u9fff]", s))


def maxmatch(seg, src, minlen):
    if not seg or not src:
        return 0
    sm = SequenceMatcher(None, seg, src, autojunk=False)
    m = sm.find_longest_match(0, len(seg), 0, len(src))
    return m.size if m.size >= minlen else 0


def main():
    ap = argparse.ArgumentParser(description="引用依据占比审计")
    ap.add_argument("--draft", nargs="+", required=True)
    ap.add_argument("--spec", nargs="*", default=[])
    ap.add_argument("--tender", nargs="*", default=[])
    ap.add_argument("--ref", nargs="*", default=[], help="参考素材（往期标书等）")
    ap.add_argument("--sub", nargs="*", default=[], help="子系统/产品资料")
    ap.add_argument("--min-len", type=int, default=15)
    args = ap.parse_args()

    groups = {
        "技术规格书": args.spec,
        "招标文件": args.tender,
        "参考素材": args.ref,
        "子系统/产品资料": args.sub,
    }
    SRC = {}
    for name, paths in groups.items():
        if not paths:
            continue
        buf = []
        for p in paths:
            if os.path.isdir(p):
                for root, _, files in os.walk(p):
                    for fn in files:
                        buf.append(read_any(os.path.join(root, fn)))
            else:
                buf.append(read_any(p))
        SRC[name] = norm("".join(buf))
    if not SRC:
        ap.error("至少提供一类来源（--spec/--tender/--ref/--sub）")

    files = collect_drafts(args.draft)
    print(f"来源：{'、'.join(f'{k}({len(v)}字)' for k, v in SRC.items())}\n")

    hdr = f"{'章节':<40}{'直接引用':>9}{'有据转述':>9}{'自主撰写':>9}{'有据率':>9}"
    print(hdr); print("-" * 74)
    total = {"直接引用": 0, "有据转述": 0, "自主撰写": 0}
    detail = {}
    for f in files:
        txt = open(f, encoding="utf-8").read()
        paras = [p.strip() for p in txt.split("\n")
                 if p.strip() and not p.strip().startswith("#")]
        c = dict.fromkeys(total, 0)
        for p in paras:
            pn = norm(p)
            if len(pn) < 8:
                continue
            best, which = 0, ""
            for name, src in SRC.items():
                k = maxmatch(pn, src, args.min_len)
                if k > best:
                    best, which = k, name
            n = cn(p)
            if best:
                c["直接引用"] += n
                detail[which] = detail.get(which, 0) + n
            elif CITE.search(p):
                c["有据转述"] += n
            else:
                c["自主撰写"] += n
        for k in c:
            total[k] += c[k]
        s = sum(c.values()) or 1
        print(f"{os.path.basename(f)[:38]:<40}{c['直接引用']:>9}{c['有据转述']:>9}"
              f"{c['自主撰写']:>9}{(c['直接引用']+c['有据转述'])/s*100:>8.1f}%")
    s = sum(total.values()) or 1
    print("-" * 74)
    print(f"{'合计':<40}{total['直接引用']:>9}{total['有据转述']:>9}{total['自主撰写']:>9}"
          f"{(total['直接引用']+total['有据转述'])/s*100:>8.1f}%")
    print(f"\n总计 {s} 字")
    print("\n【直接引用来源分布】")
    for k, v in sorted(detail.items(), key=lambda x: -x[1]):
        print(f"   {k:<18}{v:>8} 字   占全文 {v/s*100:.2f}%")
    print("\n说明：'自主撰写'比例显著偏高的章节，应重点复查是否存在招标未要求的"
          "自创内容（可配合 check_invented.py 交叉验证）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
