#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成后校验：源 markdown ↔ 产出 docx 双向核对，并对差异做完整归因。

为什么必须做这一步：
    长文档生成最危险的失败模式不是报错，而是"静默丢内容"——文档能打开、
    页数看着正常，但尾部或某些章节悄悄少了内容。只比较总字数不足以发现
    问题（可能一处丢失、另一处重复而总数相近）。本脚本做三重核对：
      1) 段落级存在性：源文件每一段是否都能在产物中找到
      2) 表格级计数：表格数 / 单元格数 / 单元格中文字数（含表头，XML 直数）
      3) 差异归因：把总字数差拆解为可解释项，残留"无法解释的差异"必须为 0

用法：
    python3 verify_docx.py --config doc_config.json
    python3 verify_docx.py --docx 输出.docx --src 01-xxx.md 02-xxx.md ...
    python3 verify_docx.py --auto            # 自动配对 NN-*.md 与同名 docx

退出码：0=通过；1=存在无法解释的差异或段落缺失（应中止交付并排查）
"""
import argparse
import glob
import json
import re
import sys

from docx import Document
from docx.oxml.ns import qn

MD_LINK = re.compile(r"\[([^\]]*)\]\([^)]*\)")
BOLD = re.compile(r"\*\*(.+?)\*\*")
DROP = [
    re.compile(r"^（第[一二三四五六七八九十]+章.*完.*）$"),
    re.compile(r"^（第[一二三四五六七八九十]+章·.*）$"),
    re.compile(r"^【本节完】$"),
]


def clean(t: str) -> str:
    t = MD_LINK.sub(r"\1", t)
    t = BOLD.sub(r"\1", t)
    return t.replace("**", "").replace("`", "").strip()


def cn(s: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", s))


def norm(s: str) -> str:
    return re.sub(r"\s", "", s)


def iter_src(path):
    """逐行产出 ('h'|'p'|'t', ...)，跳过表格数据行（表格另行统计）。"""
    lines = open(path, encoding="utf-8").read().split("\n")
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if any(d.match(s) for d in DROP):
            i += 1
            continue
        if s.startswith("|") and i + 1 < len(lines) \
                and re.match(r"^\|[\s:\-|]+\|$", lines[i + 1].strip()):
            hdr = [clean(c) for c in s.strip("|").split("|")]
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([clean(c) for c in lines[i].strip().strip("|").split("|")])
                i += 1
            yield ("t", hdr, rows)
            continue
        if re.match(r"^#{1,6}\s", s):
            t = clean(re.sub(r"^#{1,6}\s+", "", s))
            t = re.sub(r"^第[一二三四五六七八九十]+章\s*", "", t)
            yield ("h", t)
        else:
            yield ("p", clean(s))
        i += 1


def src_stats(paths):
    paras, nt, ncell, ccell = [], 0, 0, 0
    for p in paths:
        for item in iter_src(p):
            if item[0] == "t":
                nt += 1
                for c in item[1]:
                    ncell += 1
                    ccell += cn(c)
                for row in item[2]:
                    for c in row:
                        ncell += 1
                        ccell += cn(c)
            else:
                paras.append(item[1])
    return paras, nt, ncell, ccell


def docx_stats(path):
    d = Document(path)
    docx_paras = [p.text for p in d.paragraphs if p.text.strip()]
    nt = len(d.tables)
    ncell = ccell = 0
    for t in d.tables:
        for tc in t._tbl.iter(qn("w:tc")):          # XML 直数：合并单元格只算一次
            ncell += 1
            ccell += cn("".join(n.text or "" for n in tc.iter(qn("w:t"))))
    all_cn = cn("\n".join(docx_paras)) + ccell
    return docx_paras, nt, ncell, ccell, all_cn


def autodetect_title(docx_path):
    """从产物中自动读取封面标题（首个 Heading 1 段落），避免漏传 --title 造成误报。"""
    try:
        d = Document(docx_path)
        for p in d.paragraphs:
            if p.style and p.style.name == "Heading 1" and p.text.strip():
                return p.text
    except Exception:
        pass
    return None


def verify(docx_path, src_paths, title=None, verbose=True):
    sp, snt, snc, scc = src_stats(src_paths)
    dp, dnt, dnc, dcc, dall = docx_stats(docx_path)
    dnorm = norm("\n".join(dp))
    if title is None:
        title = autodetect_title(docx_path)
        if title and verbose:
            print(f"  （自动识别封面标题：{title[:40]}…）")

    # 1) 段落存在性
    missing = []
    for t in sp:
        k = norm(t)
        if len(k) >= 6 and k[:40] not in dnorm:
            missing.append(t[:70])

    # 2) 差异归因
    src_raw = sum(cn(open(p, encoding="utf-8").read()) for p in src_paths)
    dropped = 0
    for p in src_paths:
        for ln in open(p, encoding="utf-8").read().split("\n"):
            s = ln.strip()
            if any(d.match(s) for d in DROP):
                dropped += cn(s)
    stripped_h1 = 0
    for p in src_paths:
        m = re.search(r"^#\s+(第[一二三四五六七八九十]+章)", open(p, encoding="utf-8").read(), re.M)
        if m:
            stripped_h1 += cn(m.group(1))
    added_title = cn(title.replace("\n", "")) if title else 0
    expected = src_raw - dropped - stripped_h1 + added_title
    unexplained = dall - expected

    if verbose:
        print(f"\n{'='*70}\n核对：{docx_path}\n{'='*70}")
        print(f"  源文件数 {len(src_paths)}   源中文 {src_raw} 字")
        print(f"\n  【1】段落存在性")
        print(f"      源段落(非表格) {len(sp)} 段   产物段落 {len(dp)} 段")
        print(f"      未在产物中找到：{len(missing)} 处   {'✅' if not missing else '❌'}")
        for m in missing[:8]:
            print(f"        ⚠ {m}")
        print(f"\n  【2】表格核对（含表头，XML 直数）")
        ok_t = (snt == dnt and snc == dnc and scc == dcc)
        print(f"      表格 {snt} -> {dnt}   单元格 {snc} -> {dnc}   单元格中文 {scc} -> {dcc}   {'✅' if ok_t else '❌'}")
        print(f"\n  【3】总字数差异归因")
        print(f"      源文件中文                {src_raw}")
        print(f"      - 剥离章末工作标记         -{dropped}")
        print(f"      - 剥离'第X章'前缀(转自动编号) -{stripped_h1}")
        print(f"      + 新增封面标题             +{added_title}")
        print(f"      = 预期产物字数             {expected}")
        print(f"      实际产物字数               {dall}")
        print(f"      ★ 无法解释的差异           {unexplained}   {'✅' if unexplained == 0 else '❌ 需排查'}")

    ok = (not missing) and (snt == dnt) and (snc == dnc) and (scc == dcc) and (unexplained == 0)
    return ok


def main():
    ap = argparse.ArgumentParser(description="源 markdown ↔ 产出 docx 双向核对")
    ap.add_argument("--config")
    ap.add_argument("--docx")
    ap.add_argument("--src", nargs="*")
    ap.add_argument("--title", default=None)
    args = ap.parse_args()

    results = []
    if args.config:
        cfg = json.load(open(args.config, encoding="utf-8"))
        for d in cfg.get("documents", []):
            files = [c["file"] for c in d["chapters"]]
            results.append(verify(d["output"], files, d.get("title")))
    elif args.docx and args.src:
        results.append(verify(args.docx, args.src, args.title))
    else:
        ap.error("需要 --config，或 --docx 与 --src")

    print("\n" + "=" * 70)
    if all(results):
        print("✅ 全部核对通过：无内容丢失")
        return 0
    print("❌ 核对未通过：请排查上述标 ❌ 的项目，不要直接交付")
    return 1


if __name__ == "__main__":
    sys.exit(main())
