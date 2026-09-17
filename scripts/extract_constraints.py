#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
硬约束抽取与冲突对账（日期 / 工期 / 有效期 / 质保期 / 响应时效）。

背景（为什么需要）：
    招标文件常在多处规定同一件事，且彼此不一致。真实案例：某招标文件公告写
    "11月15日前完成整体施工"，投标人须知又写"自合同签订之日起60个日历日内
    完成"，而开标日为9月24日——按60日排计划会直接突破11月15日。编写者若只
    抓住其中一条，排出的进度计划就是错的，且属于实质性不响应。

    本脚本抽取全部时间类硬约束，按主题归组，指出冲突并给出"以最严者为准"的
    建议，供编制进度计划与工期承诺时使用。

用法：
    python3 extract_constraints.py 招标文件.docx 技术规格书.docx
    python3 extract_constraints.py --txt 招标文件.txt --json out.json
"""
import argparse
import json
import re
import sys

from check_invented import read_any

# 主题识别关键词
TOPICS = {
    "交付/完工节点": ["完工", "完成整体施工", "交付", "竣工", "上线运行", "具备验收条件"],
    "总工期": ["总工期", "工期", "实施周期", "施工周期"],
    "进场时间": ["进场", "开工", "入场"],
    "投标截止/开标": ["投标截止", "开标", "递交投标文件"],
    "投标有效期": ["投标有效期"],
    "保证金/澄清等节点": ["保证金", "澄清", "答疑", "踏勘"],
    "质保期": ["质保期", "质保", "免费升级", "保修"],
    "响应时效": ["响应时间", "到场", "抵达现场", "小时内响应"],
    "服务/升级年限": ["免费升级", "服务年限", "维保期"],
    "付款节点": ["付款", "支付", "预付款", "质保金"],
}

# 日期与期限模式
P_DATE = re.compile(r"(20\d{2}\s*年\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日)?|\d{1,2}\s*月\s*\d{1,2}\s*日)")
P_DUR = re.compile(r"(\d+\s*(?:个)?\s*(?:日历日|工作日|自然日|天|日|个月|月|年))")
P_DAYS = re.compile(r"(\d+)\s*(?:个)?\s*(?:日历日|工作日|自然日)")


def split_sentences(text):
    text = re.sub(r"[ \t]+", " ", text)
    parts = re.split(r"(?<=[。；;\n])", text)
    return [p.strip() for p in parts if p.strip()]


def classify(sent):
    for topic, kws in TOPICS.items():
        if any(k in sent for k in kws):
            return topic
    return None


def extract(paths):
    found = {}
    for p in paths:
        text = read_any(p)
        import os
        doc = os.path.basename(p)
        for sent in split_sentences(text):
            if len(sent) > 400:
                continue
            topic = classify(sent)
            if not topic:
                continue
            dates = P_DATE.findall(sent)
            durs = P_DUR.findall(sent)
            if not dates and not durs:
                continue
            found.setdefault(topic, []).append({
                "文件": doc, "原文": sent[:220],
                "日期": [re.sub(r"\s", "", d) for d in dates],
                "期限": [re.sub(r"\s", "", d) for d in durs],
            })
    return found


# 绝对日期节点（"11月15日""2026年11月15日"），排除"2026年9月"这类进场时点
P_ABS = re.compile(r"^(?:20\d{2}年)?\d{1,2}月\d{1,2}日$")
# 锚点：投标截止 / 开标日期
P_ANCHOR = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")


def _to_date(y, m, d):
    import datetime
    return datetime.date(int(y), int(m), int(d))


def find_anchor(paths):
    """从文件中定位投标截止/开标日期，作为推算相对期限是否可行的锚点。"""
    for p in paths:
        text = read_any(p)
        for sent in split_sentences(text):
            if re.search(r"投标截止|开标时间|递交.*截止", sent):
                m = P_ANCHOR.search(sent)
                if m:
                    try:
                        return _to_date(*m.groups()), sent.strip()[:80]
                    except Exception:
                        pass
    return None, None


def detect_conflicts(found, anchor=None):
    """检测三类冲突：
       (1) 同主题多个不同期限值
       (2) 同主题同时存在"绝对日期节点"与"相对期限" ← 实战最危险的类型
       (3) 锚点 + 相对期限 已晚于绝对节点（推算不可行）
    """
    conflicts = []
    for topic, items in found.items():
        day_vals, abs_dates = set(), set()
        for it in items:
            for d in it["期限"]:
                m = P_DAYS.search(d)
                if m:
                    day_vals.add(int(m.group(1)))
            for d in it["日期"]:
                if P_ABS.match(d):
                    abs_dates.add(d)
        if len(day_vals) > 1:
            conflicts.append({
                "主题": topic, "类型": "多个期限值并存",
                "并存期限(日)": sorted(day_vals), "绝对日期节点": sorted(abs_dates),
                "建议": f"以最严者 {min(day_vals)} 日为准；若另有绝对日期节点，取两者中更早者",
            })
        # 绝对节点 + 相对期限：必须以绝对节点倒排
        if day_vals and abs_dates:
            mx = max(day_vals)
            cands = sorted(abs_dates)
            feasible = None
            if anchor:
                import datetime
                y = anchor.year
                est = []
                for ad in cands:
                    m = re.match(r"^(?:(20\d{2})年)?(\d{1,2})月(\d{1,2})日$", ad)
                    if not m:
                        continue
                    yy = int(m.group(1)) if m.group(1) else y
                    try:
                        est.append(_to_date(yy, m.group(2), m.group(3)))
                    except Exception:
                        pass
                if est:
                    earliest = min(est)
                    last_ok = anchor + datetime.timedelta(days=mx)
                    feasible = last_ok <= earliest
            conflicts.append({
                "主题": topic, "类型": "绝对节点与相对期限并存",
                "并存期限(日)": sorted(day_vals), "绝对日期节点": cands,
                "可行性推算": ("可行" if feasible else "★ 不可行") if feasible is not None else "无法推算（缺锚点）",
                "建议": (f"存在绝对日期节点 {cands}，该节点优先级最高，全部计划须以它倒排；"
                         f"自开工/签订起算 {mx} 日与绝对节点取更早者"),
            })
    return conflicts


def main():
    ap = argparse.ArgumentParser(description="硬约束抽取与冲突对账")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", help="结果另存为 JSON")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    anchor, anchor_src = find_anchor(args.files)
    if anchor and not args.quiet:
        print(f"锚点日期（来自『{anchor_src}』）：{anchor}")
    found = extract(args.files)
    if not found:
        print("未抽取到时间类硬约束，请确认输入文件是否为招标文件/技术规格书。")
        return 0

    for topic, items in found.items():
        print(f"\n{'='*74}\n【{topic}】共 {len(items)} 条\n{'='*74}")
        for it in items:
            tag = " ".join(it["日期"] + it["期限"])
            print(f"  [{it['文件']}] {(tag or '(无明确数值)'):<22} {it['原文']}")

    conflicts = detect_conflicts(found, anchor)
    print(f"\n{'='*74}\n★ 冲突检测\n{'='*74}")
    if conflicts:
        for c in conflicts:
            print(f"  ⚠ 【{c['主题']}】{c.get('类型','')}")
            print(f"       并存期限(日)：{c['并存期限(日)']}   绝对日期节点：{c.get('绝对日期节点') or '无'}")
            if '可行性推算' in c:
                print(f"       可行性推算：{c['可行性推算']}")
            print(f"       → {c['建议']}")
        print("\n  处置要求：编制进度计划前必须先消除上述冲突；")
        print("            存在绝对日期节点时，该节点优先级最高，全部计划以它倒排。")
    else:
        print("  未发现同主题期限冲突。")

    print("\n  提醒：请特别核对'绝对日期节点'与'相对期限'是否自洽——")
    print("        开标/合同签订日 + 相对期限 是否早于 绝对日期节点。")

    if args.json:
        json.dump({"constraints": found, "conflicts": conflicts},
                  open(args.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"\n  已写出：{args.json}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
