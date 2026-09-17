#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
参考素材能力清单提取器 —— 把往期标书/产品资料拆成可对齐的结构化清单。

背景（为什么需要）：
    "多参考现有素材"是提高中标率、避免过度设计的有效手段，但素材动辄十几万
    字，直接投喂既超上下文又抓不住重点。且两类素材用法不同：
      · 制式内容（组织架构/售后/质量/人员/培训）→ 取其真实承诺强度与措辞
      · 技术方案（平台/集控/子系统接入）        → 取其已实现的技术路径与边界
    本脚本先给出结构骨架与关键事实，再由编写者按需深读。

用法：
    python3 extract_capability.py 参考素材A.docx 参考素材B.docx -o 素材能力清单.md
    python3 extract_capability.py --dir ./子系统资料 -o 资料清单.md

⚠ 当单个素材超过约 10 万字时，建议改用 subagent 并行提取（见 SKILL.md Step 1b）。
"""
import argparse
import glob
import os
import re
import sys

from check_invented import read_any

# 制式内容关注的维度
FORMAL = {
    "资质与认证": r"ISO\d+|CMMI|软件著作权|著作权登记|CS\s*信息系统|ITSS|CCRC|售后服务认证|高新技术企业|软件企业",
    "信创适配": r"信创|国产化|麒麟|统信|达梦|人大金仓|南大通用|东方通|中创|鲲鹏|龙芯|海光|飞腾",
    "售后与质保": r"质保期|免费升级|响应时间|到场|7[×x*]24|服务组织|服务总监|技术支持经理|工单|闭环",
    "质量体系": r"质量方针|质量目标|技术债务率|代码重复率|单元测试覆盖率|缺陷密度|四级测试|缺陷全生命周期",
    "人员与组织": r"项目经理|技术负责人|系统架构师|驻场|PMP|信息系统项目管理师|人员稳定|备份",
    "实施与进度": r"实施方法论|实施阶段|里程碑|进度保障|关键路径|缓冲|Sprint|迭代批次",
    "培训与知识转移": r"培训对象|培训内容|培训方式|培训教材|知识转移|操作手册|运维手册",
    "保密与安全": r"信息安全|保密|加密|等保|审计日志|数据分级|责任追究",
}
# 技术方案关注的维度
TECH = {
    "总体架构": r"分层架构|技术架构|微服务|容器|云原生|中台|总体设计原则",
    "SCADA/组态": r"组态|SCADA|I/O采集|采集端|开发站|操作员站|可视化|画面",
    "物联网平台": r"物联网|物模型|边缘网关|MQTT|CoAP|设备接入|规则引擎|持久化",
    "GIS": r"GIS|地理信息|一张图|点位|图层|空间查询|绑点",
    "数据治理": r"数据治理|数据标准|元数据|数据质量|主数据|数据资产|血缘",
    "集控/集控中心": r"集控|集中控制|远程控制|联锁|一键启停|无人值守|顺控",
    "子系统接入": r"接入方式|OPC|Modbus|Profinet|IEC\s*104|RS485|协议转换|点表",
    "性能与可靠性": r"可用性|冗余|热备|无扰切换|断点续传|RTO|RPO|并发|采集频率",
}
NUM = re.compile(r"\d+(?:\.\d+)?\s*(?:%|个日历日|个工作日|小时|分钟|秒|天|年|个|台|套|万|点/秒)")


def analyze(path):
    txt = read_any(path)
    if not txt.strip():
        return None
    lines = [l.strip() for l in txt.split("\n") if l.strip()]
    headings = [l for l in lines if len(l) < 60 and not re.search(r"[。；，]", l)]
    hits = {}
    for label, pat in {**FORMAL, **TECH}.items():
        found = []
        for m in re.finditer(pat, txt):
            s = max(0, m.start() - 60)
            seg = txt[s:m.end() + 120].replace("\n", " ")
            found.append(seg)
        if found:
            hits[label] = found[:6]
    numbers = list(dict.fromkeys(NUM.findall(txt)))
    return {"chars": len(re.findall(r"[\u4e00-\u9fff]", txt)),
            "headings": headings, "hits": hits, "numbers": numbers[:80]}


def main():
    ap = argparse.ArgumentParser(description="参考素材能力清单提取")
    ap.add_argument("files", nargs="*")
    ap.add_argument("--dir", help="目录（递归读取）")
    ap.add_argument("-o", "--output", default="素材能力清单.md")
    args = ap.parse_args()

    paths = list(args.files)
    if args.dir:
        for root, _, fs in os.walk(args.dir):
            for fn in fs:
                if fn.startswith(".") or fn.startswith("~$"):
                    continue
                paths.append(os.path.join(root, fn))
    if not paths:
        ap.error("需要文件或 --dir")

    out = ["# 参考素材能力清单", "",
           "> 用途：制式内容取其真实承诺强度与措辞；技术方案取其已实现路径与边界。",
           "> 注意：素材中的**具体数值**（质保年限、响应时效、可用率、并发规模等）"
           "多数与本项目招标要求不符，引用前必须逐项替换。", ""]
    for p in sorted(paths):
        r = analyze(p)
        if not r:
            continue
        out += [f"## {os.path.basename(p)}", "",
                f"- 中文字数：{r['chars']}", "",
                "### 结构骨架（标题行）", ""]
        for h in r["headings"][:80]:
            out.append(f"- {h}")
        out += ["", "### 主题命中（含上下文片段）", ""]
        for label, segs in r["hits"].items():
            out.append(f"**{label}**")
            for s in segs:
                out.append(f"- …{s.strip()}…")
            out.append("")
        out += ["### 出现的量化数值（须逐项核对是否适用于本项目）", "",
                "、".join(r["numbers"]) if r["numbers"] else "（无）", "", "---", ""]

    open(args.output, "w", encoding="utf-8").write("\n".join(out))
    print(f"已生成 {args.output}（覆盖 {len(paths)} 个素材文件）")
    print("\n⚠ 使用提醒：")
    print("  1) 单个素材超过约 10 万字时，建议改用 subagent 并行深读（SKILL.md Step 1b）；")
    print("  2) '量化数值'一节列出的数字必须逐个核对——素材数值直接搬入新项目是")
    print("     最常见的隐蔽偏离（如素材质保1年、本项目要求3年）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
