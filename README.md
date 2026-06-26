# Moi Bid Response Skill · 投标技术响应文件编写

![License](https://img.shields.io/badge/License-AGPL--3.0-blue?style=flat-square)
![Skill](https://img.shields.io/badge/Skill-Agent-111111?style=flat-square)
![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat-square)
![Claude Code](https://img.shields.io/badge/Claude%20Code-Supported-6B5B95?style=flat-square)

一个适配 Claude Code 等 Agent 环境的投标技能，用于根据招标文件技术规格书（docx）**自动编写投标技术响应文件**（docx）。

核心流程：解析招标文件结构 → 检测招标自带响应模板 → 生成带分类和篇幅约束的响应大纲 → 逐章编写技术响应 → 汇总输出格式化 Word 文档。

> 由 [MoiTempete](https://github.com/MoiTempete) 基于煤矿安全、信息化平台等多个真实项目的中标响应文件沉淀而成。

## 30 秒开始

```bash
npx skills add https://github.com/MoiTempete/moi-bid-response --skill moi-bid-response
```

也可以直接把这段话发给有 shell 权限的 AI Agent：

```text
帮我安装 moi-bid-response skill。请把 https://github.com/MoiTempete/moi-bid-response 克隆到 ~/.claude/skills/moi-bid-response，安装完成后检查 SKILL.md、references/、scripts/ 是否存在。
```

已经安装过的话，用这段话更新：

```text
帮我更新 moi-bid-response skill。请进入 ~/.claude/skills/moi-bid-response 执行 git pull，然后告诉我当前最新 commit。
```

安装后直接对 Agent 说：

```text
帮我根据这份招标文件写技术响应文件。
```

也可以试这些请求：

```text
分析这份招标文件，提取重点标记和响应模板要求。
生成响应大纲，先确认结构再开始写。
把这些确认的响应内容汇总成格式化 Word 文档。
```

## 效果

- 📋 **自动解析招标文件**：提取章节结构、条款编号、重点标记（加粗/标星/标色/关键词），检测附件格式和响应模板
- 🎨 **三种组织方式**：按招标模板组织 / 软件工程式组织 / 按章节一一对应
- 🏷 **智能分类标注**：区分核心·标星 / 核心 / 制式·重点 / 制式，附带篇幅约束建议
- ✍️ **两种合规风格**：隐式叙述（企业级营销型）和显式引用（工程型逐条对标）
- 📄 **输出格式化 Word**：A4 页面，国标多级自动编号（一、/ 1. / 1.1. / 1.1.1.），宋体/黑体规范排版
- 🔍 **覆盖度保障**：大纲阶段生成条款覆盖映射，确保招标每条要求都有对应响应
- ⚡ **逐章确认**：每章写完后展示给用户即时调整，核心章节优先

## 适合 / 不适合

**✅ 合适**：信息化平台 / 安全系统 / 设备采购 / 工程服务等技术型招标响应，需要逐条对标、格式规范的正式投标文件

**❌ 不合适**：纯商务标书（报价单、资质证明等）、纯法律合同、需要设计排版的宣传册式方案

## 常见使用场景

| 任务 | 推荐方式 |
|------|---------|
| 招标方有附件格式 | 选方式 A，按招标模板组织，最稳妥 |
| 招标无模板，技术导向 | 选方式 B，软件工程式结构，展示专业性 |
| 用户要求逐章对应 | 选方式 C，按招标章节一一映射 |
| 大型国企/集团招标 | 用风格 1（隐式叙述），更像企业方案 |
| 技术型单位/煤矿/工厂 | 用风格 2（显式引用），评委逐条对标的安心感 |
| 仅需分析招标文件 | 直接跑 `scripts/parse_tender.py` 看结果 |

## 为什么是 Agent Skill

- **招标文件结构多变**：每条招标文件格式不同，Agent 更适合做适应性解析和判断
- **响应需要上下文理解**：不能简单模板填空，需要理解技术条款的隐含要求
- **篇幅需要权衡**：核心功能要展开、制式内容要克制，Agent 判断力比脚本更灵活
- **逐章确认更安全**：投标文件容错率低，人机协作逐章确认保障质量
- **格式自动化**：写完内容后 Word 排版、多级编号完全自动生成，省去手动调格式

## 安装

### 方式一：一行命令安装（推荐）

```bash
npx skills add https://github.com/op741MoiTempete8/moi-bid-response-skill --skill moi-bid-response
```

### 方式二：把下面这段话直接发给 AI

> 帮我安装 `moi-bid-response` 这个 Claude Code skill。请按下面步骤做：
>
> 1. 确保 `~/.claude/skills/` 目录存在（不存在就创建）
> 2. 执行 `git clone https://github.com/MoiTempete/moi-bid-response.git ~/.claude/skills/moi-bid-response`
> 3. 验证：`ls ~/.claude/skills/moi-bid-response/` 应该看到 `SKILL.md`、`references/`、`scripts/` 三项
> 4. 告诉我安装好了，之后我说"写投标响应"之类的话就会触发这个 skill

把这段话复制粘贴给 Claude Code / Cursor / 任何有 shell 权限的 AI Agent，它会自动完成安装。

### 方式三：手动命令行

```bash
git clone https://github.com/MoiTempete/moi-bid-response.git ~/.claude/skills/moi-bid-response
```

### 触发方式

装好后，Claude Code 会在对话里自动发现并调用这个 skill。触发关键词：

- "帮我写投标技术响应文件"
- "根据这份招标文件生成响应"
- "帮我做标书应答"
- "分析招标文件并写技术方案"
- "bid response"
- "tender response"

## 使用流程

Skill 本身是结构化工作流，Agent 会逐步引导：

1. **解析招标文件** — 运行 `parse_tender.py` 提取结构、重点标记、附件格式
2. **风格配置** — 分析招标特征，推荐组织方式（A/B/C）和合规风格（1/2），用户确认
3. **生成响应大纲** — 带分类标注 + 篇幅约束 + 覆盖映射，等用户确认
4. **逐章编写响应** — 核心优先，每章写完后即时调整
5. **输出 Word 文档** — 结构化 JSON → `generate_docx.py` → 格式化 .docx

详细说明见 [`SKILL.md`](./SKILL.md)。

## 三种组织方式

Skill 支持三种响应组织方式，按优先级自动判断：

| 方式 | 触发条件 | 说明 |
|------|---------|------|
| **A. 按招标模板** | 招标含"附件格式"、"响应文件格式"等 | 直接按招标要求章节组织，最稳妥 |
| **B. 软件工程式** | 招标无模板规定 | 项目概述 → 设计原则 → 功能方案 → 性能方案 → 实施验收 → 售后服务 |
| **C. 章节对应式** | 用户明确要求 | 按招标技术规格书章节一一对应 |

## 两种合规标注风格

基于真实中标文件分析，提供两种响应写法：

**风格 1 — 隐式叙述**（企业级/营销型标书）：
> 平台基于 B/S 架构构建，支持主流浏览器访问，并兼容 PC 端与移动端一体化使用。内置统一权限管理体系，提供基于角色的细粒度权限控制。

**风格 2 — 显式引用**（工程型/技术型标书）：
> 本章节满足"（二）系统功能要求——1、PC端系统功能"中"（1）综合展示：结合煤矿现有的采掘工程平面图，多维展示风险点信息..."的相关要求。
>
> 核心数据总览看板设计：系统登录后默认展示总览看板，涵盖风险点总数...

## 目录结构

```
moi-bid-response/
├── SKILL.md              ← Skill 主文件：工作流、原则、常见错误
├── README.md             ← 本文件
├── scripts/
│   ├── parse_tender.py       ← 招标 docx 解析脚本
│   └── generate_docx.py      ← 响应 docx 生成脚本
└── references/
    └── writing-guide.md      ← 响应编写规范、模板、篇幅控制指南
```

## 核心编写原则

1. **比招标书更具体**：增加数字、指标、流程细节。招标说"支持高并发"→ 响应写"支持 1000+ 用户同时在线，每秒 200+ 并发请求"
2. **综合叙述不逐条**：一个章节内多条要求融合为连贯的技术叙述，不写成 FAQ
3. **核心功能独立展开**：每个关键场景/功能单独成段，有独立的标题和指标
4. **制式内容标准化**：售后/培训等章节按标准模板展开，但仍然具体
5. **不确定处留标记**：无法确认的技术细节标注 `【待补充】`
6. **确保全覆盖**：大纲阶段生成覆盖映射，确保每条招标要求都有对应响应

## Roadmap

- 支持更多招标格式（PDF 解析）
- 扩展行业模板库（煤矿安全、信息化、设备采购等）
- 增加商务偏离表 / 技术偏离表自动生成
- 支持评分点分析和高分策略建议
- 支持多轮迭代修改和版本对比

## FAQ

**可以输出 PDF 吗？**
当前核心交付是 docx。可以在 Word 中另存为 PDF，或使用 libreoffice 等工具命令行转换。

**支持非中文招标文件吗？**
当前解析和响应均为中文优化。英文招标文件可尝试但效果可能下降。

**能保证中标吗？**
不能。这个 Skill 帮助你高效、规范地编写技术响应，确保覆盖度和专业度，但中标取决于综合评分。

**怎么更新到最新版？**
重新运行安装命令，或在本地 skill 目录执行 `git pull`。

## 贡献

Bug、响应质量问题、新行业模板需求——欢迎开 Issue 或 PR。改动请优先：

- 在 `references/writing-guide.md` 中补充新的编写模板和案例
- 脚本改动需保持 JSON 输入输出接口兼容
- 新增解析规则同步更新 `parse_tender.py` 的标记逻辑

详见 [`CONTRIBUTING.md`](./CONTRIBUTING.md)。

## License

AGPL-3.0 © 2026 [MoiTempete](https://github.com/MoiTempete)
