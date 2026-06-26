# Bid Response Skill · Tender Technical Response Generation

![License](https://img.shields.io/badge/License-AGPL--3.0-blue?style=flat-square)
![Skill](https://img.shields.io/badge/Skill-Agent-111111?style=flat-square)
![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat-square)
![Claude Code](https://img.shields.io/badge/Claude%20Code-Supported-6B5B95?style=flat-square)

> 🌏 **中文版本：[README.md](./README.md)**

An agent skill for Claude Code and similar coding-agent environments. It generates **formatted bid technical response documents (docx)** from tender technical specifications (docx).

Core workflow: parse tender document structure → detect tender-supplied response templates → generate a classified, word-count-constrained response outline → write technical responses chapter by chapter → assemble a formatted Word document.

> Distilled by [Guizang](https://x.com/op7418) from real winning bids in coal mine safety, information platforms, and other technical projects.

## 30-second start

```bash
npx skills add https://github.com/op7418/moi-bid-response-skill --skill moi-bid-response
```

Or send this to any AI agent with shell access:

```text
Install the moi-bid-response Claude Code skill. Clone https://github.com/op7418/moi-bid-response-skill into ~/.claude/skills/moi-bid-response and verify SKILL.md, references/, and scripts/ exist.
```

After installation, just say:

```text
Help me write a technical response for this tender document.
```

## Capabilities

- 📋 **Auto-parse tender documents**: Extract chapter structure, clause numbers, emphasis markers (bold/starred/colored/keywords), detect appendix formats and response templates
- 🎨 **Three organization modes**: Follow tender template / Software-engineering structure / Chapter-by-chapter mapping
- 🏷 **Smart classification**: Distinguish core-starred / core / standard-priority / standard, with word-count recommendations
- ✍️ **Two compliance styles**: Implicit narrative (enterprise marketing style) and explicit citation (engineering traceability style)
- 📄 **Formatted Word output**: A4 pages, multi-level auto-numbering, standardized font/layout
- 🔍 **Coverage guarantee**: Clause coverage mapping at the outline stage ensures no tender requirement is missed
- ⚡ **Chapter-by-chapter confirmation**: Review after each chapter, core chapters first

## Suitable / Not suitable

**✅ Suitable for**: Technical bids (IT platforms, safety systems, equipment procurement, engineering services) that require clause-by-clause compliance and formally formatted response documents

**❌ Not suitable for**: Pure commercial bids (price quotes, qualification certificates), legal contracts, or brochure-style proposals

## Installation

### Option 1: One-liner (recommended)

```bash
npx skills add https://github.com/op7418/moi-bid-response-skill --skill moi-bid-response
```

### Option 2: Send to AI agent

Copy and paste the installation instructions from the Chinese README installation section to any AI agent with shell access.

### Option 3: Manual

```bash
git clone https://github.com/op7418/moi-bid-response-skill.git ~/.claude/skills/moi-bid-response
```

## Workflow

The Skill is a structured workflow the agent follows step by step:

1. **Parse tender** — Run `scripts/parse_tender.py` to extract structure, emphasis, and templates
2. **Style configuration** — Analyze tender characteristics, recommend organization mode (A/B/C) and compliance style (1/2), wait for user confirmation
3. **Generate outline** — Classified chapters with word budgets and coverage mapping, wait for user approval
4. **Write chapter by chapter** — Core-first, user adjusts after each chapter
5. **Output Word** — Structured JSON → `scripts/generate_docx.py` → formatted .docx

Full details in [`SKILL.md`](./SKILL.md).

## Three organization modes

| Mode | Trigger | Description |
|------|---------|-------------|
| **A. Follow template** | Tender includes "appendix format" or "response template" | Follow the tender's required structure — safest option |
| **B. Software engineering** | No template specified | Overview → Design Principles → Functional → Performance → Implementation → Support |
| **C. Chapter mapping** | User explicitly requests | One-to-one mapping to tender technical chapters |

## Two compliance styles

**Style 1 — Implicit narrative** (enterprise/marketing bids):
Integration of requirements into flowing technical prose without explicit "fully compliant" labels.

**Style 2 — Explicit citation** (engineering/technical bids):
Each section opens with a citation of the exact tender clause it addresses, then expands with technical detail.

## Directory structure

```
moi-bid-response/
├── SKILL.md              ← Main Skill file: workflow, principles
├── README.md             ← This file (Chinese)
├── README.en.md          ← English README
├── scripts/
│   ├── parse_tender.py       ← Tender docx parser
│   └── generate_docx.py      ← Response docx generator
└── references/
    └── writing-guide.md      ← Writing standards, templates, word-count guides
```

## Core principles

1. **More specific than the tender**: Add numbers, metrics, process details
2. **Integrated narrative, not FAQ**: Merge related clauses into cohesive technical descriptions
3. **Core features expanded independently**: Each key feature/scenario gets its own section with headings and metrics
4. **Standardized but concrete**: Support/training sections follow templates but remain specific
5. **Mark unknowns**: Use `【TBD】` for unconfirmed technical details
6. **Full coverage guaranteed**: Coverage mapping at outline stage ensures every clause is addressed

## Contributing

Bug reports, response quality issues, new industry templates — open an Issue or PR. Prioritize:

- Add new templates and examples to `references/writing-guide.md`
- Keep script JSON I/O interfaces compatible
- Update `parse_tender.py` marker logic for new parsing rules

See [`CONTRIBUTING.md`](./CONTRIBUTING.md).

## License

AGPL-3.0 © 2026 [op7418](https://github.com/op7418)
