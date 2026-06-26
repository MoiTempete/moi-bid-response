# Contributing

Thanks for helping improve `moi-bid-response`.

This project is a Skill for AI agents that generate bid technical response documents. The most useful contributions are specific, reproducible, and tied to real tender response output.

## Before Opening an Issue

Please check whether the problem belongs to one of these buckets:

- Parsing accuracy: the tender parser misses chapter structure, emphasis markers, or appendix templates.
- Response quality: generated content is too generic, misses clauses, or gets the compliance style wrong.
- Word output: formatting, numbering, fonts, or table layout issues in the generated .docx.
- Documentation: installation, usage, or workflow instructions are unclear.
- Script behavior: `parse_tender.py` or `generate_docx.py` produce errors or unexpected output.

When reporting, please include:

- The tender document (if shareable) or a description of its structure.
- The response generated and the specific issues.
- The agent and environment used (Claude Code, Codex, etc.).

## Pull Request Guidelines

Keep PRs focused. A small fix with a clear before/after example is easier to review than a large rewrite.

### For parsing changes (`scripts/parse_tender.py`):

- Test against at least one real tender document (or a representative fixture).
- New marker rules should have clear keyword/logic comments.
- Keep the JSON output schema backward-compatible.

### For generation changes (`scripts/generate_docx.py`):

- Verify the generated .docx opens correctly in Microsoft Word and LibreOffice.
- Check that multi-level numbering renders as expected.
- Verify table formatting, margins, and font styles.

### For writing guide changes (`references/writing-guide.md`):

- New templates should include both the pattern and a concrete example.
- Indicate which compliance style (1=implicit, 2=explicit) the template is designed for.
- Include word-count guidance for different chapter types.

### For workflow changes (`SKILL.md`):

- Keep the 5-step workflow structure clear.
- Each step should state its input, output, and exit criteria.
- Add common error cases and how to handle them.

## Good PRs Usually Include

- A short summary of the problem.
- The exact files changed.
- Before/after output snippets.
- Test results or validation notes.

## Style Notes

This Skill is opinionated by design. It prefers well-structured, compliant response generation over unlimited creative freedom, because constraints make AI-generated bid documents more reliable.

When in doubt, preserve the existing workflow structure and improve the quality around it.

## Key Design Decisions

- **Style must be confirmed before outline**: Prevents mismatch between content and format.
- **Outline must be confirmed before writing**: Prevents wasted writing on wrong structure.
- **Core chapters first**: Ensures the most important content gets attention if time is limited.
- **Coverage mapping in outline**: Makes sure every tender clause is addressed before writing starts.
- **Chapter-by-chapter confirmation**: Allows course correction without rewriting the whole document.
