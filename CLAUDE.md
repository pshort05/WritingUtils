# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

WritingUtils is a collection of Python utilities for formatting creative writing files for publishing. There are two tools:

- **`clean-markdown`** — formats Markdown files with paragraph indentation and blank-line rules (no external dependencies)
- **`clean-docx`** — formats Word `.docx` files for KDP/print-on-demand publishing (requires `python-docx`, `pyyaml`)

## Project structure

```
src/writing_utils/
    __init__.py
    clean_markdown.py
    clean_docx.py
tests/
    test_sample.md       # fixture: input
    output_sample.md     # fixture: expected output
pyproject.toml
push.sh                  # git push utility (reads from commit_message file)
```

## Git workflow

This project uses `push.sh` for all git operations. **Do not commit or push directly.**

- `push.sh` reads from a `commit_message` file in the project root, then commits and pushes all changes, then deletes the file.
- After completing any work session, update (or create) the `commit_message` file with a summary of changes made. Update it regularly as work progresses, or when asked.
- The `commit_message` file is listed in `.gitignore` and is never committed itself.

## Installation

```bash
pip install -e .
```

This installs the `clean-markdown` and `clean-docx` entry points.

---

## clean-markdown

Source: `src/writing_utils/clean_markdown.py`

Processes Markdown files to add paragraph indentation and normalize spacing.

```bash
clean-markdown -i <input_file.md> -o <output_file.md>
```

Test fixtures: `tests/test_sample.md` (input), `tests/output_sample.md` (expected output).

### Architecture

1. **`is_markdown_structure(line)`** — returns True for lines that must not be indented: headers, lists, blockquotes, code fences, horizontal rules, links, HTML tags, bold/italic-only scene metadata.
2. **`clean_markdown(content)`** — single-pass processor:
   - Removes single blank lines between regular paragraphs; preserves blank lines adjacent to structural elements and multi-blank-line scene breaks.
   - Prepends 4 spaces to the first line of each paragraph unless `is_markdown_structure()` is True.
   - "New paragraph" is detected from the **source** position (preceded by blank line in input), not the output.

---

## clean-docx

Source: `src/writing_utils/clean_docx.py`

Cleans and formats Word `.docx` files for publishing. All features are optional and combinable. Supports a YAML config file.

### Quick start

```bash
# Using a config file (recommended)
clean-docx -c thunderwing.yaml

# CLI only
clean-docx -i input.docx -o output.docx --clean --page-breaks

# Config file with CLI overrides
clean-docx -c thunderwing.yaml --no-page-breaks --log-level DEBUG
```

### All flags

| Flag | Description |
|---|---|
| `-c FILE` | YAML config file. CLI flags override config values. |
| `-i FILE` | Input `.docx` file |
| `-o FILE` | Output `.docx` file |
| `--start TEXT` | Begin range at first heading containing TEXT |
| `--end TEXT` | Stop range before first heading containing TEXT |
| `--clean` / `--no-clean` | Remove extra empty paragraphs within range |
| `--page-breaks` / `--no-page-breaks` | Insert page break before each Heading 1 in range |
| `--fix-hrules` / `--no-fix-hrules` | Convert fixed-width HR shapes to margin-relative borders |
| `--font-to FONT` | Convert all non-skipped fonts to this font |
| `--font-skip FONT` | Font to preserve unchanged (repeatable) |
| `--log-level LEVEL` | `NONE` \| `ERROR` \| `INFO` \| `DEBUG` (default: NONE) |
| `--log-file FILE` | Log destination (default: stderr when logging enabled) |

### Config file format (`thunderwing.yaml`)

```yaml
input:  "Thunderwing Book 1.docx"
output: "Thunderwing Book 1 - clean.docx"

start: "Prologue"        # heading text match (case-insensitive substring)
end:   "The Luminarch"   # processing stops BEFORE this heading

clean:       true
page-breaks: true
fix-hrules:  true

font-to: "Times New Roman"
font-skip:
  - "Roboto Mono"

log-level: INFO          # NONE | ERROR | INFO | DEBUG
log-file:  thunderwing.log   # omit to log to stderr
```

### Feature details

**`--clean`**
Removes empty paragraphs within the range using three rules (applied in this order):
1. Empty heading paragraphs are always removed (Word export artifacts).
2. Single empty `normal` paragraphs between content are removed (double-Enter typing habit).
3. Runs of 2+ consecutive empty paragraphs are collapsed to 1 (scene break preserved).
Paragraphs with a bottom border (horizontal rules) are never removed.

**`--fix-hrules`**
Google Docs exports `---` as an inline drawing shape (`mc:AlternateContent`) with a hardcoded pixel width. This breaks with mirror margins for 2-sided printing. This feature replaces them with `w:pBdr` bottom borders, which are margin-relative and always span correctly. **Must run before `--clean`** — the code enforces this automatically.

**`--page-breaks`**
Inserts an explicit `<w:br w:type="page"/>` paragraph immediately before each non-empty Heading 1 in the range.

**`--font-to` / `--font-skip`**
Updates `w:rFonts` elements in four locations: document defaults, style definitions, paragraph-level run properties, and individual run properties. If **any** font attribute on a `w:rFonts` element matches a skip font, the entire element is left untouched (preserves monospace runs wholesale). Font conversion is document-wide (not range-restricted).

### Execution order

Features always run in this fixed order regardless of flag order:
1. `--fix-hrules` (must precede clean so borders survive removal)
2. `--clean`
3. `--page-breaks`
4. `--font-to`

### Logging levels

| Level | What is logged |
|---|---|
| `NONE` | Nothing (print output only) |
| `ERROR` | Failed heading lookups, file errors |
| `INFO` | Per-feature summaries with counts and timestamps |
| `DEBUG` | Every paragraph touched, every font replacement, every page break |

### Key architectural notes

- `--start` / `--end` match headings by **case-insensitive substring** — `"Prologue"` matches `"Prologue - Sky Terror, Eleven Years Ago"`.
- The `--end` heading itself is **not** processed — the range stops at the paragraph immediately before it.
- After `--clean` removes paragraphs, paragraph indices shift. The script re-resolves `(start_idx, end_idx)` silently before each subsequent feature to keep indices accurate.
- Boolean flags use `BooleanOptionalAction` — both `--clean` and `--no-clean` exist, allowing CLI to override a `true` set in the config file.
