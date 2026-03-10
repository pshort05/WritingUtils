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

---

## format-docx — Implementation Plan

New utility for producing platform-specific publishing versions (KDP, print) from a single source `.docx`. Configured by separate YAML files (`kdp.yaml`, `print.yaml`).

### Typical pipeline

```bash
# Step 1: clean the raw export once
clean-docx -c clean.yaml

# Step 2: format for each platform (from the cleaned file)
format-docx -c kdp.yaml
format-docx -c print.yaml
```

### Architecture

- **New file:** `src/writing_utils/format_docx.py`
- **Shared utilities:** `load_config()` and `setup_logging()` extracted from `clean_docx.py` into `src/writing_utils/_util.py`; both tools import from there
- **Reused functions** imported directly from `clean_docx.py` (no modification): `fix_hrules`, `collect_removals`, `insert_page_breaks`, `convert_fonts`, `find_heading`, `resolve_range`

### Execution order inside format-docx

1. `fix-hrules` (must precede clean)
2. `clean`
3. `page-breaks`
4. `font-to` (document-wide)
5. `page-size` (section geometry)
6. `margins` (includes mirror-margins XML)
7. `spacing` (style-level defaults — applied globally)
8. `body-format` (per-paragraph overrides within range — skips headings, centered/right text)
9. `headers` (needs page geometry for tab stop positions)
10. `footers`

### New YAML keys (beyond what clean-docx already supports)

```yaml
page-size:
  width:  "5.5in"        # accepts: in, mm, cm, pt
  height: "8.5in"

margins:
  top:     "1in"
  bottom:  "1in"
  inside:  "0.875in"     # binding side (mirror margins)
  outside: "0.75in"
  gutter:  "0in"
  header:  "0.5in"       # top-of-page to header top
  footer:  "0.5in"       # bottom-of-page to footer bottom

mirror-margins: true     # enables facing-pages mode

spacing:                 # applied to named styles, not inline paragraphs
  normal:
    line-spacing:      "double"   # single | 1.5 | double | exactly:12pt | multiple:1.5
    first-line-indent: "0.5in"
    space-before:      "0pt"
    space-after:       "0pt"
  heading-1:
    space-before: "12pt"
    space-after:  "6pt"

doc-title:  "My Book Title"   # substituted for {title} token in headers/footers
doc-author: "Author Name"     # substituted for {author} token

header-mode: "odd-even"       # none | uniform | odd-even | first-different
header:
  odd:
    left:   ""
    center: "{title}"          # tokens: {title} {author} {page}
    right:  "{page}"
    font:   "Times New Roman"
    size:   "10pt"
    italic: true
  even:
    left:   "{page}"
    center: "{author}"
    right:  ""
    font:   "Times New Roman"
    size:   "10pt"
    italic: true

footer-mode: "none"            # none | uniform | odd-even
```

### python-docx vs. raw XML

Features requiring raw XML (lxml):
- **Mirror margins** — inject `<w:mirrorMargins/>` into `doc.settings.element`
- **Page number field** — `{page}` token becomes `<w:fldChar>`/`<w:instrText> PAGE </w:instrText>` sequence; no public API
- **Three-part header layout** — tab stops at computed center/right positions injected as `w:tabs` XML

Everything else (page size, margins, header text/font, odd/even header enable, style spacing) uses the python-docx public API.

### Implementation checklist

#### Phase 0 — Shared utilities
- [x] Create `src/writing_utils/_util.py` — move `load_config()`, `setup_logging()`, `_LOG_LEVELS` from `clean_docx.py`
- [x] Update `clean_docx.py` to import from `_util`; verify existing behavior unchanged

#### Phase 1 — Skeleton
- [ ] Create `format_docx.py` with `main()`, argparse, `apply_config()` — opens doc, saves unchanged (no-op)
- [ ] Add `format-docx` entry point to `pyproject.toml`; run `pip install -e .`
- [ ] Implement `parse_length(s) -> EMU int` — handles `"0.75in"`, `"19mm"`, `"12pt"`

#### Phase 2 — Page geometry
- [ ] `set_page_size(section, width_str, height_str)`
- [ ] `set_margins(section, cfg)` — maps `inside`/`outside` to `left_margin`/`right_margin`
- [ ] `enable_mirror_margins(doc)` — injects `<w:mirrorMargins/>` into `doc.settings.element`
- [ ] Wire `page-size`, `margins`, `mirror-margins` into `main()`; test with a real docx

#### Phase 3 — Style spacing
- [ ] `parse_line_spacing(s) -> (rule, value)` — handles all five format strings
- [ ] `apply_style_spacing(doc, style_name, cfg)` — normalizes style names, sets `paragraph_format.*`
- [ ] Wire `spacing:` into `main()`

#### Phase 4 — Headers and footers
- [ ] `_make_field_run(instruction)` — lxml sequence for `PAGE` field codes
- [ ] `_set_tab_stops(paragraph, center_twips, right_twips)` — injects `w:tabs` XML
- [ ] `build_header_paragraph(hdr_obj, left, center, right, font, size, italic, page_geom)`
- [ ] `set_headers(doc, section, cfg)` — dispatches by `header-mode`
- [ ] `set_footers(doc, section, cfg)`
- [ ] Wire headers and footers into `main()`

#### Phase 5 — Reuse clean_docx features
- [ ] Import and wire `fix_hrules`, `collect_removals`, `insert_page_breaks`, `convert_fonts`, `find_heading`, `resolve_range` into `format_docx.main()`

#### Phase 6 — Body paragraph formatting
Applies inline paragraph formatting (indent, spacing) to body paragraphs within the range, skipping headings, centered/right-justified text, and any explicitly listed styles. Complements `spacing:` (which sets style-level defaults globally) with per-paragraph overrides for body text only.

New YAML key:
```yaml
body-format:
  first-line-indent: "0.5in"
  line-spacing:      "double"    # same format strings as spacing:
  space-before:      "0pt"
  space-after:       "0pt"
  skip-styles:                   # styles to leave untouched (in addition to auto-skips)
    - "Block Text"
    - "Caption"
    - "Epigraph"
```

Auto-skip rules (no config needed):
- Any `Heading *` style
- Paragraph alignment is centered (`WD_ALIGN_PARAGRAPH.CENTER`)
- Paragraph alignment is right-justified (`WD_ALIGN_PARAGRAPH.RIGHT`)
- Paragraph alignment is distributed (`WD_ALIGN_PARAGRAPH.DISTRIBUTE`)

Implementation:
- [ ] `should_skip_paragraph(p, skip_styles)` — returns True for headings, non-left-aligned, or style in skip list
- [ ] `apply_body_format(paragraphs, start_idx, end_idx, cfg, skip_styles)` — walks range, calls `should_skip_paragraph`, applies `paragraph_format.*` directly to each qualifying paragraph's `p.paragraph_format`
- [ ] Wire `body-format:` into `main()` execution order (runs after `spacing:`, before headers)

#### Phase 7 — End-to-end
- [ ] Write `kdp.yaml` with KDP trim size (e.g. 5.5×8.5in) and margin specs
- [ ] Write `print.yaml` with print publisher trim size and margin specs
- [ ] End-to-end test on the Thunderwing docx
- [ ] Update this section of `CLAUDE.md` with completed `format-docx` documentation
