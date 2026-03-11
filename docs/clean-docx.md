# clean-docx

Cleans and formats Word `.docx` files for publishing. All features are optional and freely combinable. Accepts a YAML config file, CLI flags, or both.

## Quick start

```bash
# Recommended: use a config file
clean-docx -c mybook.yaml

# CLI only
clean-docx -i input.docx -o output.docx --clean --page-breaks

# Config file with CLI overrides
clean-docx -c mybook.yaml --no-page-breaks --log-level DEBUG
```

## All flags

| Flag | Description |
|---|---|
| `-c FILE` | YAML config file. CLI flags override config values. |
| `-i FILE` | Input `.docx` file |
| `-o FILE` | Output `.docx` file |
| `--start TEXT` | Begin processing at the first heading containing TEXT |
| `--end TEXT` | Stop processing before the first heading containing TEXT |
| `--clean` / `--no-clean` | Remove extra empty paragraphs within range |
| `--page-breaks` / `--no-page-breaks` | Insert page break before each Heading 1 in range |
| `--fix-hrules` / `--no-fix-hrules` | Convert fixed-width HR shapes to margin-relative borders |
| `--indent-paragraphs` / `--no-indent-paragraphs` | Add a first-line indent to body text paragraphs within range |
| `--excerpt-font FONT` | Font name that identifies embedded written excerpts; applies 0.5" block indent with blank lines before/after |
| `--font-to FONT` | Convert all non-skipped fonts to this font |
| `--font-skip FONT` | Font name to preserve unchanged (repeatable) |
| `--log-level LEVEL` | `NONE` \| `ERROR` \| `INFO` \| `DEBUG` (default: `NONE`) |
| `--log-file FILE` | Log destination (default: stderr when logging is enabled) |

At least one action flag must be supplied.

## Config file

A YAML config file can supply any option that a CLI flag can. CLI flags always take precedence over config file values. The `--no-*` variants let you override a `true` value set in the config.

```yaml
input:  "My Novel.docx"
output: "My Novel - clean.docx"

start: "Prologue"        # heading text match (case-insensitive substring)
end:   "Acknowledgements"  # processing stops BEFORE this heading

clean:       true
page-breaks: true
fix-hrules:  true
indent-paragraphs: true

excerpt-font: "Roboto Mono"   # paragraphs using this font get 0.5" block indent

font-to: "Times New Roman"
font-skip:
  - "Roboto Mono"        # preserves the excerpt font unchanged during conversion

log-level: INFO          # NONE | ERROR | INFO | DEBUG
log-file:  clean.log     # omit to log to stderr
```

## Processing range

`--start` and `--end` limit which paragraphs each feature acts on. Both match headings by **case-insensitive substring**: `"Prologue"` matches `"Prologue — Sky Terror, Eleven Years Ago"`.

- The `--start` heading is **included** in the range.
- The `--end` heading is **excluded** — processing stops at the paragraph immediately before it.
- Omitting either flag extends the range to the beginning or end of the document.

Font conversion (`--font-to`) is always **document-wide** and ignores the range.

## Feature details

### `--clean`

Removes unwanted empty paragraphs within the range. Three rules are applied in order:

1. **Empty Heading 1 paragraphs** are removed. These are common Word/Google Docs export artifacts that appear alongside real chapter headings. Empty Heading 2 and deeper paragraphs are left in place — they are often intentional visual scene-break separators styled by the heading definition.
2. **Single empty normal paragraphs** between content paragraphs are removed. These come from the habit of pressing Enter twice between paragraphs.
3. **Runs of two or more consecutive empty paragraphs** are collapsed to one. The single surviving empty paragraph is treated as a scene break.

Paragraphs that carry a bottom border (horizontal rules, whether original or converted by `--fix-hrules`) are never removed by any rule.

### `--fix-hrules`

Google Docs exports `---` scene dividers as inline drawing shapes (`mc:AlternateContent`) with a hardcoded pixel width. This is fine for single-sided documents but breaks visually when you switch to mirror margins for two-sided print layout — the line no longer spans the text area correctly.

This feature replaces those drawing shapes with `w:pBdr` paragraph bottom borders. A `w:pBdr` border is always margin-relative and spans the full text width regardless of page geometry or margin settings.

`--fix-hrules` also handles **empty Heading 2+ scene-break paragraphs** that have no bottom border. These are used as scene dividers, but the Heading 2 style carries no border definition, so without an explicit `w:pBdr` they render as invisible whitespace. `--fix-hrules` applies the same bottom border to them, making them visible as horizontal rules. Heading 2 paragraphs that already contain a `<w:br type="page"/>` run are skipped — those are chapter-break containers, not scene separators, and should not be given a border.

**Must run before `--clean`** — the tool enforces this automatically regardless of flag order. Running before `--clean` also ensures the border is in place before `--clean` evaluates whether to preserve or remove these paragraphs.

The border applied in all cases is a 0.75pt single rule in medium gray (`#A0A0A0`).

### `--page-breaks`

Inserts an explicit page-break paragraph (`<w:br w:type="page"/>`) immediately before each non-empty Heading 1 within the range. Empty Heading 1 paragraphs (artifacts) are skipped. If the paragraph immediately preceding the Heading 1 in the XML already contains a page-break run (e.g. a Heading 2 chapter-break container from the original document), no additional page break is inserted.

This is distinct from Word's "page break before" paragraph property: the inserted paragraph is a standalone element that some export pipelines handle more reliably.

### `--indent-paragraphs`

Adds a first-line indent of 0.25 inches (approximately 3 characters at 12pt) to body text paragraphs within the range. This is applied as `paragraph_format.first_line_indent` directly on the paragraph, which takes precedence over any style-level default.

Paragraphs skipped automatically:

- **Empty paragraphs** — nothing visible to indent
- **Headings** — any paragraph with a `Heading *` style
- **Non-left-aligned paragraphs** — center-justified, right-justified, and distributed alignment are left untouched
- **Title components** — paragraphs where every non-empty run is bold, italic, or both, that appear immediately after a heading. This catches chapter subtitles (place names, character names, POV labels, dates) styled as normal text with emphasis formatting rather than a dedicated heading style. Paragraphs that mix bold and italic runs (e.g. a bold location followed by an italic character name in the same paragraph) are also recognized. Once a plain-text body paragraph is encountered, the title-component detection stops for that heading block — later bold/italic paragraphs in the body are indented normally.
- **Excerpt paragraphs** — paragraphs identified by `--excerpt-font` are skipped; they receive their own block indentation from that feature instead.

This is useful when importing from Google Docs or another source that uses blank lines between paragraphs instead of indentation, after `--clean` has removed those blank lines.

### `--excerpt-font`

Formats embedded written excerpts (letters, journal entries, diary pages, etc.) that are distinguished from narrative prose by using a different font.

All paragraphs using the specified font are treated as a block-quote excerpt and receive:
- **0.5" left indent** and **0.5" right indent** applied directly to each paragraph
- **A blank paragraph inserted before** the excerpt group if one is not already present
- **A blank paragraph inserted after** the excerpt group if one is not already present

Consecutive excerpt paragraphs (including any empty paragraphs between them) are treated as a single group, so only one blank line appears around the entire block rather than one per paragraph.

Font detection searches all `w:rFonts` elements within each paragraph, covering both run-level and paragraph-level font settings.

When used together with `--indent-paragraphs`, excerpt paragraphs are automatically excluded from first-line indentation — the block indent is applied instead.

When used together with `--font-to`, set `--font-skip` to the same font to preserve it during conversion:

```yaml
excerpt-font: "Roboto Mono"
font-to:      "Times New Roman"
font-skip:
  - "Roboto Mono"
```

### `--font-to` / `--font-skip`

Replaces fonts throughout the entire document with the specified font. Covers all four locations where fonts are stored in a `.docx`:

1. **Document defaults** — `w:docDefaults` in the styles part
2. **Style definitions** — `w:style` elements
3. **Paragraph-level run properties** — `w:pPr/w:rPr` overrides
4. **Individual run properties** — `w:r/w:rPr` inline overrides

`--font-skip` preserves fonts by name. If **any** font attribute on a `w:rFonts` element matches a skip font, the entire element is left untouched. This means a monospace run (e.g., `Roboto Mono`) is preserved wholesale even if it also sets `w:cs` or `w:eastAsia` attributes.

`--font-skip` is repeatable:

```bash
clean-docx -i input.docx -o output.docx --font-to "Garamond" --font-skip "Roboto Mono" --font-skip "Courier New"
```

## Execution order

Features always run in this fixed sequence regardless of the order flags appear on the command line:

1. `--fix-hrules` — converts drawing shapes to borders before `--clean` can touch them
2. `--clean` — removes empty paragraphs
3. `--page-breaks` — inserts page break paragraphs
4. `--indent-paragraphs` — adds first-line indents to body text (skips excerpt paragraphs)
5. `--excerpt-font` — applies block indent and surrounding blank lines to excerpt groups
6. `--font-to` — converts fonts document-wide

After `--clean` removes paragraphs, paragraph indices shift. The script re-resolves `(start_idx, end_idx)` silently before each subsequent step to keep indices accurate.

## Logging

| Level | What is logged |
|---|---|
| `NONE` | Nothing — only print output (default) |
| `ERROR` | Failed heading lookups, file errors |
| `INFO` | Per-feature summaries with counts and timestamps |
| `DEBUG` | Every paragraph touched, every font replacement, every page break |

Log output goes to stderr by default. Supply `--log-file` (or `log-file` in the config) to write to a file instead.

## Typical workflow

For a novel exported from Google Docs:

```bash
# 1. Create a config file for the book
cat > thunderwing.yaml <<EOF
input:  "Thunderwing Book 1.docx"
output: "Thunderwing Book 1 - clean.docx"
start:  "Prologue"
end:    "The Luminarch"
clean:       true
page-breaks: true
fix-hrules:  true
indent-paragraphs: true
excerpt-font: "Roboto Mono"
font-to:     "Times New Roman"
font-skip:
  - "Roboto Mono"
log-level: INFO
EOF

# 2. Run
clean-docx -c thunderwing.yaml
```

The cleaned output is then ready to pass to `format-docx` for platform-specific layout (KDP, print-on-demand, etc.).
