# WritingUtils Project Memory

## Current state (2026-03-10)

**clean-docx**: Complete with all 6 features working and tested on Thunderwing1.docx.
**format-docx**: Not started (Phases 1–7 pending). Entry point registered in pyproject.toml.

## Key architectural lessons

### OOXML pPr child ordering (critical)
`w:pBdr` MUST appear BEFORE `w:rPr` in `w:pPr`. Appending to end of pPr places it after rPr,
which is schema-invalid — Word silently ignores the border. Always insert before rPr:
```python
rPr_elem = pPr.find(qn("w:rPr"))
if rPr_elem is not None:
    rPr_elem.addprevious(new_elem)
else:
    pPr.append(new_elem)
```
Correct order: pStyle → keepNext → keepLines → pageBreakBefore → pBdr → shd → tabs → spacing → ind → rPr

### fix_hrules: two cases
- Case 1: `mc:AlternateContent` drawing shape (Google Docs `---` export) — remove run, add pBdr
- Case 2: empty Heading 2+ paragraphs (scene break separators) — add pBdr, skip if has page-break run

### excerpt-font font detection
Use `w:rFonts` XML (not python-docx `r.font.name`) — covers paragraph-level and run-level settings.
Check all four attributes: `w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`.

### --clean interaction with excerpt-font
`--clean` Rule 2 removes single empty paragraphs between content — strips blanks around individual
excerpt paragraphs. `format_excerpts` runs AFTER clean and re-inserts blank paragraphs as needed.

## Next session: start format-docx Phase 1

User has remaining issues from clean-docx test (said "several items, starting with #1").
Issue #1 (horizontal rules) is now fixed. Ask user about remaining issues before starting format-docx.
