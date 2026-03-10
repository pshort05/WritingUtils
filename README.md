# WritingUtils

Python utilities for formatting creative writing files for publishing.

- **`clean-markdown`** — formats Markdown files with paragraph indentation and blank-line rules
- **`clean-docx`** — formats Word `.docx` files for KDP/print-on-demand publishing

## Installation

```bash
pip install -e .
```

This installs the `clean-markdown` and `clean-docx` commands.

## Usage

```bash
# Format a Markdown file
clean-markdown -i input.md -o output.md

# Format a .docx file (with config file)
clean-docx -c mybook.yaml

# Format a .docx file (CLI only)
clean-docx -i input.docx -o output.docx --clean --page-breaks
```

See `CLAUDE.md` for full documentation of all options.

## License

MIT — see [LICENSE](LICENSE).
