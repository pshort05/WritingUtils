# WritingUtils

Python utilities for formatting creative writing files for publishing.

- **`clean-markdown`** — formats Markdown files with paragraph indentation and blank-line rules
- **`clean-docx`** — cleans Word `.docx` files: removes artifacts, normalizes paragraphs, converts fonts
- **`format-docx`** — applies platform-specific layout (page size, margins, headers/footers) for KDP and print publishing *(in development)*

## Installation

```bash
pip install -e .
```

This installs the `clean-markdown`, `clean-docx`, and `format-docx` commands.

## Usage

```bash
# Format a Markdown file
clean-markdown -i input.md -o output.md

# Clean a .docx file (with config file — recommended)
clean-docx -c mybook.yaml

# Clean a .docx file (CLI only)
clean-docx -i input.docx -o output.docx --clean --page-breaks
```

## Documentation

- [clean-docx — full reference](docs/clean-docx.md)

## License

MIT — see [LICENSE](LICENSE).
