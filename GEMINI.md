# WritingUtils - Instructional Context

## Project Overview
WritingUtils is a collection of utilities designed to assist with writing-related tasks, specifically focused on formatting and cleaning up Markdown files for creative writing and novels.

- **Primary Language:** Python
- **Status:** Active development. The core `clean_markdown.py` utility is implemented, recently refined to accurately parse and preserve complex Markdown structures (like headers, TOC links, and scene metadata in novel manuscripts), and successfully tested against full-length novels.

## Building and Running
The primary utility is a standalone Python script.

- **Run Command:**
  ```bash
  python3 clean_markdown.py -i <input_file.md> -o <output_file.md>
  ```
- **Dependencies:** Uses Python standard library (`argparse`, `sys`, `re`). No external packages are currently required.
- **TODO:** Create a `pyproject.toml` or `requirements.txt` file if external dependencies are added in the future.

## Development Conventions
- **Language:** Python 3.x is expected.
- **Style:** Follow PEP 8 guidelines. The script should intelligently distinguish between structural Markdown elements (which should not be indented) and standard prose paragraphs (which should be indented with 4 spaces).
- **Testing:** Ensure formatting rules correctly preserve headers, images, links, blockquotes, code blocks, horizontal rules, and stylized scene metadata lines.

## Directory Structure
- `/`: Root directory containing project configuration and documentation.
- `/clean_markdown.py`: The main utility script for processing Markdown files.
- `/README.md`: Usage documentation.
- `/test_sample.md`, `/output_sample.md`: Basic test files for the utility.
- (Future) `tests/`: Intended location for formal test suites (e.g., `pytest`).
