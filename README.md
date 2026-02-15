# md2docx

Convert Markdown (`.md`) files to Word documents (`.docx`) from the command line or from your own scripts. Supports single files or entire folders, and is designed to work with AI-generated Markdown and automation (e.g. AI agents).

**Repository:** [github.com/RagingScout97/md2docx](https://github.com/RagingScout97/md2docx)

---

## What it does

- **Single file:** Convert one `.md` file to one `.docx` file.
- **Folder:** Convert all `.md` files in a directory to separate `.docx` files. Optionally run recursively (include subfolders) and write all outputs to one folder.
- **Styling:** Optional reference `.docx` so every output uses the same fonts, headings, and margins (great for AI-generated content).

Conversion uses [Pandoc](https://pandoc.org/) under the hood via [pypandoc_binary](https://pypi.org/project/pypandoc-binary/), so you get high-quality DOCX output without installing Pandoc yourself.

---

## Requirements

- **Python 3.10+** (for `str | Path` type hints; can be relaxed to 3.8+ if needed)

---

## Installation

1. Clone the repo (or download the code):

   ```bash
   git clone https://github.com/RagingScout97/md2docx.git
   cd md2docx
   ```

2. Create a virtual environment (recommended):

   ```bash
   python -m venv .venv
   .venv\Scripts\activate   # Windows
   # source .venv/bin/activate   # macOS/Linux
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

   This installs `pypandoc_binary`, which bundles Pandoc—no separate Pandoc install needed.

---

## Usage

### Single file

Convert one Markdown file. The output `.docx` is created next to the input by default (same name, `.docx` extension).

```bash
python convert.py --file path/to/readme.md
```

Specify the output path:

```bash
python convert.py --file readme.md --output report.docx
```

Short options:

```bash
python convert.py -f readme.md -o report.docx
```

### Folder (all `.md` files in a directory)

Convert every `.md` file in a folder. Each output `.docx` is placed next to its source file by default.

```bash
python convert.py --folder path/to/docs
```

Put all outputs in one directory (handles name clashes by using the relative path in the filename):

```bash
python convert.py --folder ./docs --output-dir ./docx
```

Include subfolders:

```bash
python convert.py --folder ./docs --recursive --output-dir ./docx
```

Short options:

```bash
python convert.py -d ./docs -r --output-dir ./docx
```

### Using a reference document (consistent formatting)

To apply the same styles (fonts, headings, margins) to every converted file, use a reference `.docx`:

```bash
python convert.py --file readme.md --reference-doc my-template.docx
python convert.py --folder ./docs --output-dir ./docx --reference-doc my-template.docx
```

You can create a base template with Pandoc once and reuse it:

```bash
pandoc -o reference.docx --print-default-data-file reference.docx
# Then edit reference.docx in Word (fonts, heading styles, margins) and use it as above.
```

---

## How it works (code flow)

1. **CLI** – `convert.py` uses `argparse` to handle `--file` or `--folder` (and options like `--output`, `--output-dir`, `--recursive`, `--reference-doc`).
2. **Single-file path** – If `--file` is set, the script resolves the output path (from `--output` or from the input path), then calls `convert_md_to_docx()` once.
3. **Folder path** – If `--folder` is set, it collects all `.md` paths (optionally with `--recursive`), then for each file calls `convert_md_to_docx()`, writing either next to the source or into `--output-dir` with a unique name.
4. **Conversion** – `convert_md_to_docx()` uses `pypandoc.convert_file()` with format `markdown+hard_line_breaks` and optional `--reference-doc` for styling.

You can also import and call the conversion function from your own Python code (see below).

---

## Use from scripts or AI agents

### As a subprocess

Run the script from any environment that has Python and the project dependencies:

```bash
python convert.py --file report.md
python convert.py --folder ./content --output-dir ./docx
```

### As a Python module

Import the converter and call it programmatically:

```python
from pathlib import Path
from convert import convert_md_to_docx

# Single file
convert_md_to_docx("readme.md", "readme.docx")

# With optional reference document for styling
convert_md_to_docx("readme.md", "readme.docx", reference_doc="template.docx")
```

This is useful in automation pipelines or when an AI agent runs Python code and needs to produce DOCX from Markdown.

---

## Project structure

```
md2docx/
├── convert.py          # CLI and conversion logic (single file + folder)
├── requirements.txt    # pypandoc_binary
├── README.md           # This file
└── LICENSE             # MIT
```

---

## License

This project is licensed under the **MIT License** – free to use, modify, and distribute. See [LICENSE](LICENSE) for the full text.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Contributing

Contributions are welcome. Open an issue or a pull request on [GitHub](https://github.com/RagingScout97/md2docx).
