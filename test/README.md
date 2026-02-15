# Test folder

Use this folder to try the md2docx converter.

**Single file:**

```bash
python convert.py --file test/sample.md
# Creates test/sample.docx
```

**Entire folder:**

```bash
python convert.py --folder test
# Converts every .md here to .docx alongside each file
```

**Folder with output to a directory:**

```bash
python convert.py --folder test --output-dir test/output
# Puts all .docx files in test/output/
```

Run from the project root (`MD_TO_DOC_convertor`).
