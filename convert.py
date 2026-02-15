#!/usr/bin/env python3
"""
md2docx – Convert Markdown (.md) files to Word documents (.docx).

Supports:
  - Single file: convert one .md file to one .docx
  - Folder: convert all .md files in a directory (optionally recursive) to separate .docx files

Can be run from the command line or imported as a module for use in scripts and AI agents.

Usage (CLI):
  python convert.py --file path/to/file.md
  python convert.py --folder path/to/dir [--recursive] [--output-dir path/to/out]
"""

import argparse
import os
import sys
from pathlib import Path


def convert_md_to_docx(
    md_path: str | Path,
    docx_path: str | Path,
    reference_doc: str | Path | None = None,
) -> None:
    """
    Convert a single Markdown file to a DOCX file.

    Uses pypandoc (Pandoc) with markdown+hard_line_breaks for better
    paragraph and line-break handling in Word.

    Args:
        md_path: Path to the input .md file.
        docx_path: Path for the output .docx file (created/overwritten).
        reference_doc: Optional path to a reference .docx; its styles (fonts, headings,
                       margins) are applied to the output. Use for consistent formatting.

    Raises:
        FileNotFoundError: If md_path (or reference_doc) does not exist.
        RuntimeError: If conversion fails (e.g. pypandoc/Pandoc error).
    """
    md_path = Path(md_path).resolve()
    docx_path = Path(docx_path).resolve()

    if not md_path.is_file():
        raise FileNotFoundError(f"Markdown file not found: {md_path}")

    if reference_doc is not None:
        reference_doc = Path(reference_doc).resolve()
        if not reference_doc.is_file():
            raise FileNotFoundError(f"Reference document not found: {reference_doc}")

    # Ensure output directory exists
    docx_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        import pypandoc
    except ImportError:
        raise RuntimeError(
            "pypandoc is required. Install with: pip install pypandoc_binary"
        ) from None

    # Build extra args for Pandoc
    # Resolve images (e.g. ![alt](flowchart.png)) relative to the .md file's directory
    extra_args = [f"--resource-path={md_path.parent}"]
    if reference_doc is not None:
        extra_args.extend(["--reference-doc", str(reference_doc)])

    # Convert: markdown+hard_line_breaks improves line breaks in DOCX
    # Pass a list (never None); pypandoc extends internal args with extra_args.
    pypandoc.convert_file(
        str(md_path),
        "docx",
        outputfile=str(docx_path),
        format="markdown+hard_line_breaks",
        extra_args=extra_args,
    )


def collect_md_files(folder: Path, recursive: bool) -> list[Path]:
    """
    Collect all .md files under folder.

    Args:
        folder: Root directory to search.
        recursive: If True, include subdirectories; if False, only direct children.

    Returns:
        Sorted list of Paths to .md files.
    """
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"Not a directory: {folder}")

    if recursive:
        paths = list(folder.rglob("*.md"))
    else:
        paths = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".md"]

    return sorted(paths)


def run_single_file(
    md_path: Path,
    output: Path | None,
    reference_doc: Path | None = None,
) -> Path:
    """
    Convert one .md file to .docx.

    Output path is either the given output or same dir + same stem + .docx.
    Returns the path where the DOCX was written.
    """
    if output is not None:
        out_path = Path(output).resolve()
    else:
        out_path = md_path.with_suffix(".docx")

    convert_md_to_docx(md_path, out_path, reference_doc=reference_doc)
    return out_path


def run_folder(
    folder: Path,
    output_dir: Path | None,
    recursive: bool,
    reference_doc: Path | None,
) -> list[Path]:
    """
    Convert all .md files in folder to .docx.

    If output_dir is set, all DOCX files are written there (names from relative
    path to avoid collisions). Otherwise each DOCX is written next to its .md file.

    Returns list of output DOCX paths.
    """
    md_files = collect_md_files(folder, recursive)
    if not md_files:
        return []

    folder = Path(folder).resolve()
    out_paths = []

    for md_path in md_files:
        if output_dir is not None:
            # Preserve relative path under folder to avoid name clashes
            try:
                rel = md_path.relative_to(folder)
            except ValueError:
                rel = md_path.name
            # e.g. subdir/readme.md -> output_dir/subdir_readme.docx
            stem = str(rel.with_suffix("")).replace(os.sep, "_")
            docx_path = Path(output_dir).resolve() / f"{stem}.docx"
        else:
            docx_path = md_path.with_suffix(".docx")

        convert_md_to_docx(md_path, docx_path, reference_doc=reference_doc)
        out_paths.append(docx_path)

    return out_paths


def main() -> int:
    """
    CLI entrypoint: parse arguments and run single-file or folder conversion.
    """
    parser = argparse.ArgumentParser(
        description="Convert Markdown (.md) files to Word (.docx). Single file or whole folder.",
        epilog="Examples:\n"
        "  %(prog)s --file readme.md\n"
        "  %(prog)s --file readme.md --output report.docx\n"
        "  %(prog)s --folder ./docs\n"
        "  %(prog)s --folder ./docs --recursive --output-dir ./docx\n",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--file",
        "-f",
        type=Path,
        metavar="PATH",
        help="Path to a single .md file to convert",
    )
    group.add_argument(
        "--folder",
        "-d",
        type=Path,
        metavar="PATH",
        help="Path to a folder; convert all .md files inside (see --recursive)",
    )

    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        metavar="PATH",
        help="Output .docx path (single-file mode only). Default: same dir, same name.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        metavar="PATH",
        help="Output directory for folder mode. Default: same dir as each .md file.",
    )
    parser.add_argument(
        "--recursive",
        "-r",
        action="store_true",
        help="In folder mode, include .md files in subdirectories",
    )
    parser.add_argument(
        "--reference-doc",
        type=Path,
        metavar="PATH",
        help="Optional reference .docx for styles (fonts, headings, margins)",
    )

    args = parser.parse_args()

    try:
        if args.file is not None:
            out = run_single_file(
                args.file, args.output, reference_doc=args.reference_doc
            )
            print(f"Created: {out}")
        else:
            out_list = run_folder(
                args.folder,
                args.output_dir,
                args.recursive,
                args.reference_doc,
            )
            if not out_list:
                print("No .md files found.", file=sys.stderr)
                return 1
            for p in out_list:
                print(f"Created: {p}")
    except (FileNotFoundError, NotADirectoryError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
