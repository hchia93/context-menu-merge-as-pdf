import re
import os
import sys
import traceback
from pathlib import Path
from typing import List, Union
from datetime import datetime
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter

# Setup logging
LOG_FILE = Path(__file__).parent / "merge-pdf-debug.log"

def log(message):
    """Write log message to file and print to console"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_msg = f"[{timestamp}] {message}"
    print(log_msg)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_msg + "\n")
    except:
        pass

# Log the arguments
log(f"sys.argv: {sys.argv}")


def parse_page_spec(spec: str) -> List[int]:
    if not spec.startswith('[') or not spec.endswith(']'):
        raise ValueError("Invalid page specification format.")
    inner = spec[1:-1]
    pages = []
    for part in inner.split(','):
        part = part.strip()
        if ':' in part:
            start, end = map(int, part.split(':'))
            pages.extend(range(start - 1, end))
        else:
            pages.append(int(part) - 1)
    return pages


def add_pdf(writer: PdfWriter, filepath: Path, pages: Union[List[int], None]):
    try:
        reader = PdfReader(str(filepath))
        total = len(reader.pages)
        if pages is None:
            pages = list(range(total))
        for page_num in pages:
            if 0 <= page_num < total:
                writer.add_page(reader.pages[page_num])
            else:
                log(f"Skipping invalid page {page_num+1} in {filepath.name}")
    except Exception as e:
        log(f"Error adding PDF {filepath}: {e}")
        raise


def add_image(writer: PdfWriter, filepath: Path):
    try:
        image = Image.open(filepath).convert("RGB")
        tmp_path = filepath.with_suffix(".temp.pdf")
        image.save(tmp_path, "PDF", resolution=100.0)
        reader = PdfReader(str(tmp_path))
        writer.add_page(reader.pages[0])
        tmp_path.unlink(missing_ok=True)
    except Exception as e:
        log(f"Error adding image {filepath}: {e}")
        raise


def handle_merge_mode(inputs: List[str], output_path: str):
    writer = PdfWriter()
    for item in inputs:
        match = re.match(r'^(.+\.(pdf|PDF))\s*(\[.*\])?$', item)
        pages = None

        if match:
            filepath = Path(match.group(1)).expanduser().resolve()
            if match.group(3):
                try:
                    pages = parse_page_spec(match.group(3))
                except Exception as e:
                    log(f"Page spec error in {item}: {e}")
                    continue
        else:
            filepath = Path(item).expanduser().resolve()

        if not filepath.exists():
            log(f"File not found: {filepath}")
            continue

        if filepath.suffix.lower() in ['.png', '.jpg', '.jpeg']:
            add_image(writer, filepath)
        elif filepath.suffix.lower() == '.pdf':
            add_pdf(writer, filepath, pages)
        else:
            log(f"Unsupported file type: {filepath}")

    with open(output_path, 'wb') as f:
        writer.write(f)
    log(f"Merged PDF written to: {output_path}")


def handle_blob_mode(blob_dir: str, output_path: str):
    writer = PdfWriter()
    dir_path = Path(blob_dir).expanduser().resolve()
    if not dir_path.is_dir():
        log(f"Directory not found: {blob_dir}")
        return

    files = sorted(dir_path.glob("*"))
    for file in files:
        if file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
            add_image(writer, file)
        elif file.suffix.lower() == '.pdf':
            add_pdf(writer, file, None)

    with open(output_path, 'wb') as f:
        writer.write(f)
    log(f"Blob PDF created from folder: {output_path}")


def handle_auto_mode(inputs: List[str]):
    """
    Auto mode: intelligently handle mixed files and folders.
    """
    try:
        log(f"=== Starting auto mode with inputs: {inputs}")
        
        if not inputs:
            log("No inputs provided.")
            input("Press Enter to exit...")
            return
        
        paths = [Path(p).expanduser().resolve() for p in inputs]
        files, dirs = [], []

        for p in paths:
            log(f"Processing path: {p} (exists: {p.exists()}, is_file: {p.is_file()}, is_dir: {p.is_dir()})")
            if p.is_file() and p.suffix.lower() in ['.png', '.jpg', '.jpeg', '.pdf']:
                files.append(p)
                log(f"  -> Added as file")
            elif p.is_dir():
                dirs.append(p)
                log(f"  -> Added as directory")
            else:
                log(f"Skipping unsupported item: {p}")

        log(f"Found {len(files)} files and {len(dirs)} directories")

        if not files and not dirs:
            log("No valid inputs found.")
            input("Press Enter to exit...")
            return

        writer = PdfWriter()
        page_count = 0

        # Add individual files
        for f in sorted(files):
            log(f"Adding file: {f.name}")
            if f.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                add_image(writer, f)
                page_count += 1
            elif f.suffix.lower() == '.pdf':
                reader = PdfReader(str(f))
                pages = len(reader.pages)
                add_pdf(writer, f, None)
                page_count += pages
                log(f"  -> Added {pages} pages from PDF")

        # Add contents from folders
        for d in dirs:
            log(f"Adding contents from folder: {d.name}")
            folder_files = sorted(d.glob("*"))
            for file in folder_files:
                if file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                    log(f"  - {file.name}")
                    add_image(writer, file)
                    page_count += 1
                elif file.suffix.lower() == '.pdf':
                    log(f"  - {file.name}")
                    reader = PdfReader(str(file))
                    pages = len(reader.pages)
                    add_pdf(writer, file, None)
                    page_count += pages

        log(f"Total pages to merge: {page_count}")

        if page_count == 0:
            log("No valid PDF/image content found to merge.")
            input("Press Enter to exit...")
            return

        # Decide output directory
        if files:
            # Files selected: output in same directory as first file
            base_dir = files[0].parent
        else:
            # Only folders selected: output in parent directory of first folder
            base_dir = dirs[0].parent

        log(f"Output directory: {base_dir}")

        # Generate unique output filename
        output_path = base_dir / "merged.pdf"
        counter = 1
        while output_path.exists():
            output_path = base_dir / f"merged_{counter}.pdf"
            counter += 1

        log(f"Writing to: {output_path}")
        
        with open(output_path, 'wb') as f:
            writer.write(f)

        log(f"Successfully merged {page_count} pages into: {output_path}")
        
        # Open the generated PDF
        try:
            os.startfile(str(output_path))
            log("PDF opened successfully")
        except Exception as e:
            log(f"Could not open PDF automatically: {e}")
        
        # Keep window open for 2 seconds to show success message
        import time
        time.sleep(2)
        
    except Exception as e:
        log(f"FATAL ERROR in handle_auto_mode: {e}")
        log(f"Traceback: {traceback.format_exc()}")
        input("Press Enter to exit...")


def main():
    try:
        log("=== merge-pdf.py started ===")
        
        import argparse
        parser = argparse.ArgumentParser(description="Merge selective PDF pages and images into a single PDF.")
        group = parser.add_mutually_exclusive_group(required=False)
        group.add_argument('-merge', nargs='+', help='Input files with optional page spec, e.g. "/abs/path/file.pdf [1,3:5]"')
        group.add_argument('-blob', help='Merge all PDF/image files in a folder (alphabetical order)')
        parser.add_argument('-auto', nargs='+', help='Auto-detect mode based on input paths (for context menu)')
        parser.add_argument('-o', '--output', help='Output PDF path')
        args = parser.parse_args()

        log(f"Parsed args: {args}")

        if args.merge:
            if not args.output:
                log("Output path required for -merge mode.")
                return
            handle_merge_mode(args.merge, args.output)
        elif args.blob:
            if not args.output:
                log("Output path required for -blob mode.")
                return
            handle_blob_mode(args.blob, args.output)
        elif args.auto:
            handle_auto_mode(args.auto)
        else:
            parser.print_help()
            
    except Exception as e:
        log(f"FATAL ERROR in main: {e}")
        log(f"Traceback: {traceback.format_exc()}")
        input("Press Enter to exit...")


if __name__ == "__main__":
    main()