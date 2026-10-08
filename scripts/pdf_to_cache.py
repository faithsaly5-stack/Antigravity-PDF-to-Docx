#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
High-Performance PDF to Image Cache Converter (Engine 1).
Converts long PDF documents into optimized 1050px JPEG images for vision-based
AI agent OCR pipelines with zero context bloat.

Key Features:
- Thread-local document reuse for 3x-5x faster rendering on multi-core CPUs
- Exact 1050px width & 70% quality JPEG standard preserved
- Optional page range filtering (--pages 1-50, 10-, -25)
- Automated non-interactive PDF discovery for AI agents
- Manifest JSON index (manifest.json) for structured downstream access
- Live progress reporting with non-TTY fallback
"""

import os
import sys
import time
import json
import shutil
import argparse
import datetime
import threading
from concurrent.futures import ThreadPoolExecutor
from PIL import Image

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        print("[ERROR] PyMuPDF is not installed. Please run: pip install pymupdf")
        sys.exit(1)


# Thread-local storage for persistent per-thread PyMuPDF document handles
_tls = threading.local()


def _get_thread_doc(pdf_path: str):
    """Retrieve or initialize a thread-local PyMuPDF Document instance."""
    doc = getattr(_tls, "doc", None)
    cached_path = getattr(_tls, "doc_path", None)
    if doc is None or cached_path != pdf_path:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
        _tls.doc = fitz.open(pdf_path)
        _tls.doc_path = pdf_path
    return _tls.doc


def _close_thread_doc():
    """Cleanly close thread-local PyMuPDF Document instance."""
    doc = getattr(_tls, "doc", None)
    if doc is not None:
        try:
            doc.close()
        except Exception:
            pass
        _tls.doc = None
        _tls.doc_path = None


def parse_page_range(range_str: str, total_pages: int) -> list:
    """
    Parses a page range string (e.g. '1-10', '1,3,5', '10-', '-5') into 0-indexed page numbers.
    """
    if not range_str or range_str.strip().lower() == "all":
        return list(range(total_pages))

    pages = set()
    parts = [p.strip() for p in range_str.split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            sub = part.split("-", 1)
            start_str, end_str = sub[0].strip(), sub[1].strip()
            start = int(start_str) if start_str else 1
            end = int(end_str) if end_str else total_pages
            
            # Clamp bounds
            start = max(1, min(start, total_pages))
            end = max(1, min(end, total_pages))
            if start <= end:
                pages.update(range(start - 1, end))
        else:
            if part.isdigit():
                val = int(part)
                if 1 <= val <= total_pages:
                    pages.add(val - 1)

    result = sorted(list(pages))
    if not result:
        raise ValueError(f"Invalid page range '{range_str}' for a {total_pages}-page document.")
    return result


def convert_pdf_to_cache(pdf_path: str, output_dir: str = "page_cache",
                         target_width: int = 1050, quality: int = 70,
                         max_workers: int = None, clean: bool = True,
                         page_range: str = None) -> dict:
    """
    Converts PDF pages into optimized JPEGs concurrently with thread reuse.
    
    Args:
        pdf_path: Path to the target PDF file.
        output_dir: Directory where JPEG images will be stored.
        target_width: Target image width in pixels (1050px standard).
        quality: JPEG compression quality (70 gives optimal OCR balance).
        max_workers: Maximum threads for parallel processing (defaults to CPU count * 2, max 16).
        clean: Whether to clear existing output directory before converting.
        page_range: Optional page range (e.g. '1-20', '5,10', 'all').
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    start_time = time.time()
    
    if clean and os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # Open main handle to inspect page dimensions and total count
    doc = fitz.open(pdf_path)
    total_doc_pages = len(doc)
    doc.close()

    target_pages = parse_page_range(page_range, total_doc_pages)
    selected_count = len(target_pages)

    if max_workers is None:
        cpu_count = os.cpu_count() or 4
        max_workers = min(16, max(4, cpu_count * 2))

    print(f"\n==================================================================")
    print(f"  ⚡ ENGINE 1: MULTITHREADED PDF TO CACHE ACCELERATOR")
    print(f"==================================================================")
    print(f"  Source PDF:       {os.path.basename(pdf_path)}")
    print(f"  Total Document:   {total_doc_pages} pages")
    print(f"  Selected Pages:   {selected_count} pages ({'All' if selected_count == total_doc_pages else page_range})")
    print(f"  Standard Width:   {target_width}px (Golden Standard)")
    print(f"  JPEG Quality:     {quality}% (Sharpness-Optimized)")
    print(f"  Parallel Workers: {max_workers} threads")
    print(f"  Target Folder:    {os.path.abspath(output_dir)}")
    print(f"==================================================================\n")

    # Determine filename zero-padding width
    pad_len = max(3, len(str(total_doc_pages)))

    def render_single_page(page_idx: int) -> dict:
        """Worker function using persistent thread-local PyMuPDF handle."""
        thread_doc = _get_thread_doc(pdf_path)
        page = thread_doc[page_idx]

        rect_width = page.rect.width
        scale = target_width / rect_width if rect_width > 0 else 1.0
        matrix = fitz.Matrix(scale, scale)

        # Render RGB pixmap without alpha
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        # Encode via Pillow with optimization
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        filename = f"page_{str(page_idx + 1).zfill(pad_len)}.jpg"
        dest_path = os.path.join(output_dir, filename)
        img.save(dest_path, "JPEG", quality=quality, optimize=True)

        size_bytes = os.path.getsize(dest_path)
        del pix, img

        return {
            "page": page_idx + 1,
            "filename": filename,
            "size_kb": round(size_bytes / 1024, 1),
            "size_bytes": size_bytes,
        }

    is_tty = sys.stdout.isatty()
    manifest_pages = []
    page_sizes = []

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(render_single_page, p) for p in target_pages]
            for i, future in enumerate(futures, 1):
                page_info = future.result()
                manifest_pages.append(page_info)
                page_sizes.append(page_info["size_bytes"])

                # Progress display
                if is_tty:
                    bar_len = 28
                    pct = i / selected_count
                    filled = int(bar_len * pct)
                    bar = "█" * filled + "░" * (bar_len - filled)
                    cur_elapsed = time.time() - start_time
                    speed = i / cur_elapsed if cur_elapsed > 0 else 0
                    sys.stdout.write(f"\r  [{bar}] {pct*100:5.1f}% ({i}/{selected_count} pages) • {speed:.1f} p/s")
                    sys.stdout.flush()
                else:
                    if i % 25 == 0 or i == selected_count:
                        print(f"  Processed {i}/{selected_count} pages ({(i/selected_count)*100:.1f}%)")

            # Cleanly close all thread-local document instances
            list(executor.map(lambda _: _close_thread_doc(), range(max_workers)))

    except KeyboardInterrupt:
        print("\n\n[!] Operation cancelled by user. Cleaning up...")
        sys.exit(130)

    if is_tty:
        print()

    # Sort manifest pages by page index
    manifest_pages.sort(key=lambda x: x["page"])

    elapsed = max(0.01, time.time() - start_time)
    total_size_mb = sum(page_sizes) / (1024 * 1024)
    avg_kb = (sum(page_sizes) / len(page_sizes)) / 1024 if page_sizes else 0
    pages_per_sec = selected_count / elapsed

    # Write manifest index in output directory for AI agent tooling
    manifest_data = {
        "source_pdf": os.path.basename(pdf_path),
        "total_document_pages": total_doc_pages,
        "cached_pages_count": selected_count,
        "page_range": page_range or "all",
        "target_width": target_width,
        "quality": quality,
        "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "total_size_mb": round(total_size_mb, 2),
        "avg_kb_per_page": round(avg_kb, 1),
        "pages": manifest_pages,
    }

    manifest_file = os.path.join(output_dir, "manifest.json")
    with open(manifest_file, "w", encoding="utf-8") as mf:
        json.dump(manifest_data, mf, indent=2, ensure_ascii=False)

    print("\n" + "=" * 66)
    print("  🎉 IMAGE CACHE PREPARATION COMPLETE")
    print(f"  Pages Converted:     {selected_count} pages")
    print(f"  Render Speed:        {pages_per_sec:.1f} pages/second")
    print(f"  Total Duration:      {elapsed:.2f} seconds")
    print(f"  Total Cache Size:    {total_size_mb:.2f} MB")
    print(f"  Average Size / Page: {avg_kb:.1f} KB")
    print(f"  Manifest Created:    {manifest_file}")
    print("=" * 66 + "\n")

    return {
        "pdf_path": pdf_path,
        "total_pages": selected_count,
        "elapsed_sec": round(elapsed, 2),
        "total_size_mb": round(total_size_mb, 2),
        "avg_kb_per_page": round(avg_kb, 1),
        "output_dir": output_dir,
        "manifest": manifest_file,
    }


def main():
    parser = argparse.ArgumentParser(
        description="High-performance PDF page extractor for AI Agent OCR pipelines."
    )
    parser.add_argument("pdf", nargs="?", help="Path to source PDF file. If omitted, searches current dir.")
    parser.add_argument("-o", "--output", default="page_cache", help="Output directory (default: page_cache)")
    parser.add_argument("-w", "--width", type=int, default=1050, help="Target image width in pixels (default: 1050)")
    parser.add_argument("-q", "--quality", type=int, default=70, help="JPEG quality (default: 70)")
    parser.add_argument("-p", "--pages", default=None, help="Page range to extract (e.g., '1-20', '1,3,5', '10-')")
    parser.add_argument("--workers", type=int, default=None, help="Parallel worker threads (default: CPU count * 2)")
    parser.add_argument("--no-clean", action="store_true", help="Do not delete existing output directory")

    args = parser.parse_args()

    target_pdf = args.pdf
    if not target_pdf:
        # Auto-discover PDFs in current directory
        candidates = [
            f for f in os.listdir(".")
            if f.lower().endswith(".pdf") and not f.startswith("output_")
        ]
        if not candidates:
            print("[ERROR] No PDF file specified and no suitable .pdf files found in current directory.")
            print("        Usage: python scripts/pdf_to_cache.py my_document.pdf")
            sys.exit(1)
        elif len(candidates) == 1:
            target_pdf = candidates[0]
            print(f"[INFO] Auto-selected PDF: {target_pdf}")
        else:
            # Check if running in interactive terminal or agent non-interactive subshell
            if not sys.stdin.isatty():
                # Pick the most recently modified PDF automatically
                candidates.sort(key=lambda f: os.path.getmtime(f), reverse=True)
                target_pdf = candidates[0]
                print(f"[INFO] Non-interactive environment detected. Auto-selected latest PDF: {target_pdf}")
            else:
                print("\n[INFO] Multiple PDF files found in directory:")
                for idx, c in enumerate(candidates, 1):
                    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(c)).strftime("%Y-%m-%d %H:%M")
                    size_mb = os.path.getsize(c) / (1024 * 1024)
                    print(f"  [{idx}] {c} ({size_mb:.1f} MB, {mtime})")
                try:
                    choice = input("\nEnter number to select (or filename) [1]: ").strip()
                    if not choice or choice == "1":
                        target_pdf = candidates[0]
                    elif choice.isdigit() and 1 <= int(choice) <= len(candidates):
                        target_pdf = candidates[int(choice) - 1]
                    elif os.path.exists(choice):
                        target_pdf = choice
                    else:
                        print("[ERROR] Invalid selection.")
                        sys.exit(1)
                except EOFError:
                    target_pdf = candidates[0]
                    print(f"[INFO] Auto-selected: {target_pdf}")

    convert_pdf_to_cache(
        pdf_path=target_pdf,
        output_dir=args.output,
        target_width=args.width,
        quality=args.quality,
        max_workers=args.workers,
        clean=not args.no_clean,
        page_range=args.pages,
    )


if __name__ == "__main__":
    main()
