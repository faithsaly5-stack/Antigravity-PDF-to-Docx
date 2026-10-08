#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
High-Performance PDF to Image Cache Converter.
Converts long PDF documents into optimized 1050px JPEG images for vision-based
AI agent OCR pipelines without memory overhead.
"""

import os
import sys
import time
import shutil
import argparse
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


def convert_pdf_to_cache(pdf_path: str, output_dir: str = "page_cache",
                         target_width: int = 1050, quality: int = 70,
                         max_workers: int = 8, clean: bool = True) -> dict:
    """
    Converts PDF pages into optimized JPEGs concurrently.
    
    Args:
        pdf_path: Path to the target PDF file.
        output_dir: Directory where JPEG images will be stored.
        target_width: Target image width in pixels (1050px is the golden standard).
        quality: JPEG compression quality (70 gives optimal balance of speed & OCR sharpness).
        max_workers: Maximum threads for parallel processing.
        clean: Whether to clear existing output directory before converting.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    start_time = time.time()
    
    if clean and os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    
    print(f"\n[INFO] Processing: {os.path.basename(pdf_path)} ({total_pages} pages)")
    print(f"[INFO] Target Width: {target_width}px | Quality: {quality}% | Workers: {max_workers}")
    print(f"[INFO] Output Directory: {os.path.abspath(output_dir)}\n")

    # Determine padding width for filenames (e.g., page_001.jpg vs page_0001.jpg)
    pad_len = max(3, len(str(total_pages)))

    def process_page(page_num: int) -> int:
        """Worker function to render and save a single page."""
        # Open separate doc handle per thread for safe concurrency
        thread_doc = fitz.open(pdf_path)
        page = thread_doc[page_num]
        
        # Calculate scale factor to match target_width exactly
        rect_width = page.rect.width
        scale = target_width / rect_width if rect_width > 0 else 1.0
        matrix = fitz.Matrix(scale, scale)
        
        # Render pixmap (alpha=False ensures RGB without alpha channel for JPEG)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        
        # Save via PIL for optimized JPEG encoding
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        filename = f"page_{str(page_num + 1).zfill(pad_len)}.jpg"
        dest_path = os.path.join(output_dir, filename)
        img.save(dest_path, "JPEG", quality=quality, optimize=True)
        
        thread_doc.close()
        return os.path.getsize(dest_path)

    page_sizes = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_page, p) for p in range(total_pages)]
        for i, future in enumerate(futures, 1):
            size_bytes = future.result()
            page_sizes.append(size_bytes)
            if i % 25 == 0 or i == total_pages:
                print(f"  Processed {i}/{total_pages} pages ({(i/total_pages)*100:.1f}%)")

    doc.close()
    elapsed = time.time() - start_time
    total_size_mb = sum(page_sizes) / (1024 * 1024)
    avg_kb = (sum(page_sizes) / len(page_sizes)) / 1024 if page_sizes else 0

    stats = {
        "pdf_path": pdf_path,
        "total_pages": total_pages,
        "elapsed_sec": round(elapsed, 2),
        "total_size_mb": round(total_size_mb, 2),
        "avg_kb_per_page": round(avg_kb, 1),
        "output_dir": output_dir,
    }

    print("\n" + "=" * 50)
    print(" [SUCCESS] Image Cache Preparation Complete")
    print(f" Total Pages:        {total_pages}")
    print(f" Time Elapsed:       {elapsed:.2f} seconds ({total_pages/elapsed:.1f} pages/sec)")
    print(f" Total Folder Size:  {total_size_mb:.2f} MB")
    print(f" Average Size/Page:  {avg_kb:.1f} KB")
    print("=" * 50 + "\n")

    return stats


def main():
    parser = argparse.ArgumentParser(
        description="High-performance PDF page extractor for AI Agent OCR pipelines."
    )
    parser.add_argument("pdf", nargs="?", help="Path to source PDF file. If omitted, searches current dir.")
    parser.add_argument("-o", "--output", default="page_cache", help="Output directory (default: page_cache)")
    parser.add_argument("-w", "--width", type=int, default=1050, help="Target image width in pixels (default: 1050)")
    parser.add_argument("-q", "--quality", type=int, default=70, help="JPEG quality (default: 70)")
    parser.add_argument("--workers", type=int, default=8, help="Parallel worker threads (default: 8)")
    parser.add_argument("--no-clean", action="store_true", help="Do not delete existing output directory")

    args = parser.parse_args()

    target_pdf = args.pdf
    if not target_pdf:
        # Auto-discover PDFs in current directory
        candidates = [f for f in os.listdir(".") if f.lower().endswith(".pdf")]
        if not candidates:
            print("[ERROR] No PDF file specified and no .pdf files found in current directory.")
            sys.exit(1)
        elif len(candidates) == 1:
            target_pdf = candidates[0]
            print(f"[INFO] Auto-selected PDF: {target_pdf}")
        else:
            print("[INFO] Multiple PDF files found in directory:")
            for idx, c in enumerate(candidates, 1):
                print(f"  [{idx}] {c}")
            choice = input("Enter number to select (or filename): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(candidates):
                target_pdf = candidates[int(choice) - 1]
            elif os.path.exists(choice):
                target_pdf = choice
            else:
                print("[ERROR] Invalid selection.")
                sys.exit(1)

    convert_pdf_to_cache(
        pdf_path=target_pdf,
        output_dir=args.output,
        target_width=args.width,
        quality=args.quality,
        max_workers=args.workers,
        clean=not args.no_clean,
    )


if __name__ == "__main__":
    main()
