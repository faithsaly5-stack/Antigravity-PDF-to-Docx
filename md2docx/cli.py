"""
Command Line Interface for Perfect Markdown to DOCX Converter.
"""

import sys
import os
import time
import argparse
import subprocess
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint

from .config import TypographyConfig, DocumentMetadata
from .converter import MarkdownToDocx

console = Console()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="Perfect Markdown to DOCX & PDF Converter Engine (Persian Academic & Mixed-Script Specialized)",
    )
    parser.add_argument("input", nargs="?", default=None, help="Path to input Markdown file (optional; opens GUI if omitted)")
    parser.add_argument("--gui", action="store_true", help="Launch graphical user interface (Windows GUI)")
    parser.add_argument("-o", "--output", help="Path to output .docx file (default: same name as input)")
    parser.add_argument("--pdf", action="store_true", help="Export to native PDF via Word COM engine")
    parser.add_argument("--template", help="Path to custom .docx master template")
    parser.add_argument("--academic", action="store_true", help="Use built-in academic master template with cover & TOC skeleton")
    parser.add_argument("--cover", action="store_true", help="Enable academic cover page")
    parser.add_argument("--toc", action="store_true", help="Insert Table of Contents")
    parser.add_argument("--font-fa", default="B Nazanin", help="Persian body font (default: 'B Nazanin')")
    parser.add_argument("--font-fa-heading", default="B Titr", help="Persian heading font (default: 'B Titr')")
    parser.add_argument("--font-en", default="Times New Roman", help="Latin body font (default: 'Times New Roman')")
    parser.add_argument("--margin", type=float, default=2.0, help="Page margins in cm (default: 2.0 cm)")
    parser.add_argument("--validate", action="store_true", help="Validate output DOCX using officecli validate")
    parser.add_argument("--watch", action="store_true", help="Watch input markdown file for changes and recompile")
    return parser


def run_conversion(args) -> bool:
    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        console.print(f"[bold red]Error:[/bold red] Input file not found: {input_path}")
        return False

    out_docx = args.output
    if not out_docx:
        base_no_ext = os.path.splitext(input_path)[0]
        out_docx = f"{base_no_ext}.docx"

    out_pdf = f"{os.path.splitext(out_docx)[0]}.pdf" if args.pdf else None

    # Resolve template
    template_path = args.template
    if args.academic and not template_path:
        builtin_tmpl = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "templates", "default_academic.docx"
        )
        if os.path.exists(builtin_tmpl):
            template_path = builtin_tmpl

    cfg = TypographyConfig(
        font_fa_body=args.font_fa,
        font_fa_heading=args.font_fa_heading,
        font_en_body=args.font_en,
        margin_top_cm=args.margin,
        margin_bottom_cm=args.margin,
        margin_left_cm=args.margin,
        margin_right_cm=args.margin,
    )

    converter = MarkdownToDocx(config=cfg)

    with console.status("[bold green]Compiling markdown to publication-grade DOCX...", spinner="dots"):
        t0 = time.time()
        res = converter.convert_file(
            input_file=input_path,
            output_docx=out_docx,
            output_pdf=out_pdf,
            template=template_path,
            cover_page=args.cover if args.cover else None,
            toc=args.toc if args.toc else None,
            export_pdf=args.pdf,
        )
        t_elapsed = time.time() - t0

    console.print(Panel.fit(
        f"[bold cyan]Perfect Markdown to DOCX Converter[/bold cyan]\n"
        f"[green]✔[/green] Document compiled successfully in [bold]{t_elapsed:.2f}s[/bold]!",
        border_style="green"
    ))

    # Summary Table
    table = Table(title="Conversion Summary", border_style="blue")
    table.add_column("Property", style="cyan", no_wrap=True)
    table.add_column("Value", style="magenta")

    table.add_row("Input File", input_path)
    table.add_row("Output DOCX", res["docx_path"])
    if res["pdf_path"]:
        table.add_row("Output PDF", res["pdf_path"])
    
    stats = res["stats"]
    table.add_row("Headings", str(stats["headings"]))
    table.add_row("Paragraphs", str(stats["paragraphs"]))
    table.add_row("Tables", str(stats["tables"]))
    table.add_row("Code Blocks", str(stats["code_blocks"]))
    table.add_row("Math Equations", str(stats["math_blocks"]))
    table.add_row("Callout Cards", str(stats["callouts"]))
    table.add_row("Images", str(stats["images"]))
    console.print(table)

    # Optional validation with officecli
    if args.validate:
        console.print("[dim]Running OpenXML schema validation via officecli...[/dim]")
        try:
            val_res = subprocess.run(
                ["officecli", "validate", res["docx_path"]],
                capture_output=True,
                text=True,
                check=False
            )
            if val_res.returncode == 0:
                console.print("[bold green]✔ OfficeCLI Validation Passed: OpenXML Schema is 100% compliant.[/bold green]")
            else:
                console.print(f"[yellow]OfficeCLI Validation Note:[/yellow] {val_res.stdout or val_res.stderr}")
        except FileNotFoundError:
            console.print("[yellow]Notice:[/yellow] officecli is not installed or not in PATH. Skipping validation.")

    return True


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.gui or not args.input:
        from .gui import launch_gui
        launch_gui()
        sys.exit(0)

    if not args.watch:
        success = run_conversion(args)
        sys.exit(0 if success else 1)

    # Watch Mode
    input_file = os.path.abspath(args.input)
    if not os.path.exists(input_file):
        console.print(f"[bold red]Error:[/bold red] Input file {input_file} not found.")
        sys.exit(1)

    console.print(f"[bold blue]Watching {input_file} for changes... (Press Ctrl+C to exit)[/bold blue]")
    last_mtime = os.path.getmtime(input_file)
    run_conversion(args)

    try:
        while True:
            time.sleep(1.0)
            cur_mtime = os.path.getmtime(input_file)
            if cur_mtime != last_mtime:
                last_mtime = cur_mtime
                console.print(f"\n[yellow]Change detected in {input_file}. Recompiling...[/yellow]")
                run_conversion(args)
    except KeyboardInterrupt:
        console.print("\n[dim]Stopped watching.[/dim]")


if __name__ == "__main__":
    main()
