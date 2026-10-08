"""
Perfect Markdown to DOCX Converter Engine.
Specialized for Persian Academic Word Document Engineering, Mixed-Script Typography, and Native Math.
"""

from .config import TypographyConfig, ColorTheme, DocumentMetadata
from .converter import MarkdownToDocx
from .word_com import update_toc_and_export_pdf

__version__ = "1.0.0"
__all__ = [
    "MarkdownToDocx",
    "TypographyConfig",
    "ColorTheme",
    "DocumentMetadata",
    "update_toc_and_export_pdf",
]
