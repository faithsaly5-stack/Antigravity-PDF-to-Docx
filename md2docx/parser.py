"""
Markdown Parser and AST Builder.
Extracts YAML Frontmatter, GFM elements, LaTeX Math ($ and $$), Callouts, and Tables.
"""

import re
from typing import Dict, Any, List, Optional, Tuple
import yaml
import markdown_it
from .config import DocumentMetadata


CALLOUT_PREFIXES = {
    "[!NOTE]": ("note", "یادداشت", "Note"),
    "[!TIP]": ("tip", "نکته کاربردی", "Tip"),
    "[!IMPORTANT]": ("important", "مهم", "Important"),
    "[!WARNING]": ("warning", "هشدار", "Warning"),
    "[!CAUTION]": ("caution", "احتیاط", "Caution"),
}


def extract_frontmatter(text: str) -> Tuple[DocumentMetadata, str]:
    """
    Extracts YAML front matter from markdown text if present.
    Returns DocumentMetadata and remaining markdown body.
    """
    meta = DocumentMetadata()
    body = text

    # Check for YAML front matter at the start
    if text.startswith("---"):
        parts = text.split("---", 2)
        if len(parts) >= 3:
            yaml_content = parts[1].strip()
            body = parts[2].lstrip("\r\n")
            try:
                data = yaml.safe_load(yaml_content)
                if isinstance(data, dict):
                    meta.title = data.get("title")
                    meta.subtitle = data.get("subtitle")
                    
                    # Authors
                    authors_val = data.get("author") or data.get("authors") or []
                    if isinstance(authors_val, str):
                        meta.authors = [authors_val]
                    elif isinstance(authors_val, list):
                        meta.authors = [str(a) for a in authors_val]

                    # Supervisors
                    sup_val = data.get("supervisor") or data.get("supervisors") or []
                    if isinstance(sup_val, str):
                        meta.supervisors = [sup_val]
                    elif isinstance(sup_val, list):
                        meta.supervisors = [str(s) for s in sup_val]

                    meta.date = str(data.get("date")) if data.get("date") is not None else None
                    meta.institution = data.get("institution") or data.get("university")
                    meta.faculty = data.get("faculty")
                    meta.cover_page = bool(data.get("cover_page", False))
                    meta.toc = bool(data.get("toc", False))
                    meta.toc_title = data.get("toc_title", "فهرست مطالب")
                    meta.lang = data.get("lang", "fa")
                    meta.template_path = data.get("template")
                    meta.extra = data
            except Exception:
                pass

    return meta, body


class MarkdownParser:
    """
    Advanced Markdown parser with GFM, Math, and Persian Academic extensions.
    """
    def __init__(self):
        self.md = (
            markdown_it.MarkdownIt("commonmark", {"breaks": False, "html": True})
            .enable("table")
            .enable("strikethrough")
        )

    def preprocess_math(self, text: str) -> str:
        """
        Normalizes block math ($$...$$) so markdown-it handles it smoothly as code/math blocks.
        """
        # Replace display math $$ ... $$ with fenced math blocks ```math ... ```
        pattern = re.compile(r'^\$\$\s*\n(.*?)\n\s*\$\$', re.DOTALL | re.MULTILINE)
        
        def repl(match):
            eq = match.group(1).strip()
            return f"\n```math\n{eq}\n```\n"

        text = pattern.sub(repl, text)
        
        # Also handle inline $$eq$$ on their own line
        pattern_single_line = re.compile(r'^\$\$(.+?)\$\$$', re.MULTILINE)
        text = pattern_single_line.sub(r'\n```math\n\1\n```\n', text)
        return text

    def parse(self, text: str):
        """Parses markdown text into tokens."""
        clean_text = self.preprocess_math(text)
        tokens = self.md.parse(clean_text)
        return tokens
