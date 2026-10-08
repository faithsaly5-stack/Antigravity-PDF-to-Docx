"""
Main Converter Engine: Markdown to High-Quality Academic DOCX / PDF Converter.
"""

import os
import re
from typing import Optional, Dict, Any, List
from .config import DocumentMetadata, TypographyConfig, ColorTheme
from .parser import extract_frontmatter, MarkdownParser, CALLOUT_PREFIXES
from .docx_builder import DocxBuilder
from .math_parser import insert_display_equation
from .code_highlighter import insert_code_block
from .cover_engine import insert_standalone_toc
from .bidi import has_persian_chars
from .word_com import update_toc_and_export_pdf


class MarkdownToDocx:
    def __init__(
        self,
        config: Optional[TypographyConfig] = None,
        theme: Optional[ColorTheme] = None,
        default_template: Optional[str] = None,
    ):
        self.config = config or TypographyConfig()
        self.theme = theme or ColorTheme()
        self.default_template = default_template
        self.parser = MarkdownParser()

    def convert_file(
        self,
        input_file: str,
        output_docx: Optional[str] = None,
        output_pdf: Optional[str] = None,
        template: Optional[str] = None,
        cover_page: Optional[bool] = None,
        toc: Optional[bool] = None,
        export_pdf: bool = False,
        update_toc: bool = True,
    ) -> Dict[str, Any]:
        """
        Converts a markdown file to a publication-ready Word (.docx) document
        and optionally exports to a native PDF.
        """
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"Input markdown file not found: {input_file}")

        base_dir = os.path.dirname(os.path.abspath(input_file))
        with open(input_file, "r", encoding="utf-8") as f:
            content = f.read()

        if output_docx is None:
            base_name = os.path.splitext(input_file)[0]
            output_docx = f"{base_name}.docx"

        if export_pdf and output_pdf is None:
            base_name = os.path.splitext(output_docx)[0]
            output_pdf = f"{base_name}.pdf"

        return self.convert_string(
            markdown_text=content,
            output_docx=output_docx,
            output_pdf=output_pdf,
            base_dir=base_dir,
            template=template,
            cover_page=cover_page,
            toc=toc,
            export_pdf=export_pdf,
            update_toc=update_toc,
        )

    def convert_string(
        self,
        markdown_text: str,
        output_docx: str,
        output_pdf: Optional[str] = None,
        base_dir: str = ".",
        template: Optional[str] = None,
        cover_page: Optional[bool] = None,
        toc: Optional[bool] = None,
        export_pdf: bool = False,
        update_toc: bool = True,
    ) -> Dict[str, Any]:
        """Converts raw markdown string to Word document."""
        meta, body_text = extract_frontmatter(markdown_text)

        # Apply CLI overrides if explicitly passed
        if template is not None:
            meta.template_path = template
        elif self.default_template is not None and not meta.template_path:
            meta.template_path = self.default_template

        if cover_page is not None:
            meta.cover_page = cover_page
        if toc is not None:
            meta.toc = toc

        # Parse AST tokens
        tokens = self.parser.parse(body_text)

        builder = DocxBuilder(
            metadata=meta,
            config=self.config,
            theme=self.theme,
            base_dir=base_dir,
        )
        builder.initialize_document()

        # Stats tracking
        stats = {
            "headings": 0,
            "paragraphs": 0,
            "tables": 0,
            "code_blocks": 0,
            "math_blocks": 0,
            "callouts": 0,
            "lists": 0,
            "images": 0,
        }

        # Process AST tokens
        i = 0
        n_tokens = len(tokens)
        while i < n_tokens:
            tok = tokens[i]
            ttype = tok.type

            # --- Headings ---
            if ttype == 'heading_open':
                level = int(tok.tag[1:])
                inline_tok = tokens[i + 1]
                builder.add_heading(inline_tok.content, level)
                stats["headings"] += 1
                i += 3
                continue

            # --- Paragraphs ---
            elif ttype == 'paragraph_open':
                inline_tok = tokens[i + 1]
                
                # Check for explicit PageBreak
                raw_text = inline_tok.content.strip()
                if raw_text in (r'\pagebreak', '<!-- pagebreak -->', '<!-- page_break -->'):
                    builder.add_page_break()
                    i += 3
                    continue

                # Check for explicit TOC placeholder
                if raw_text.lower() in ('[toc]', '[[toc]]', '<!-- toc -->'):
                    insert_standalone_toc(builder.doc, meta.toc_title, self.config)
                    i += 3
                    continue

                # Check if paragraph contains an Image
                if inline_tok.children and any(c.type == 'image' for c in inline_tok.children):
                    for c in inline_tok.children:
                        if c.type == 'image':
                            src = c.attrs.get('src', '') if c.attrs else ''
                            caption = c.content or (c.children[0].content if c.children else "")
                            builder.add_image_with_caption(src, caption)
                            stats["images"] += 1
                    i += 3
                    continue

                # Normal Paragraph
                builder.add_paragraph(inline_tok.children or [inline_tok], raw_text=inline_tok.content)
                stats["paragraphs"] += 1
                i += 3
                continue

            # --- Fenced Code Blocks & Math ---
            elif ttype == 'fence':
                lang = (tok.info or "").strip().lower()
                if lang in ('math', 'latex'):
                    insert_display_equation(
                        builder.doc,
                        tok.content,
                        space_before_pt=self.config.space_before_equation,
                        space_after_pt=self.config.space_after_equation,
                    )
                    stats["math_blocks"] += 1
                else:
                    insert_code_block(
                        builder.doc,
                        tok.content,
                        lang=lang,
                        font_name=self.config.font_code,
                        font_size_pt=self.config.size_code,
                    )
                    stats["code_blocks"] += 1
                i += 1
                continue

            elif ttype == 'code_block':
                insert_code_block(
                    builder.doc,
                    tok.content,
                    font_name=self.config.font_code,
                    font_size_pt=self.config.size_code,
                )
                stats["code_blocks"] += 1
                i += 1
                continue

            # --- Tables ---
            elif ttype == 'table_open':
                tbl_res, consumed = self._parse_table(tokens, i)
                headers, rows, alignments = tbl_res
                builder.add_table(headers, rows, alignments)
                stats["tables"] += 1
                i += consumed
                continue

            # --- Blockquotes & Callouts ---
            elif ttype == 'blockquote_open':
                bq_res, consumed = self._parse_blockquote(tokens, i)
                callout_type, title, body_items, inner_tables, is_rtl = bq_res
                if body_items:
                    builder.add_callout_box(callout_type, title, body_items, is_rtl=is_rtl)
                    stats["callouts"] += 1
                for tbl_h, tbl_r, tbl_a in inner_tables:
                    builder.add_table(tbl_h, tbl_r, tbl_a)
                    stats["tables"] += 1
                i += consumed
                continue

            # --- HTML Block (Page Breaks & Markers) ---
            elif ttype == 'html_block':
                content = tok.content
                if 'START OF PAGE:' in content:
                    page_m = re.search(r'START OF PAGE:\s*page_(\d+)\.jpg', content)
                    if page_m:
                        page_num = int(page_m.group(1))
                        if page_num > 1:
                            builder.add_page_break()
                            stats["page_breaks"] = stats.get("page_breaks", 0) + 1
                    else:
                        builder.add_page_break()
                        stats["page_breaks"] = stats.get("page_breaks", 0) + 1
                i += 1
                continue

            # --- Lists ---
            elif ttype in ('bullet_list_open', 'ordered_list_open'):
                consumed = self._parse_list(tokens, i, builder)
                stats["lists"] += 1
                i += consumed
                continue

            # --- Horizontal Rule ---
            elif ttype == 'hr':
                builder.add_horizontal_rule()
                i += 1
                continue

            i += 1

        # Finalize and enforce OpenXML schema compliance
        builder.finalize()

        # Save DOCX
        out_dir = os.path.dirname(os.path.abspath(output_docx))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        builder.doc.save(output_docx)

        # Word COM Automation (TOC update & PDF export)
        pdf_generated = False
        if update_toc or export_pdf:
            pdf_path = output_pdf if export_pdf else None
            try:
                pdf_generated = update_toc_and_export_pdf(output_docx, pdf_path)
            except Exception as e:
                print(f"[Word COM] Failed: {e}")

        return {
            "status": "success",
            "docx_path": os.path.abspath(output_docx),
            "pdf_path": os.path.abspath(output_pdf) if (export_pdf and pdf_generated) else None,
            "stats": stats,
            "metadata": meta,
        }

    def _parse_table(self, tokens: List[Any], start_idx: int):
        """Extracts headers, data rows, and alignments from markdown table tokens."""
        i = start_idx + 1
        headers: List[List[Any]] = []
        rows: List[List[List[Any]]] = []
        alignments: List[str] = []
        in_thead = False
        in_tbody = False
        current_row: List[List[Any]] = []

        while i < len(tokens):
            t = tokens[i]
            if t.type == 'table_close':
                i += 1
                break
            elif t.type == 'thead_open':
                in_thead = True
            elif t.type == 'thead_close':
                in_thead = False
            elif t.type == 'tbody_open':
                in_tbody = True
            elif t.type == 'tbody_close':
                in_tbody = False
            elif t.type == 'tr_open':
                current_row = []
            elif t.type == 'tr_close':
                if in_thead:
                    headers = current_row
                else:
                    rows.append(current_row)
            elif t.type in ('th_open', 'td_open'):
                align = 'center'
                if t.attrs and 'style' in t.attrs:
                    style_str = t.attrs['style']
                    if 'right' in style_str:
                        align = 'right'
                    elif 'left' in style_str:
                        align = 'left'
                if in_thead and len(alignments) < len(current_row) + 1:
                    alignments.append(align)
                
                # Next token is inline content
                inline_tok = tokens[i + 1]
                current_row.append(inline_tok.children or [inline_tok])
                i += 2  # Skip th/td_open and inline
                continue

            i += 1

        consumed = i - start_idx
        return (headers, rows, alignments), consumed

    def _parse_blockquote(self, tokens: List[Any], start_idx: int):
        """
        Extracts callout or quote content from blockquote tokens.
        Supports paragraphs, nested lists, and inner tables.
        """
        i = start_idx + 1
        body_items: List[Any] = []
        inner_tables: List[Any] = []
        list_stack: List[Dict[str, Any]] = []

        while i < len(tokens):
            t = tokens[i]
            if t.type == 'blockquote_close':
                i += 1
                break
            elif t.type == 'paragraph_open':
                inline_tok = tokens[i + 1]
                body_items.append(('paragraph', inline_tok.children or [inline_tok], inline_tok.content))
                i += 3
                continue
            elif t.type in ('bullet_list_open', 'ordered_list_open'):
                is_ordered = (t.type == 'ordered_list_open')
                list_stack.append({"is_ordered": is_ordered, "counter": 1})
                i += 1
            elif t.type in ('bullet_list_close', 'ordered_list_close'):
                if list_stack:
                    list_stack.pop()
                i += 1
            elif t.type == 'list_item_open':
                curr_level = max(0, len(list_stack) - 1)
                curr_info = list_stack[-1] if list_stack else {"is_ordered": False, "counter": 1}
                next_tok = tokens[i + 1]
                if next_tok.type == 'paragraph_open':
                    inline_tok = tokens[i + 2]
                    body_items.append(('list_item', inline_tok.children or [inline_tok], inline_tok.content, curr_info["is_ordered"], curr_info["counter"], curr_level))
                    curr_info["counter"] += 1
                    i += 4
                    continue
                elif next_tok.type == 'inline':
                    body_items.append(('list_item', next_tok.children or [next_tok], next_tok.content, curr_info["is_ordered"], curr_info["counter"], curr_level))
                    curr_info["counter"] += 1
                    i += 2
                    continue
                else:
                    i += 1
            elif t.type == 'list_item_close':
                i += 1
            elif t.type == 'table_open':
                tbl_res, tbl_consumed = self._parse_table(tokens, i)
                inner_tables.append(tbl_res)
                i += tbl_consumed
                continue
            else:
                i += 1

        consumed = i - start_idx
        callout_type = "quote"
        title = None
        is_rtl = True

        if body_items and body_items[0][0] == 'paragraph':
            first_text = body_items[0][2].strip()
            # Check for GitHub callout prefixes
            for prefix, (c_type, fa_title, en_title) in CALLOUT_PREFIXES.items():
                if first_text.startswith(prefix):
                    callout_type = c_type
                    clean_first = first_text[len(prefix):].strip()
                    is_rtl = has_persian_chars(clean_first) or True
                    title = fa_title if is_rtl else en_title
                    first_children = [
                        c for c in body_items[0][1]
                        if not (hasattr(c, 'content') and prefix in c.content)
                    ]
                    body_items[0] = ('paragraph', first_children or body_items[0][1], clean_first)
                    break

        return (callout_type, title, body_items, inner_tables, is_rtl), consumed

    def _parse_list(self, tokens: List[Any], start_idx: int, builder: DocxBuilder) -> int:
        """Parses nested ordered/unordered lists and passes them to DocxBuilder."""
        stack = []
        i = start_idx

        while i < len(tokens):
            t = tokens[i]
            if t.type in ('bullet_list_open', 'ordered_list_open'):
                is_ordered = (t.type == 'ordered_list_open')
                stack.append({"is_ordered": is_ordered, "counter": 1})
                i += 1
            elif t.type in ('bullet_list_close', 'ordered_list_close'):
                stack.pop()
                i += 1
                if not stack:
                    break
            elif t.type == 'list_item_open':
                curr_level = len(stack) - 1
                curr_info = stack[-1]
                # Look for paragraph / inline
                next_tok = tokens[i + 1]
                if next_tok.type == 'paragraph_open':
                    inline_tok = tokens[i + 2]
                    raw_text = inline_tok.content
                    builder.add_list_item(
                        inline_tokens=inline_tok.children or [inline_tok],
                        raw_text=raw_text,
                        is_ordered=curr_info["is_ordered"],
                        order_index=curr_info["counter"],
                        indent_level=curr_level,
                    )
                    curr_info["counter"] += 1
                    i += 4  # list_item_open, paragraph_open, inline, paragraph_close
                elif next_tok.type == 'inline':
                    raw_text = next_tok.content
                    builder.add_list_item(
                        inline_tokens=next_tok.children or [next_tok],
                        raw_text=raw_text,
                        is_ordered=curr_info["is_ordered"],
                        order_index=curr_info["counter"],
                        indent_level=curr_level,
                    )
                    curr_info["counter"] += 1
                    i += 2
                else:
                    i += 1
            elif t.type == 'list_item_close':
                i += 1
            else:
                i += 1

        return i - start_idx
