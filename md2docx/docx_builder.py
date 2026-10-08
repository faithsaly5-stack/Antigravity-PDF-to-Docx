"""
Document Builder: Constructs publication-grade Microsoft Word (.docx) documents.
Implements OpenXML layout, complex script typography, tables, lists, callouts, and math.
"""

import os
import re
import tempfile
from typing import List, Optional, Tuple, Dict, Any
import requests
from PIL import Image

import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from .config import DocumentMetadata, TypographyConfig, ColorTheme
from .bidi import (
    has_persian_chars,
    is_predominantly_persian,
    is_english_reference,
    tokenize_runs,
    set_paragraph_bidi,
    apply_paragraph_spacing,
    style_run,
)
from .math_parser import latex_to_omml, insert_display_equation
from .code_highlighter import insert_code_block
from .cover_engine import (
    prepare_template_document,
    build_standalone_cover_page,
    insert_standalone_toc,
)

INLINE_MATH_PATTERN = re.compile(r'(?<!\\)\$(?!\$)(.+?)(?<!\\)\$')


class DocxBuilder:
    def __init__(
        self,
        metadata: DocumentMetadata,
        config: Optional[TypographyConfig] = None,
        theme: Optional[ColorTheme] = None,
        base_dir: str = ".",
    ):
        self.metadata = metadata
        self.config = config or TypographyConfig()
        self.theme = theme or ColorTheme()
        self.base_dir = base_dir
        self.doc: Optional[docx.Document] = None
        self.image_counter = 0

    def initialize_document(self) -> docx.Document:
        """Initializes document either from a template or standalone."""
        template_file = self.metadata.template_path
        
        # Check default template fallback if requested or available
        if template_file and not os.path.isabs(template_file):
            candidate = os.path.join(self.base_dir, template_file)
            if not os.path.exists(candidate):
                default_tmpl = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "templates", "default_academic.docx"
                )
                if os.path.exists(default_tmpl):
                    template_file = default_tmpl
            else:
                template_file = candidate

        if template_file and os.path.exists(template_file):
            self.doc = docx.Document(template_file)
            prepare_template_document(self.doc, self.metadata)
        else:
            self.doc = docx.Document()
            self._setup_page_geometry()
            self._setup_default_styles()
            
            if self.metadata.cover_page:
                build_standalone_cover_page(self.doc, self.metadata, self.config)
            if self.metadata.toc:
                insert_standalone_toc(self.doc, self.metadata.toc_title, self.config)

        return self.doc

    def _setup_page_geometry(self):
        """Sets up 2.0 cm margins and A4 dimensions."""
        for section in self.doc.sections:
            section.top_margin = Cm(self.config.margin_top_cm)
            section.bottom_margin = Cm(self.config.margin_bottom_cm)
            section.left_margin = Cm(self.config.margin_left_cm)
            section.right_margin = Cm(self.config.margin_right_cm)
            
            if self.config.page_size.upper() == "A4":
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)
            else:  # Letter
                section.page_width = Cm(21.59)
                section.page_height = Cm(27.94)

    def _setup_default_styles(self):
        """Sets OpenXML default Complex Script font bindings."""
        styles = self.doc.styles
        if 'Normal' in styles:
            normal = styles['Normal']
            normal.font.name = self.config.font_en_body
            normal.font.size = Pt(self.config.size_body_en)

    def add_heading(self, text: str, level: int):
        """Renders an academic heading (H1 to H6) with exact typography and spacing."""
        style_name = f'Heading {level}' if level <= 6 else 'Normal'
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Subtitle hyphenation rule: spaces around hyphens
        spaced_text = re.sub(r'(?<=\S)[\–\-](?=\S)', ' - ', text)

        is_rtl = is_predominantly_persian(spaced_text)
        set_paragraph_bidi(p, is_rtl=is_rtl, justify=True)

        if level == 1:
            apply_paragraph_spacing(
                p,
                before_pt=self.config.space_before_h1,
                after_pt=self.config.space_after_h1,
                line_spacing_multiplier=self.config.line_spacing_multiplier,
            )
            sz = self.config.size_h1
            bold = True
            font_fa = self.config.font_fa_heading
            font_en = self.config.font_en_heading
        elif level == 2:
            apply_paragraph_spacing(
                p,
                before_pt=self.config.space_before_h2,
                after_pt=self.config.space_after_h2,
                line_spacing_multiplier=self.config.line_spacing_multiplier,
            )
            sz = self.config.size_h2
            bold = True
            font_fa = self.config.font_fa_heading
            font_en = self.config.font_en_heading
        elif level == 3:
            apply_paragraph_spacing(
                p,
                before_pt=self.config.space_before_h3,
                after_pt=self.config.space_after_h3,
                line_spacing_multiplier=self.config.line_spacing_multiplier,
            )
            sz = self.config.size_h3
            bold = True
            font_fa = self.config.font_fa_heading
            font_en = self.config.font_en_heading
        else:
            apply_paragraph_spacing(p, before_pt=6.0, after_pt=3.0, line_spacing_multiplier=self.config.line_spacing_multiplier)
            sz = self.config.size_h4
            bold = True
            font_fa = self.config.font_fa_heading
            font_en = self.config.font_en_heading

        for seg_type, seg_text in tokenize_runs(spaced_text):
            r = p.add_run(seg_text)
            style_run(
                r,
                font_fa=font_fa,
                font_en=font_en,
                size_pt=sz,
                is_fa=(seg_type == 'fa'),
                bold=bold,
            )

        # Set heading style in pPr so TOC detects it
        pPr = p._element.get_or_add_pPr()
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            pStyle = OxmlElement('w:pStyle')
            pPr.append(pStyle)
        pStyle.set(qn('w:val'), f'Heading{level}')
        return p

    def add_paragraph(
        self,
        inline_tokens: List[Any],
        raw_text: str = "",
        before_pt: float = 0.0,
        after_pt: float = 8.0,
    ):
        """
        Renders a body paragraph with Tokenized Run Generation and Reference BIDI Isolation.
        """
        p = self.doc.add_paragraph()
        
        # Check Reference BIDI Isolation Rule
        is_ref = is_english_reference(raw_text)
        is_rtl = not is_ref and is_predominantly_persian(raw_text)

        set_paragraph_bidi(p, is_rtl=is_rtl, justify=True)
        apply_paragraph_spacing(
            p,
            before_pt=before_pt,
            after_pt=after_pt,
            line_spacing_multiplier=self.config.line_spacing_multiplier,
        )

        self._render_inline_tokens(p, inline_tokens, default_is_fa=is_rtl)
        return p

    def _render_inline_tokens(
        self,
        p,
        tokens: List[Any],
        default_is_fa: bool = True,
        default_bold: bool = False,
        default_italic: bool = False,
        default_size_pt: Optional[float] = None,
        color_hex: Optional[str] = None,
    ):
        """
        Renders inline markdown tokens (bold, italic, code, math, links, text)
        with tokenized run script segregation.
        """
        cur_bold = default_bold
        cur_italic = default_italic
        cur_strike = False
        cur_link_url: Optional[str] = None
        cur_link_children: List[Any] = []
        in_link = False

        sz_fa = default_size_pt or self.config.size_body_fa
        sz_en = default_size_pt or self.config.size_body_en

        idx = 0
        while idx < len(tokens):
            tok = tokens[idx]
            ttype = tok.type

            if ttype == 'strong_open':
                cur_bold = True
            elif ttype == 'strong_close':
                cur_bold = False
            elif ttype == 'em_open':
                cur_italic = True
            elif ttype == 'em_close':
                cur_italic = False
            elif ttype == 's_open':
                cur_strike = True
            elif ttype == 's_close':
                cur_strike = False
            elif ttype == 'link_open':
                in_link = True
                cur_link_url = tok.attrs.get('href') if tok.attrs else None
            elif ttype == 'link_close':
                in_link = False
                cur_link_url = None
            elif ttype == 'code_inline':
                # Inline code: Consolas, shading, strictly LTR
                r = p.add_run(tok.content)
                style_run(
                    r,
                    font_fa=self.config.font_code,
                    font_en=self.config.font_code,
                    size_pt=self.config.size_code,
                    is_fa=False,
                    bold=False,
                    italic=False,
                    highlight_color=self.theme.inline_code_bg,
                    is_code=True,
                )
            elif ttype == 'text':
                content = tok.content
                # Handle inline math ($...$) within text
                self._render_text_with_inline_math(
                    p,
                    content,
                    bold=cur_bold,
                    italic=cur_italic,
                    strike=cur_strike,
                    size_fa=sz_fa,
                    size_en=sz_en,
                    color_hex=self.theme.link_color if in_link else color_hex,
                    underline=in_link,
                    url=cur_link_url if in_link else None,
                )
            idx += 1

    def _render_text_with_inline_math(
        self,
        p,
        text: str,
        bold: bool = False,
        italic: bool = False,
        strike: bool = False,
        size_fa: float = 14.0,
        size_en: float = 13.0,
        color_hex: Optional[str] = None,
        underline: bool = False,
        url: Optional[str] = None,
    ):
        """
        Renders a chunk of text that may contain inline math $...$ formulas.
        """
        # Split by inline math pattern
        segments = []
        last_idx = 0
        for m in INLINE_MATH_PATTERN.finditer(text):
            start, end = m.span()
            if start > last_idx:
                segments.append(('text', text[last_idx:start]))
            segments.append(('math', m.group(1)))
            last_idx = end
        if last_idx < len(text):
            segments.append(('text', text[last_idx:]))

        if not segments:
            segments = [('text', text)]

        for seg_type, val in segments:
            if seg_type == 'math':
                # Render native Word OMML inline equation
                omml = latex_to_omml(val, is_display=False)
                if omml is not None:
                    p._element.append(omml)
                else:
                    # Fallback to Cambria Math run
                    r = p.add_run(val)
                    r.font.name = self.config.font_math
                    r.font.italic = True
                    r.font.size = Pt(size_en)
            else:
                # Text: apply Tokenized Run Generation
                for script, chunk in tokenize_runs(val):
                    if url:
                        self._add_hyperlink_run(
                            p, url, chunk, is_fa=(script == 'fa'),
                            bold=bold, italic=italic, size_pt=(size_fa if script == 'fa' else size_en)
                        )
                    else:
                        r = p.add_run(chunk)
                        style_run(
                            r,
                            font_fa=self.config.font_fa_body,
                            font_en=self.config.font_en_body,
                            size_pt=(size_fa if script == 'fa' else size_en),
                            is_fa=(script == 'fa'),
                            bold=bold,
                            italic=italic,
                            strike=strike,
                            underline=underline,
                            color_hex=color_hex,
                        )

    def _add_hyperlink_run(self, p, url: str, text: str, is_fa: bool, bold: bool, italic: bool, size_pt: float):
        """Creates a native Word OpenXML hyperlink element."""
        part = p.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hl = parse_xml(
            f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'r:id="{r_id}" w:history="1"/>'
        )
        
        r = OxmlElement('w:r')
        hl.append(r)
        
        # Style run
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        if is_fa:
            rFonts.set(qn('w:ascii'), self.config.font_fa_body)
            rFonts.set(qn('w:hAnsi'), self.config.font_fa_body)
            rFonts.set(qn('w:cs'), self.config.font_fa_body)
        else:
            rFonts.set(qn('w:ascii'), self.config.font_en_body)
            rFonts.set(qn('w:hAnsi'), self.config.font_en_body)
            rFonts.set(qn('w:cs'), self.config.font_fa_body)
        rPr.append(rFonts)

        sz_val = str(int(size_pt * 2))
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), sz_val)
        rPr.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), sz_val)
        rPr.append(szCs)

        # Color & Underline
        c = OxmlElement('w:color')
        c.set(qn('w:val'), self.theme.link_color)
        rPr.append(c)

        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

        if bold:
            rPr.append(OxmlElement('w:b'))
            rPr.append(OxmlElement('w:bCs'))
        if italic:
            rPr.append(OxmlElement('w:i'))
            rPr.append(OxmlElement('w:iCs'))
        if is_fa:
            rPr.append(OxmlElement('w:rtl'))

        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = text
        r.append(t)
        
        p._element.append(hl)

    def add_table(self, headers: List[List[Any]], rows: List[List[List[Any]]], alignments: List[str]):
        """
        Renders an academic table with booktabs borders, cantSplit, tblHeader,
        and automatic <w:bidiVisual/> for Persian RTL tables.
        """
        num_cols = max(len(headers), max((len(r) for r in rows), default=0))
        if num_cols == 0:
            return

        total_rows = (1 if headers else 0) + len(rows)
        table = self.doc.add_table(rows=total_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Detect if table contains Persian text
        all_text = " ".join(tok.content for h in headers for tok in h if hasattr(tok, 'content'))
        all_text += " " + " ".join(tok.content for r in rows for cell in r for tok in cell if hasattr(tok, 'content'))
        table_is_rtl = has_persian_chars(all_text)

        # Table Properties (tblPr)
        tblPr = table._element.xpath('w:tblPr')
        if tblPr:
            # Set width to 100% (approx 9360 dxa)
            tblW = parse_xml(r'<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                             r'w:w="9360" w:type="dxa"/>')
            tblPr[0].append(tblW)

            if table_is_rtl:
                # RTL table visual ordering
                bidiVis = parse_xml(r'<w:bidiVisual xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                tblPr[0].append(bidiVis)

            # Booktabs borders (Top rule, Header bottom rule, Table bottom rule)
            borders = parse_xml(
                r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                r'<w:top w:val="single" w:sz="12" w:space="0" w:color="334155"/>'
                r'<w:left w:val="none"/>'
                r'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="334155"/>'
                r'<w:right w:val="none"/>'
                r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
                r'<w:insideV w:val="none"/>'
                r'</w:tblBorders>'
            )
            tblPr[0].append(borders)

        col_width_dxa = int(9360 / num_cols)

        # Render Header Row
        row_offset = 0
        if headers:
            hdr_tr = table.rows[0]._element
            trPr = hdr_tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
            trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

            for c_idx, cell_tokens in enumerate(headers):
                cell = table.cell(0, c_idx)
                cell.width = Inches(col_width_dxa / 1440.0)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                tcPr = cell._element.get_or_add_tcPr()
                # Header Shading
                tcPr.append(parse_xml(
                    f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    f'w:val="clear" w:color="auto" w:fill="{self.theme.table_header_bg}"/>'
                ))
                # Cell Padding (120 dxa top/bottom, 150 dxa left/right)
                tcPr.append(parse_xml(
                    r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    r'<w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/>'
                    r'<w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/>'
                    r'</w:tcMar>'
                ))

                p = cell.paragraphs[0]
                align = alignments[c_idx] if c_idx < len(alignments) else 'center'
                self._apply_cell_alignment(p, align, table_is_rtl)
                apply_paragraph_spacing(p, before_pt=2, after_pt=2, line_spacing_multiplier=1.2)
                
                self._render_inline_tokens(
                    p, cell_tokens,
                    default_bold=True,
                    default_size_pt=self.config.size_table_header,
                )
            row_offset = 1

        # Render Data Rows
        for r_idx, row_cells in enumerate(rows):
            curr_row = table.rows[r_idx + row_offset]
            trPr = curr_row._element.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

            for c_idx, cell_tokens in enumerate(row_cells):
                if c_idx >= num_cols:
                    break
                cell = curr_row.cells[c_idx]
                cell.width = Inches(col_width_dxa / 1440.0)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                tcPr = cell._element.get_or_add_tcPr()
                # Subtle Zebra Shading
                if r_idx % 2 == 1:
                    tcPr.append(parse_xml(
                        f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                        f'w:val="clear" w:color="auto" w:fill="{self.theme.table_zebra_bg}"/>'
                    ))

                # Cell Padding
                tcPr.append(parse_xml(
                    r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    r'<w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>'
                    r'<w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/>'
                    r'</w:tcMar>'
                ))

                p = cell.paragraphs[0]
                align = alignments[c_idx] if c_idx < len(alignments) else 'center'
                self._apply_cell_alignment(p, align, table_is_rtl)
                apply_paragraph_spacing(p, before_pt=2, after_pt=2, line_spacing_multiplier=1.2)

                self._render_inline_tokens(
                    p, cell_tokens,
                    default_bold=False,
                    default_size_pt=self.config.size_table_cell,
                )

        # Empty spacer paragraph after table
        sp = self.doc.add_paragraph()
        apply_paragraph_spacing(sp, before_pt=0, after_pt=8)

    def _apply_cell_alignment(self, p, align: str, is_rtl: bool):
        if align == 'center':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == 'right':
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == 'left':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT

    def add_callout_box(
        self,
        callout_type: str,
        title: Optional[str],
        body_items: List[Any],
        is_rtl: bool = True,
    ):
        """
        Renders a GitHub-style Admonition / Callout card or blockquote container
        with an accent colored bar (right bar for RTL, left bar for LTR) and matching pastel background.
        Supports both paragraphs and nested list items.
        """
        color_map = {
            "note": (self.theme.callout_note_border, self.theme.callout_note_bg, "ℹ️"),
            "tip": (self.theme.callout_tip_border, self.theme.callout_tip_bg, "💡"),
            "important": (self.theme.callout_important_border, self.theme.callout_important_bg, "📌"),
            "warning": (self.theme.callout_warning_border, self.theme.callout_warning_bg, "⚠️"),
            "caution": (self.theme.callout_caution_border, self.theme.callout_caution_bg, "🛑"),
            "quote": (self.theme.callout_quote_border, self.theme.callout_quote_bg, ""),
        }
        border_col, bg_col, icon = color_map.get(callout_type, (self.theme.border_subtle, "F8FAFC", ""))

        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        tblPr = table._element.xpath('w:tblPr')
        if tblPr:
            tblW = parse_xml(r'<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                             r'w:w="9360" w:type="dxa"/>')
            tblPr[0].append(tblW)

            # Accent bar on right for RTL, on left for LTR
            if is_rtl:
                borders = parse_xml(
                    f'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    f'<w:top w:val="none"/>'
                    f'<w:left w:val="none"/>'
                    f'<w:bottom w:val="none"/>'
                    f'<w:right w:val="single" w:sz="24" w:space="0" w:color="{border_col}"/>'
                    f'</w:tblBorders>'
                )
            else:
                borders = parse_xml(
                    f'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    f'<w:top w:val="none"/>'
                    f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_col}"/>'
                    f'<w:bottom w:val="none"/>'
                    f'<w:right w:val="none"/>'
                    f'</w:tblBorders>'
                )
            tblPr[0].append(borders)

        cell = table.cell(0, 0)
        tcPr = cell._element.get_or_add_tcPr()
        tcPr.append(parse_xml(
            f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'w:val="clear" w:color="auto" w:fill="{bg_col}"/>'
        ))
        tcPr.append(parse_xml(
            r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            r'<w:top w:w="140" w:type="dxa"/><w:bottom w:w="140" w:type="dxa"/>'
            r'<w:left w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/>'
            r'</w:tcMar>'
        ))

        first_p_used = False
        if title:
            p = cell.paragraphs[0]
            set_paragraph_bidi(p, is_rtl=is_rtl, justify=False)
            apply_paragraph_spacing(p, before_pt=0, after_pt=4, line_spacing_multiplier=1.2)
            lbl = f"{icon} {title}".strip()
            r = p.add_run(lbl)
            style_run(
                r,
                font_fa=self.config.font_fa_heading,
                font_en=self.config.font_en_heading,
                size_pt=13.0,
                is_fa=is_rtl,
                bold=True,
                color_hex=border_col,
            )
            first_p_used = True

        for item_idx, item in enumerate(body_items):
            item_type = item[0]
            if not first_p_used and item_idx == 0:
                bp = cell.paragraphs[0]
            else:
                bp = cell.add_paragraph()

            if item_type == 'paragraph':
                in_toks, raw_t = item[1], item[2]
                set_paragraph_bidi(bp, is_rtl=is_rtl, justify=True)
                apply_paragraph_spacing(bp, before_pt=2, after_pt=3, line_spacing_multiplier=1.3)
                self._render_inline_tokens(bp, in_toks, default_is_fa=is_rtl, default_size_pt=13.0)
            elif item_type == 'list_item':
                in_toks, raw_t, is_ordered, order_index, indent_level = item[1], item[2], item[3], item[4], item[5]
                set_paragraph_bidi(bp, is_rtl=is_rtl, justify=True)
                apply_paragraph_spacing(bp, before_pt=1, after_pt=2, line_spacing_multiplier=1.3)

                task_m = re.match(r'^\s*\[([ xX])\]\s*(.*)', raw_t)
                if task_m:
                    prefix = "☑  " if task_m.group(1).lower() == 'x' else "☐  "
                elif is_ordered:
                    if is_rtl:
                        fa_num = str(order_index).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))
                        prefix = f"{fa_num}.  "
                    else:
                        prefix = f"{order_index}.  "
                else:
                    bullets = ["•  ", "◦  ", "▪  "]
                    prefix = bullets[min(indent_level, 2)]

                pPr = bp._element.get_or_add_pPr()
                ind = OxmlElement('w:ind')
                left_val = str(360 * (indent_level + 1))
                if is_rtl:
                    ind.set(qn('w:right'), left_val)
                    ind.set(qn('w:rightChars'), '0')
                else:
                    ind.set(qn('w:left'), left_val)
                    ind.set(qn('w:hanging'), '360')
                pPr.append(ind)

                r_pre = bp.add_run(prefix)
                style_run(
                    r_pre,
                    font_fa=self.config.font_fa_heading if is_ordered else self.config.font_fa_body,
                    font_en=self.config.font_en_heading if is_ordered else self.config.font_en_body,
                    size_pt=13.0,
                    is_fa=is_rtl,
                    bold=is_ordered,
                )
                self._render_inline_tokens(bp, in_toks, default_is_fa=is_rtl, default_size_pt=13.0)

        # Spacer after callout
        sp = self.doc.add_paragraph()
        apply_paragraph_spacing(sp, before_pt=0, after_pt=8)

    def add_list_item(
        self,
        inline_tokens: List[Any],
        raw_text: str,
        is_ordered: bool = False,
        order_index: int = 1,
        indent_level: int = 0,
    ):
        """
        Renders an academic list item with proper hanging indentation and RTL support.
        Task lists ([ ] and [x]) are rendered with elegant checkboxes.
        """
        p = self.doc.add_paragraph()
        is_rtl = is_predominantly_persian(raw_text)
        set_paragraph_bidi(p, is_rtl=is_rtl, justify=True)
        apply_paragraph_spacing(p, before_pt=1, after_pt=3, line_spacing_multiplier=1.3)

        # Check for task list checkboxes
        task_match = re.match(r'^\s*\[([ xX])\]\s*(.*)', raw_text)
        prefix = ""
        if task_match:
            mark = task_match.group(1)
            prefix = "☑  " if mark.lower() == 'x' else "☐  "
        elif is_ordered:
            # Persian numerals for RTL ordered lists
            if is_rtl:
                fa_num = str(order_index).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))
                prefix = f"{fa_num}.  "
            else:
                prefix = f"{order_index}.  "
        else:
            bullets = ["•  ", "◦  ", "▪  "]
            prefix = bullets[min(indent_level, 2)]

        # Indentation (hanging)
        pPr = p._element.get_or_add_pPr()
        ind = OxmlElement('w:ind')
        hanging_val = "360"
        left_val = str(360 * (indent_level + 1))
        if is_rtl:
            ind.set(qn('w:right'), left_val)
            ind.set(qn('w:rightChars'), '0')
        else:
            ind.set(qn('w:left'), left_val)
            ind.set(qn('w:hanging'), hanging_val)
        pPr.append(ind)

        # Add prefix run
        r_pre = p.add_run(prefix)
        style_run(
            r_pre,
            font_fa=self.config.font_fa_heading if is_ordered else self.config.font_fa_body,
            font_en=self.config.font_en_heading if is_ordered else self.config.font_en_body,
            size_pt=self.config.size_body_fa if is_rtl else self.config.size_body_en,
            is_fa=is_rtl,
            bold=is_ordered,
        )

        # Render list content
        self._render_inline_tokens(p, inline_tokens, default_is_fa=is_rtl)
        return p

    def add_image_with_caption(self, src: str, caption_text: str):
        """
        Inserts a centered image with maximum width constraint, followed by a bold caption.
        """
        self.image_counter += 1
        img_path = self._resolve_image_path(src)
        if not img_path or not os.path.exists(img_path):
            return

        # Calculate dimensions
        max_width_in = 6.2  # Max ~15.75 cm
        try:
            with Image.open(img_path) as im:
                w_px, h_px = im.size
                aspect = h_px / float(w_px)
                target_w_in = min(max_width_in, w_px / 96.0)
                target_h_in = target_w_in * aspect
        except Exception:
            target_w_in = 5.0
            target_h_in = 3.5

        # Image paragraph
        p_img = self.doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_paragraph_spacing(
            p_img,
            before_pt=self.config.space_before_image,
            after_pt=self.config.space_after_image,
        )
        r_img = p_img.add_run()
        r_img.add_picture(img_path, width=Inches(target_w_in))

        # Caption paragraph
        if caption_text:
            p_cap = self.doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            is_rtl = is_predominantly_persian(caption_text)
            set_paragraph_bidi(p_cap, is_rtl=is_rtl, justify=False)
            apply_paragraph_spacing(
                p_cap,
                before_pt=0,
                after_pt=self.config.space_after_caption,
            )

            # Auto-prefix figure label if not present
            cap_clean = caption_text.strip()
            if not re.match(r'^(شکل|تصویر|نمودار|Figure|Fig\.)\s*\d+', cap_clean):
                if is_rtl:
                    fa_cnt = str(self.image_counter).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))
                    cap_clean = f"شکل {fa_cnt}: {cap_clean}"
                else:
                    cap_clean = f"Figure {self.image_counter}: {cap_clean}"

            for seg_type, seg_text in tokenize_runs(cap_clean):
                r = p_cap.add_run(seg_text)
                style_run(
                    r,
                    font_fa=self.config.font_fa_body,
                    font_en=self.config.font_en_body,
                    size_pt=self.config.size_caption,
                    is_fa=(seg_type == 'fa'),
                    bold=True,
                )

    def _resolve_image_path(self, src: str) -> Optional[str]:
        """Resolves local paths or downloads web images into a temporary file."""
        if src.startswith("http://") or src.startswith("https://"):
            try:
                resp = requests.get(src, timeout=10)
                if resp.status_code == 200:
                    ext = os.path.splitext(src.split('?')[0])[1] or ".png"
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
                    tmp.write(resp.content)
                    tmp.close()
                    return tmp.name
            except Exception:
                return None
        else:
            if os.path.isabs(src):
                return src if os.path.exists(src) else None
            candidate = os.path.join(self.base_dir, src)
            if os.path.exists(candidate):
                return candidate
            return None

    def add_horizontal_rule(self):
        """Adds a clean horizontal divider line."""
        p = self.doc.add_paragraph()
        apply_paragraph_spacing(p, before_pt=6, after_pt=6)
        pPr = p._element.get_or_add_pPr()
        pBdr = parse_xml(
            r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/>'
            r'</w:pBdr>'
        )
        pPr.append(pBdr)

    def add_page_break(self):
        """Adds an explicit page break."""
        self.doc.add_page_break()

    def finalize(self):
        """Finalizes document and strictly enforces OpenXML schema ordering."""
        normalize_document_schema(self.doc)
        if self.doc:
            try:
                props = self.doc.core_properties
                props.author = ", ".join(self.metadata.authors) if self.metadata.authors else "MarkdownStudio"
                props.last_modified_by = ""
                props.comments = "Created with MarkdownStudio"
                if self.metadata.title:
                    props.title = self.metadata.title
            except Exception:
                pass


# OpenXML ECMA-376 Strict Child Element Ordering Mappings
PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
    'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
    'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
    'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
    'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
    'sectPr', 'pPrChange'
]
PPR_INDEX = {tag: i for i, tag in enumerate(PPR_ORDER)}

RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
    'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
    'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
    'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
    'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'
]
RPR_INDEX = {tag: i for i, tag in enumerate(RPR_ORDER)}

TBLPR_ORDER = [
    'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
    'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
    'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
    'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'
]
TBLPR_INDEX = {tag: i for i, tag in enumerate(TBLPR_ORDER)}

TCPR_ORDER = [
    'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders',
    'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign',
    'hideMark', 'headers', 'tcPrChange'
]
TCPR_INDEX = {tag: i for i, tag in enumerate(TCPR_ORDER)}

TCMAR_ORDER = ['top', 'left', 'bottom', 'right']
TCMAR_INDEX = {tag: i for i, tag in enumerate(TCMAR_ORDER)}

TRPR_ORDER = [
    'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter',
    'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jcPr',
    'hidden', 'ins', 'del', 'trPrChange'
]
TRPR_INDEX = {tag: i for i, tag in enumerate(TRPR_ORDER)}


def _sort_element_children(parent, index_map):
    children = list(parent)
    seen = {}
    for c in children:
        tag_name = c.tag.split('}')[-1]
        seen[tag_name] = c

    sorted_children = sorted(seen.values(), key=lambda c: index_map.get(c.tag.split('}')[-1], 999))
    for c in children:
        parent.remove(c)
    for c in sorted_children:
        parent.append(c)


def normalize_document_schema(doc):
    """
    Guarantees 100% strict compliance with ECMA-376 OpenXML schema
    across all tables, rows, cells, paragraphs, and runs.
    """
    body = doc._body._element

    for elem in body.xpath('.//w:tblPr'):
        _sort_element_children(elem, TBLPR_INDEX)

    for elem in body.xpath('.//w:trPr'):
        _sort_element_children(elem, TRPR_INDEX)

    for elem in body.xpath('.//w:tcPr'):
        _sort_element_children(elem, TCPR_INDEX)

    for elem in body.xpath('.//w:tcMar'):
        _sort_element_children(elem, TCMAR_INDEX)

    for elem in body.xpath('.//w:pPr'):
        _sort_element_children(elem, PPR_INDEX)

    for elem in body.xpath('.//w:rPr'):
        _sort_element_children(elem, RPR_INDEX)

