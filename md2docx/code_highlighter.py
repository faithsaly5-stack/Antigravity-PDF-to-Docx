"""
Syntax Highlighting and Code Block Formatting Engine for Word Documents.
Uses Pygments to tokenize code and maps tokens to OpenXML colored runs.
"""

import docx
from typing import List, Tuple, Optional
from pygments import lex
from pygments.lexers import get_lexer_by_name, guess_lexer, TextLexer
from pygments.token import Token
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

# Token Color Mapping (clean GitHub-like syntax theme)
TOKEN_COLOR_MAP = {
    Token.Keyword: "0550AE",           # Blue
    Token.Keyword.Constant: "0550AE",
    Token.Keyword.Declaration: "0550AE",
    Token.Keyword.Namespace: "CF222E",   # Red
    Token.Keyword.Type: "953800",        # Orange/Brown
    Token.Name.Class: "953800",
    Token.Name.Function: "8250DF",       # Purple
    Token.Name.Builtin: "8250DF",
    Token.String: "0A3069",             # Dark Blue/Green
    Token.String.Doc: "57606A",          # Gray (Docstring)
    Token.Comment: "6E7781",            # Gray
    Token.Comment.Single: "6E7781",
    Token.Comment.Multiline: "6E7781",
    Token.Number: "0550AE",             # Blue
    Token.Operator: "24292F",
    Token.Punctuation: "24292F",
    Token.Name.Decorator: "8250DF",
    Token.Generic.Deleted: "82071E",
    Token.Generic.Inserted: "116329",
}


def get_token_color(token_type) -> str:
    """Finds the best matching hex color for a pygments token type."""
    curr = token_type
    while curr is not None:
        if curr in TOKEN_COLOR_MAP:
            return TOKEN_COLOR_MAP[curr]
        curr = curr.parent
    return "24292F"  # Default charcoal black


def tokenize_code(code: str, lang: Optional[str] = None) -> List[Tuple[str, str, bool, bool]]:
    """
    Tokenizes code into (text_chunk, hex_color, is_bold, is_italic).
    """
    lexer = None
    if lang:
        try:
            lexer = get_lexer_by_name(lang.strip().lower())
        except Exception:
            pass

    if lexer is None:
        try:
            lexer = guess_lexer(code)
        except Exception:
            lexer = TextLexer()

    tokens = []
    for token_type, value in lex(code, lexer):
        color = get_token_color(token_type)
        is_bold = token_type in Token.Keyword
        is_italic = token_type in Token.Comment
        tokens.append((value, color, is_bold, is_italic))

    return tokens


def insert_code_block(doc, code: str, lang: Optional[str] = None, font_name: str = "Consolas", font_size_pt: float = 10.0):
    """
    Renders a syntax-highlighted code block inside a styled single-cell card in Word.
    Ensures strict LTR layout (no w:bidi, no w:rtl), background shading, and borders.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Configure Table Properties (tblPr)
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        # Table width 100% (approx 9360 dxa for A4 / Letter)
        tblW = parse_xml(r'<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                         r'w:w="9360" w:type="dxa"/>')
        tblPr[0].append(tblW)
        
        # Border
        borders = parse_xml(
            r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            r'<w:top w:val="single" w:sz="6" w:space="0" w:color="D0D5DD"/>'
            r'<w:left w:val="single" w:sz="6" w:space="0" w:color="D0D5DD"/>'
            r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="D0D5DD"/>'
            r'<w:right w:val="single" w:sz="6" w:space="0" w:color="D0D5DD"/>'
            r'</w:tblBorders>'
        )
        tblPr[0].append(borders)

    cell = table.cell(0, 0)
    # Cell Shading & Padding
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    r'w:val="clear" w:color="auto" w:fill="F8F9FA"/>')
    tcPr.append(shd)

    # Margins inside cell (dxa)
    mar = parse_xml(
        r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:w="140" w:type="dxa"/>'
        r'<w:bottom w:w="140" w:type="dxa"/>'
        r'<w:left w:w="180" w:type="dxa"/>'
        r'<w:right w:w="180" w:type="dxa"/>'
        r'</w:tcMar>'
    )
    tcPr.append(mar)

    # Clear initial paragraph
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Configure Paragraph Properties (pPr) for LTR
    pPr = p._element.get_or_add_pPr()
    bidi_elem = pPr.find(qn('w:bidi'))
    if bidi_elem is not None:
        pPr.remove(bidi_elem)

    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '0')
    sp.set(qn('w:after'), '0')
    sp.set(qn('w:line'), '240')  # Single line spacing
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

    tokens = tokenize_code(code.rstrip('\r\n'), lang)
    lines_tokens: List[List[Tuple[str, str, bool, bool]]] = [[]]
    
    for val, col, bld, itl in tokens:
        if '\n' in val:
            parts = val.split('\n')
            for idx, part in enumerate(parts):
                if part:
                    lines_tokens[-1].append((part, col, bld, itl))
                if idx < len(parts) - 1:
                    lines_tokens.append([])
        else:
            lines_tokens[-1].append((val, col, bld, itl))

    sz_val = str(int(font_size_pt * 2))

    for line_idx, line in enumerate(lines_tokens):
        if line_idx > 0:
            p = cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr = p._element.get_or_add_pPr()
            bidi_elem = pPr.find(qn('w:bidi'))
            if bidi_elem is not None:
                pPr.remove(bidi_elem)
            sp = OxmlElement('w:spacing')
            sp.set(qn('w:before'), '0')
            sp.set(qn('w:after'), '0')
            sp.set(qn('w:line'), '240')
            sp.set(qn('w:lineRule'), 'auto')
            pPr.append(sp)

        if not line:
            # Empty line
            r = p.add_run(" ")
            rPr = r._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rPr.append(rFonts)
            continue

        for txt, col, is_b, is_i in line:
            r = p.add_run(txt)
            rPr = r._element.get_or_add_rPr()
            
            # Fonts
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rPr.append(rFonts)

            # Font Size
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), sz_val)
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), sz_val)
            rPr.append(szCs)

            # Color
            c = OxmlElement('w:color')
            c.set(qn('w:val'), col)
            rPr.append(c)

            # Bold / Italic
            if is_b:
                rPr.append(OxmlElement('w:b'))
            if is_i:
                rPr.append(OxmlElement('w:i'))

            # Strictly ensure NO w:rtl flag
            rtl_elem = rPr.find(qn('w:rtl'))
            if rtl_elem is not None:
                rPr.remove(rtl_elem)

    # Empty spacer paragraph after code block
    spacer = doc.add_paragraph()
    spPr = spacer._element.get_or_add_pPr()
    sp_spacing = OxmlElement('w:spacing')
    sp_spacing.set(qn('w:before'), '0')
    sp_spacing.set(qn('w:after'), '120')
    spPr.append(sp_spacing)
