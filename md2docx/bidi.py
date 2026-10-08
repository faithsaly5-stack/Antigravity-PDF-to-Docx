"""
Bidirectional Text Engine for Persian / RTL & Mixed-Script Typography.
Implements Tokenized Run Generation and Reference BIDI Isolation.
"""

import re
from typing import List, Tuple
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Unicode Ranges
PERSIAN_ARABIC_PATTERN = re.compile(r'[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]')
LATIN_CHARS_PATTERN = re.compile(r'[a-zA-Z²³⁴⁺⁻₂₃₄½ελμπαβγδεθλσωΩ]')

# Regex for English words, formulas, scientific notation, units, and parentheses enclosing them
MIXED_SCRIPT_PATTERN = re.compile(
    r'([a-zA-Z0-9\+\-\=\/\(\)\[\]\<\>\±\≈\×\·\_\^\~²³⁴⁺⁻₂₃₄½ελμπαβγδεθλσωΩ°%]+'
    r'(?:\s+[a-zA-Z0-9\+\-\=\/\(\)\[\]\<\>\±\≈\×\·\_\^\~²³⁴⁺⁻₂₃₄½ελμπαβγδεθλσωΩ°%]+)*)'
)


def has_persian_chars(text: str) -> bool:
    """Check if text contains any Persian/Arabic script characters."""
    return bool(PERSIAN_ARABIC_PATTERN.search(text))


def is_predominantly_persian(text: str) -> bool:
    """
    Determine if a paragraph or text block is predominantly Persian/RTL.
    Returns False if it is purely Latin or looks like an English citation.
    """
    if not text or not text.strip():
        return True  # default to doc direction
    
    # Check if text looks like an English reference/citation
    if is_english_reference(text):
        return False
        
    fa_count = len(PERSIAN_ARABIC_PATTERN.findall(text))
    en_count = len(re.findall(r'[a-zA-Z]', text))
    
    if fa_count > 0 and fa_count >= en_count * 0.4:
        return True
    if fa_count > 0 and en_count == 0:
        return True
    return False


def is_english_reference(text: str) -> bool:
    """
    Check if a text line is an English reference citation (APA, IEEE, Vancouver, etc.)
    e.g. "1. Chang, R. (2010). Chemistry." or "[1] Smith, J., & Doe, A. (2020)."
    """
    s = text.strip()
    # Strip leading bullet, number, or citation bracket e.g. "1. ", "[1] ", "- "
    s_clean = re.sub(r'^(?:\[\d+\]|\d+[\.\-\)]|\-|\*)\s*', '', s).strip()
    
    # If cleaned text starts with Persian characters, it's definitely NOT an English ref
    if PERSIAN_ARABIC_PATTERN.match(s_clean):
        return False
        
    # Check for typical Latin citation markers: Author name, year in parens, Latin letters
    if re.search(r'^[A-Z][a-zA-Z\.\s\-\,]+(?:\([12]\d{3}\)|\,\s*[12]\d{3}|:\s*[12]\d{3})', s_clean):
        return True
        
    fa_count = len(PERSIAN_ARABIC_PATTERN.findall(s))
    en_count = len(re.findall(r'[a-zA-Z]', s))
    if en_count > 15 and fa_count == 0:
        return True
    return False


def tokenize_runs(text: str) -> List[Tuple[str, str]]:
    """
    Tokenized Run Generation:
    Splits text into Persian ('fa') runs and Latin/Formula ('en') runs.
    Ensures that chemical formulas (e.g. Cl⁻, NaOH), Latin variables, and numbers
    are NOT subjected to w:rtl, which prevents parenthesis inversion.
    """
    if not text:
        return []
        
    # If the text has no Persian characters at all, treat whole text as 'en'
    if not has_persian_chars(text):
        return [('en', text)]
        
    # If the text has no Latin/formula characters, treat whole text as 'fa'
    if not LATIN_CHARS_PATTERN.search(text):
        return [('fa', text)]

    segments: List[Tuple[str, str]] = []
    last_idx = 0

    for match in MIXED_SCRIPT_PATTERN.finditer(text):
        start, end = match.span()
        tok = match.group()

        # Only classify as 'en' if it contains actual Latin letters, math symbols, or greek
        if LATIN_CHARS_PATTERN.search(tok):
            if start > last_idx:
                fa_part = text[last_idx:start]
                if fa_part:
                    segments.append(('fa', fa_part))
            segments.append(('en', tok))
            last_idx = end

    if last_idx < len(text):
        remaining = text[last_idx:]
        if remaining:
            segments.append(('fa', remaining))

    if not segments:
        segments = [('fa', text)]

    return segments


def set_paragraph_bidi(p, is_rtl: bool = True, justify: bool = True):
    """
    Configures OpenXML paragraph bidirectional properties.
    For Persian: adds <w:bidi/> and sets justification (both).
    For English/References: ensures NO <w:bidi/> is present.
    """
    pPr = p._element.get_or_add_pPr()
    bidi_elem = pPr.find(qn('w:bidi'))
    
    if is_rtl:
        if bidi_elem is None:
            pPr.append(OxmlElement('w:bidi'))
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        if bidi_elem is not None:
            pPr.remove(bidi_elem)
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def apply_paragraph_spacing(p, before_pt: float = 0.0, after_pt: float = 8.0, line_spacing_multiplier: float = 1.5):
    """
    Configures OpenXML paragraph spacing and line height.
    Line spacing 1.5 lines = 360 dxa in OpenXML (240 dxa = 1.0 line).
    """
    pPr = p._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is not None:
        pPr.remove(sp)
        
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(int(before_pt * 20)))
    sp.set(qn('w:after'), str(int(after_pt * 20)))
    sp.set(qn('w:line'), str(int(line_spacing_multiplier * 240)))
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)


def style_run(
    r,
    font_fa: str = "B Nazanin",
    font_en: str = "Times New Roman",
    size_pt: float = 14.0,
    is_fa: bool = True,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    strike: bool = False,
    subscript: bool = False,
    superscript: bool = False,
    color_hex: str = None,
    highlight_color: str = None,
    is_code: bool = False,
):
    """
    Applies comprehensive OpenXML run properties (rPr) for typography.
    Strictly follows Complex Script (w:cs) and w:rtl isolation.
    """
    rPr = r._element.get_or_add_rPr()

    # Font Assignment
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        rPr.remove(rFonts)
    rFonts = OxmlElement('w:rFonts')

    if is_code:
        rFonts.set(qn('w:ascii'), font_en)
        rFonts.set(qn('w:hAnsi'), font_en)
        rFonts.set(qn('w:cs'), font_en)
    elif is_fa:
        rFonts.set(qn('w:ascii'), font_fa)
        rFonts.set(qn('w:hAnsi'), font_fa)
        rFonts.set(qn('w:cs'), font_fa)
    else:
        rFonts.set(qn('w:ascii'), font_en)
        rFonts.set(qn('w:hAnsi'), font_en)
        rFonts.set(qn('w:cs'), font_fa)
    rPr.append(rFonts)

    # Size (half-points)
    sz_val = str(int(size_pt * 2))
    sz = rPr.find(qn('w:sz'))
    if sz is not None:
        rPr.remove(sz)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), sz_val)
    rPr.append(sz)

    szCs = rPr.find(qn('w:szCs'))
    if szCs is not None:
        rPr.remove(szCs)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), sz_val)
    rPr.append(szCs)

    # Bold
    if bold:
        if rPr.find(qn('w:b')) is None:
            rPr.append(OxmlElement('w:b'))
        if rPr.find(qn('w:bCs')) is None:
            rPr.append(OxmlElement('w:bCs'))

    # Italic
    if italic:
        if rPr.find(qn('w:i')) is None:
            rPr.append(OxmlElement('w:i'))
        if rPr.find(qn('w:iCs')) is None:
            rPr.append(OxmlElement('w:iCs'))

    # Underline
    if underline:
        u = rPr.find(qn('w:u'))
        if u is None:
            u = OxmlElement('w:u')
            u.set(qn('w:val'), 'single')
            rPr.append(u)

    # Strikethrough
    if strike:
        if rPr.find(qn('w:strike')) is None:
            rPr.append(OxmlElement('w:strike'))

    # Subscript / Superscript
    if subscript:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), 'subscript')
        rPr.append(va)
    elif superscript:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), 'superscript')
        rPr.append(va)

    # Font Color
    if color_hex:
        clean_color = color_hex.lstrip('#')
        c = rPr.find(qn('w:color'))
        if c is not None:
            rPr.remove(c)
        c = OxmlElement('w:color')
        c.set(qn('w:val'), clean_color)
        rPr.append(c)

    # Highlight (shading)
    if highlight_color:
        shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="clear" w:color="auto" w:fill="{highlight_color.lstrip("#")}"/>')
        rPr.append(shd)

    # RTL Run Flag: strictly for Persian runs, NEVER for Latin or Code runs!
    rtl_elem = rPr.find(qn('w:rtl'))
    if is_fa and not is_code:
        if rtl_elem is None:
            rPr.append(OxmlElement('w:rtl'))
    else:
        if rtl_elem is not None:
            rPr.remove(rtl_elem)
