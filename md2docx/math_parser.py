"""
Math Equation Engine: Converts LaTeX equations to native Microsoft Word OMML (OfficeMath).
Uses latex2mathml and Microsoft's official MML2OMML.XSL stylesheet.
"""

import os
import re
from typing import Optional
import lxml.etree as ET
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Path to bundled MML2OMML.XSL
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
DEFAULT_XSLT_PATH = os.path.join(PROJECT_ROOT, "assets", "MML2OMML.XSL")

# Cache transformer
_XSLT_TRANSFORMER: Optional[ET.XSLT] = None


def get_xslt_transformer() -> Optional[ET.XSLT]:
    """Load and cache the MML2OMML XSLT transformer."""
    global _XSLT_TRANSFORMER
    if _XSLT_TRANSFORMER is not None:
        return _XSLT_TRANSFORMER

    xslt_candidates = [
        DEFAULT_XSLT_PATH,
        r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    ]

    for path in xslt_candidates:
        if os.path.exists(path):
            try:
                xslt_tree = ET.parse(path)
                _XSLT_TRANSFORMER = ET.XSLT(xslt_tree)
                return _XSLT_TRANSFORMER
            except Exception:
                continue

    return None


def latex_to_omml(latex_code: str, is_display: bool = False) -> Optional[OxmlElement]:
    """
    Converts a LaTeX equation string to a native Word OMML XML element.
    Returns <m:oMath> for inline equations, or <m:oMathPara> for display block equations.
    """
    latex_clean = latex_code.strip()
    if not latex_clean:
        return None

    # Strip surrounding $ or $$ if present
    if latex_clean.startswith("$$") and latex_clean.endswith("$$"):
        latex_clean = latex_clean[2:-2].strip()
    elif latex_clean.startswith("$") and latex_clean.endswith("$"):
        latex_clean = latex_clean[1:-1].strip()

    # Preprocess common LaTeX macros for cleaner MathML conversion
    latex_clean = re.sub(r'\\left\(', '(', latex_clean)
    latex_clean = re.sub(r'\\right\)', ')', latex_clean)
    latex_clean = re.sub(r'\\left\[', '[', latex_clean)
    latex_clean = re.sub(r'\\right\]', ']', latex_clean)

    transformer = get_xslt_transformer()
    if not transformer:
        return None

    try:
        import latex2mathml.converter
        mathml_str = latex2mathml.converter.convert(latex_clean)
        mathml_tree = ET.fromstring(mathml_str)
        omml_tree = transformer(mathml_tree)
        omml_xml = ET.tostring(omml_tree, encoding='unicode')

        # Clean namespace prefix issues if any
        if is_display:
            # Wrap in <m:oMathPara>
            if "<m:oMathPara" not in omml_xml:
                omml_xml = (
                    f'<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                    f'{omml_xml}'
                    f'</m:oMathPara>'
                )
        return parse_xml(omml_xml)
    except Exception as e:
        # Fallback will be handled gracefully by caller
        return None


def insert_display_equation(doc, latex_code: str, space_before_pt: float = 8.0, space_after_pt: float = 8.0):
    """
    Inserts a standalone centered display equation in docx.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Configure spacing
    pPr = p._element.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(int(space_before_pt * 20)))
    sp.set(qn('w:after'), str(int(space_after_pt * 20)))
    sp.set(qn('w:line'), '240')
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

    omml = latex_to_omml(latex_code, is_display=True)
    if omml is not None:
        p._element.append(omml)
    else:
        # Fallback to Cambria Math styled run
        r = p.add_run(latex_code.strip('$'))
        r.font.name = 'Cambria Math'
        r.font.italic = True
    return p
