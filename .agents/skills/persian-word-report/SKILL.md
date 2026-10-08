---
name: persian-word-report
description: >-
  Use this skill whenever you need to create, format, or edit flawless Persian (Farsi) or RTL academic documents,
  laboratory reports, theses, or research papers in Microsoft Word (.docx) with PDF export.
  Ensures professional typography (B Nazanin, B Titr, Cambria Math), automated Table of Contents (TOC),
  correct bidirectional mixed-text handling (avoiding flipped parentheses and distorted formulas),
  reference bidi isolation (preventing inverted APA English citations),
  university cover page layouts, academic data tables, high-resolution centered images with captions,
  and native Word COM automation.
---

# Persian Academic Word Document & Report Engineering

This skill provides a complete, battle-tested standard for generating and maintaining publication-quality Persian/RTL Microsoft Word (`.docx`) academic documents and converting them to flawless PDFs.

---

## 1. Core Principles & Golden Typography Rules

### 1. The Mixed-Script Dilemma (Persian + Latin / Chemical Formulas)
When Persian text and Latin terms/chemical formulas (e.g., $\text{Cl}^-$, $\text{Fe}^{3+}$, $\text{FeSCN}^{2+}$, $\text{NaOH}$) share a paragraph:
- Applying `w:rtl` to an entire run containing English terms or chemical formulas **inverts parentheses** (e.g. `(Cl⁻)` becomes `)Cl⁻(`), swaps minus signs, and distorts superscripts.
- Using plain `python-docx` default `add_paragraph` strips away OpenXML Complex Script (`w:cs`), bidirectional marks (`w:bidi`), and line spacing rules.
- **The Solution:** Use **Tokenized Run Generation**. Every paragraph is split by regex into Persian runs (`B Nazanin`, `w:rtl`) and Latin/Formula runs (`Times New Roman` or `Cambria Math`, strictly **without** `w:rtl`).

### 2. The Reference BIDI Isolation Rule (CRITICAL)
In academic documents containing both English (APA, Vancouver) and Persian references:
- **English References (e.g., Chang, Harris, Skoog, Brown):**
  - **MUST NOT** have `<w:bidi/>` in the paragraph properties (`pPr`).
  - Font: `Times New Roman`, LTR run direction, `WD_ALIGN_PARAGRAPH.JUSTIFY`.
  - *Why:* If `<w:bidi/>` is enabled on an English citation, Word treats the paragraph as RTL, which inverts year parentheses (e.g. `(2016).` flips), places periods on the far left, and scrambles author initials.
- **Persian References:**
  - **MUST** have `<w:bidi/>` in the paragraph properties (`pPr`).
  - Font: `B Nazanin` with `<w:rtl/>`, `WD_ALIGN_PARAGRAPH.JUSTIFY`.

### 3. Formula Naming Standard (Pure-Latin Identifiers)
- Never embed Persian words directly inside ASCII parentheses adjacent to Latin variables:
  - ❌ **Incorrect:** `M(اسید مجهول) = ۰٫۱۲۰ ± ۰٫۰۱۰ mol/L` (causes bi-directional parenthesis flip in Word).
  - ✅ **Correct:** `M(acid) = ۰٫۱۲۰ ± ۰٫۰۱۰ mol/L` (pure Latin symbols ensure zero flipping).

### 4. Heading & Body Justification
- In Persian RTL layout, set `WD_ALIGN_PARAGRAPH.JUSTIFY` (`w:jc w:val="both"`) on **both Headings (`Heading 1`, `Heading 2`) and Normal body text**.
- In Arabic/Persian typesetting, Justify anchors text cleanly to the right margin with baseline justification, eliminating ragged line ends and matching formal book/thesis publication standards.

### 5. Persian Subtitle Hyphenation
- In titles and headings, use spaced hyphens (` - `) rather than tight en-dashes (`–`):
  - ❌ `تیتراسیون اسید–باز` (in bold fonts like `B Titr`, the dash collides visually with adjacent characters like 'د' and 'ب').
  - ✅ `تیتراسیون اسید - باز` (provides clean visual breathing room).

### 6. Template Inheritance (The Golden Rule)
Never start a complex Persian document from a blank `docx.Document()`. Always load a proven master template (e.g., `templates/default_academic.docx`) that already contains:
- Pre-configured page margins ($2.0\text{ cm}$ all sides, Letter or A4).
- Defined styles (`Heading 1`, `Heading 2`, `Normal`, `TOC 1`, `TOC 2`).
- Complex Script font defaults (`B Nazanin`, `B Titr`).
- Header/Footer geometry.
- Embedded vector artwork, cover page borders, and university logos.

---

## 2. Document Architecture

A standard university or professional report consists of three main structural zones:

```
[ Page 1 ] Cover Page (Artwork/Logo + Titles + Professors + Students + Date)
           ── Page Break (<w:br w:type="page"/>) ──
[ Page 2 ] Table of Contents (Heading "فهرست مطالب" + <w:sdt> TOC Field)
           ── Page Break (<w:br w:type="page"/>) ──
[ Page 3+] Main Document Body (Headings, Body Text, Equations, Tables, Images, References)
```

### Surgical Body Reset (Preserving Cover & TOC Skeleton)
To populate new content while keeping the cover page and TOC intact from a template:
```python
doc = docx.Document("master_template.docx")
body = doc._body._element
children = list(body)

# In standard templates:
# 0..7: Cover page elements
# 8: Page break (Page 1 -> Page 2)
# 9: TOC heading ("فهرست مطالب")
# 10: <w:sdt> TOC field
# 11: Empty spacer
# 12: Page break (Page 2 -> Page 3)
# 13+: Old body content

sectPr = body.find(qn('w:sectPr'))
# Remove old body elements (index 13 onwards) while preserving sectPr
for child in children[13:]:
    if child != sectPr:
        body.remove(child)
```

### In-Place Text Replacement on Cover Page
Modify the `<w:t>` text nodes directly to avoid wiping underlying formatting properties:
```python
def update_cover_text(p_element, new_text):
    t_nodes = p_element.xpath('.//w:t')
    if t_nodes:
        t_nodes[0].text = new_text
        for extra_t in t_nodes[1:]:
            extra_t.getparent().remove(extra_t)

update_cover_text(body[1], "گزارش کار آزمایشگاه‌های شیمی عمومی و شیمی تجزیه")
update_cover_text(body[2], "شناسایی آنیون‌ها  |  تعیین ثابت تعادل  |  تیتراسیون اسید - باز")
update_cover_text(body[4], "اساتید: نام استاد راهنما، نام استاد مشاور")
update_cover_text(body[5], "دانشجو: نام دانشجو / پژوهشگر")
update_cover_text(body[7], "تیرماه ۱۴۰۵")
update_cover_text(body[9], "فهرست مطالب")
```

---

## 3. Typography & Formatting Standards

### Font Matrix
| Element | Font (Persian) | Font (Latin / Math) | Size | Attributes | Alignment |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Cover Title** | B Nazanin | — | 20 pt (40 dxa) | Bold | Center |
| **Cover Subtitle**| B Nazanin | — | 18 pt (36 dxa) | Bold | Center |
| **Cover Metadata**| B Nazanin | — | 15 pt (30 dxa) | Regular | Center |
| **Cover Date** | B Nazanin | — | 16 pt (32 dxa) | Bold | Center |
| **TOC Title** | B Titr | — | 16 pt (32 dxa) | Bold | Center / Justify |
| **Heading 1** | B Titr | B Titr | 17 pt (34 dxa) | Bold | Justify (`jc="both"`) |
| **Heading 2** | B Titr | B Titr | 14 pt (28 dxa) | Bold | Justify (`jc="both"`) |
| **Normal Body** | B Nazanin | Times New Roman | 14 pt (28 dxa) | Regular | Justify (`jc="both"`) |
| **Equations** | — | Cambria Math | 14 pt (28 dxa) | Bold | Justify / Center |
| **Table Header**| B Titr | B Titr | 12 pt (24 dxa) | Bold | Center / Justify |
| **Table Data** | B Nazanin | Times New Roman | 13 pt (26 dxa) | Regular | Center / Justify |
| **Image Caption**| B Nazanin | Times New Roman | 12 pt (24 dxa) | Bold | Center |
| **English Ref** | — | Times New Roman | 12 pt (24 dxa) | Regular (NO bidi) | Justify (`jc="both"`) |
| **Persian Ref** | B Nazanin | — | 13 pt (26 dxa) | Regular (bidi=True)| Justify (`jc="both"`) |

### Spacing Standards
- **Body Text:** Line spacing $1.5$ lines (`w:line="360" w:lineRule="auto"`), Space after $8\text{ pt}$ (`w:after="160"`), Space before $0\text{ pt}$.
- **Heading 1:** Space before $14\text{ pt}$ (`w:before="280"`), Space after $6\text{ pt}$ (`w:after="120"`).
- **Heading 2:** Space before $10\text{ pt}$ (`w:before="200"`), Space after $5\text{ pt}$ (`w:after="100"`).
- **Standalone Equation:** Space before $8\text{ pt}$ (`160 dxa`), Space after $8\text{ pt}$ (`160 dxa`).
- **Images:** Space before $7\text{ pt}$ (`140 dxa`), Space after $7\text{ pt}$ (`140 dxa`), Captions Space after $8\text{ pt}$.

---

## 4. Implementation Helper Functions (Python)

```python
import re
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_rich_paragraph(doc, text, style='Normal', level=0):
    style_name = 'Heading 1' if level == 1 else ('Heading 2' if level == 2 else 'Normal')
    p = doc.add_paragraph(style=style_name)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    pPr = p._element.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None:
        pPr.append(OxmlElement('w:bidi'))
        
    sp = OxmlElement('w:spacing')
    if level == 1:
        sp.set(qn('w:before'), '280')
        sp.set(qn('w:after'), '120')
        sp.set(qn('w:line'), '360')
    elif level == 2:
        sp.set(qn('w:before'), '200')
        sp.set(qn('w:after'), '100')
        sp.set(qn('w:line'), '360')
    else:
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '160')
        sp.set(qn('w:line'), '360')
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

    pattern = re.compile(r'([a-zA-Z0-9\+\-\=\/\(\)\[\]\<\>\±\≈\×\·\_\^\~²³⁴⁺⁻₂₃₄½ελ]+(?:\s+[a-zA-Z0-9\+\-\=\/\(\)\[\]\<\>\±\≈\×\·\_\^\~²³⁴⁺⁻₂₃₄½ελ]+)*)')
    segments = []
    last_idx = 0
    for m in pattern.finditer(text):
        start, end = m.span()
        tok = m.group()
        if any(c in tok for c in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ²³⁴⁺⁻₂₃₄½ελ'):
            if start > last_idx:
                segments.append(('fa', text[last_idx:start]))
            segments.append(('en', tok))
            last_idx = end
            
    if last_idx < len(text):
        segments.append(('fa', text[last_idx:]))
        
    if not segments:
        segments = [('fa', text)]
        
    for seg_type, seg_text in segments:
        r = p.add_run(seg_text)
        rPr = r._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        
        if level == 1:
            font_fa, font_en, sz_val, bold_flag = "B Titr", "B Titr", "34", True
        elif level == 2:
            font_fa, font_en, sz_val, bold_flag = "B Titr", "B Titr", "28", True
        else:
            font_fa, font_en, sz_val, bold_flag = "B Nazanin", "Times New Roman", "28", False
            
        if seg_type == 'fa':
            rFonts.set(qn('w:ascii'), font_fa); rFonts.set(qn('w:hAnsi'), font_fa); rFonts.set(qn('w:cs'), font_fa)
            rPr.append(rFonts)
            sz = OxmlElement('w:sz'); sz.set(qn('w:val'), sz_val); rPr.append(sz)
            szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), sz_val); rPr.append(szCs)
            if bold_flag:
                rPr.append(OxmlElement('w:b'))
                rPr.append(OxmlElement('w:bCs'))
            rPr.append(OxmlElement('w:rtl'))
        else:
            rFonts.set(qn('w:ascii'), font_en); rFonts.set(qn('w:hAnsi'), font_en); rFonts.set(qn('w:cs'), font_fa)
            rPr.append(rFonts)
            sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '26' if level == 0 else sz_val); rPr.append(sz)
            szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), sz_val); rPr.append(szCs)
            if bold_flag:
                rPr.append(OxmlElement('w:b'))
                rPr.append(OxmlElement('w:bCs'))
    return p

def make_reference(doc, num_str, eng_text=None, fa_text=None):
    """
    CRITICAL RULE:
    - English references MUST NOT have w:bidi on the paragraph.
    - Persian references MUST have w:bidi on the paragraph.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pPr = p._element.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '0')
    sp.set(qn('w:after'), '100')
    sp.set(qn('w:line'), '360')
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)
    
    if eng_text:
        r = p.add_run(f"{num_str}. {eng_text}")
        rPr = r._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'B Nazanin')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '24'); rPr.append(sz)
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '24'); rPr.append(szCs)
    else:
        if pPr.find(qn('w:bidi')) is None:
            pPr.append(OxmlElement('w:bidi'))
        r = p.add_run(f"{num_str}. {fa_text}")
        rPr = r._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'B Nazanin')
        rFonts.set(qn('w:hAnsi'), 'B Nazanin')
        rFonts.set(qn('w:cs'), 'B Nazanin')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '26'); rPr.append(sz)
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '26'); rPr.append(szCs)
        rPr.append(OxmlElement('w:rtl'))
    return p
```

---

## 5. Word COM Engine Automation (TOC Update & PDF Export)

Never rely on third-party HTML converters to export Persian Word documents to PDF. Use Microsoft Word's native COM interface via `pywin32` (`win32com.client`). This ensures 100% native font rendering, exact hyphenation, perfect RTL rendering, and recalculates the Table of Contents page numbers accurately.

```python
import win32com.client
import os

def finalize_docx_and_export_pdf(docx_path, pdf_path):
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    
    os.system("taskkill /F /IM WINWORD.EXE >nul 2>&1")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    try:
        doc = word.Documents.Open(abs_docx)
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Save()
        doc.ExportAsFixedFormat(abs_pdf, 17) # 17 = wdExportFormatPDF
        doc.Close()
    finally:
        word.Quit()
```

---

## 6. Verification & Quality Checklist

1. **Reference BIDI Audit:**
   Verify that all English references have `bidi=False` and all Persian references have `bidi=True`.
2. **Formula Parenthesis Audit:**
   Verify that variables in formulas use pure Latin identifiers (e.g. `M(acid)`) rather than mixed Persian in parentheses.
3. **Heading Alignment Audit:**
   Verify that `Heading 1` and `Heading 2` have `WD_ALIGN_PARAGRAPH.JUSTIFY` for consistent right-margin alignment.
4. **TOC Page Numbers:**
   Verify that all chapter page numbers in the Table of Contents match the physical PDF page numbers.
