"""
Configuration and Typography Settings for Perfect Markdown to DOCX Converter.
Adheres strictly to the Persian Academic Word Document Engineering standard.
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, Any


@dataclass
class TypographyConfig:
    # Fonts
    font_fa_body: str = "B Nazanin"
    font_fa_heading: str = "B Titr"
    font_en_body: str = "Times New Roman"
    font_en_heading: str = "Times New Roman"
    font_code: str = "Consolas"
    font_math: str = "Cambria Math"

    # Font Sizes (in points)
    size_title: float = 20.0
    size_subtitle: float = 16.0
    size_h1: float = 17.0
    size_h2: float = 14.0
    size_h3: float = 13.0
    size_h4: float = 12.0
    size_body_fa: float = 14.0
    size_body_en: float = 13.0
    size_table_header: float = 12.0
    size_table_cell: float = 13.0
    size_caption: float = 12.0
    size_code: float = 10.0
    size_ref_fa: float = 13.0
    size_ref_en: float = 12.0

    # Spacing (in points)
    space_before_h1: float = 14.0
    space_after_h1: float = 6.0
    space_before_h2: float = 10.0
    space_after_h2: float = 5.0
    space_before_h3: float = 8.0
    space_after_h3: float = 4.0
    space_before_body: float = 0.0
    space_after_body: float = 8.0
    line_spacing_multiplier: float = 1.5  # 1.5 lines = 360 dxa in OpenXML

    # Equation spacing
    space_before_equation: float = 8.0
    space_after_equation: float = 8.0

    # Image spacing
    space_before_image: float = 7.0
    space_after_image: float = 4.0
    space_after_caption: float = 8.0

    # Geometry (cm)
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0
    margin_left_cm: float = 2.0
    margin_right_cm: float = 2.0
    page_size: str = "A4"  # A4 or Letter


@dataclass
class ColorTheme:
    text_primary: str = "000000"
    text_muted: str = "4B5563"
    border_subtle: str = "CBD5E1"
    border_accent: str = "334155"
    table_header_bg: str = "F1F5F9"
    table_zebra_bg: str = "F8FAFC"
    code_bg: str = "F8F9FA"
    code_border: str = "E2E8F0"
    inline_code_bg: str = "F1F5F9"
    link_color: str = "1D4ED8"

    # Callout colors (border, background, title)
    callout_note_border: str = "2563EB"
    callout_note_bg: str = "EFF6FF"
    callout_tip_border: str = "16A34A"
    callout_tip_bg: str = "F0FDF4"
    callout_important_border: str = "9333EA"
    callout_important_bg: str = "FAF5FF"
    callout_warning_border: str = "D97706"
    callout_warning_bg: str = "FFFBEB"
    callout_caution_border: str = "DC2626"
    callout_caution_bg: str = "FEF2F2"
    callout_quote_border: str = "64748B"
    callout_quote_bg: str = "F8FAFC"


@dataclass
class DocumentMetadata:
    title: Optional[str] = None
    subtitle: Optional[str] = None
    authors: list[str] = field(default_factory=list)
    supervisors: list[str] = field(default_factory=list)
    date: Optional[str] = None
    institution: Optional[str] = None
    faculty: Optional[str] = None
    cover_page: bool = False
    toc: bool = False
    toc_title: str = "فهرست مطالب"
    lang: str = "fa"  # 'fa' or 'en'
    template_path: Optional[str] = None
    extra: Dict[str, Any] = field(default_factory=dict)
