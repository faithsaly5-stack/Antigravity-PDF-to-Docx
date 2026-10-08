"""
Microsoft Word COM Automation Engine.
Updates Table of Contents, refreshes field codes, and exports to native PDF.
"""

import os
import sys
from typing import Optional


def update_toc_and_export_pdf(docx_path: str, pdf_path: Optional[str] = None) -> bool:
    """
    Automates Microsoft Word via COM to:
    1. Recalculate Table of Contents page numbers accurately.
    2. Recalculate field codes (e.g. PAGE, NUMPAGES).
    3. Save document and optionally export to 100% native PDF.
    """
    if sys.platform != "win32":
        return False

    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path) if pdf_path else None

    if not os.path.exists(abs_docx):
        raise FileNotFoundError(f"DOCX file not found: {abs_docx}")

    try:
        import win32com.client
    except ImportError:
        print("[WARNING] pywin32 is not installed. Skipping Word COM automation.")
        return False

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
    except Exception as e:
        print(f"[WARNING] Could not start Microsoft Word application: {e}")
        return False

    try:
        doc = word.Documents.Open(abs_docx)
        
        # Update all Table of Contents
        try:
            for toc in doc.TablesOfContents:
                toc.Update()
        except Exception:
            pass

        # Update all field codes
        try:
            for field in doc.Fields:
                field.Update()
        except Exception:
            pass

        doc.Save()

        # Export to PDF if requested
        if abs_pdf:
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(abs_pdf, 17)

        doc.Close()
        return True
    except Exception as e:
        print(f"[ERROR] Word automation error: {e}")
        return False
    finally:
        try:
            word.Quit()
        except Exception:
            pass
