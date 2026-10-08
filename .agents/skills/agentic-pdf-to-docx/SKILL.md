---
name: agentic-pdf-to-docx
description: >-
  Use this skill whenever the user wants to convert, extract, or OCR a long PDF (book, pamphlet, research paper, scanned notes)
  into Markdown and then compile it into a publication-quality Microsoft Word (.docx) and PDF document using the Antigravity AI Agent.
  Implements a high-speed, zero-context-bloat pipeline: Multithreaded 1050px JPEG caching, styling reference generation,
  autonomous sequential OCR loop, and one-click Persian/RTL academic Word compilation.
---

# Agentic Bulk PDF to DOCX & PDF Pipeline

This skill guides the AI Agent to autonomously convert multi-page PDF documents (in Persian, English, or mixed-script) into clean Markdown and compile them into flawless Microsoft Word (`.docx`) and PDF documents.

---

## 🎯 The Core Philosophy: Why Other Tools Fail

1. **Context Bloat Problem**: Sending 50+ pages of text into the agent's memory window causes memory saturation, massive slowdowns, hallucinated outputs, and crashes.
2. **The RTL & Math Dilemma**: Standard OCR engines invert Persian parentheses (`(Cl⁻)` -> `)Cl⁻(`) and output broken LaTeX or low-res math images.
3. **The Solution**:
   - **Phase 1**: Render pages to lightweight, standardized 1050px JPEGs via multithreading (`scripts/pdf_to_cache.py`).
   - **Phase 2**: Create a golden `styling_reference.md` from a single sample page.
   - **Phase 3**: Autonomous loop (`/goal`) extracting one page at a time into `temp_page.md` and appending it to `document.md` via Python CLI (keeping the agent's context 100% clean).
   - **Phase 4**: Compile `document.md` into Word & PDF via `md2docx` with native OMML equations, B Nazanin/B Titr typography, and APA references.

---

## 🛠️ Step-by-Step Autonomous Workflow

### Step 1: Pre-process PDF to Image Cache
Run the high-speed cache generator in the workspace:
```bash
python scripts/pdf_to_cache.py "<filename.pdf>"
```
- This generates optimized 1050px JPEGs in the `page_cache/` folder.
- Speed: ~25-50 pages per second with multi-threading.
- Average image size: ~70-120 KB per page.

### Step 2: Extract Sample Page & Create `styling_reference.md`
1. Read the first representative image: `page_cache/page_001.jpg` (or first content page).
2. Transcribe and format the text with strict Markdown:
   - Proper heading hierarchy (`# `, `## `, `### `) with space after `#`.
   - Blockquotes (`> [!NOTE]` or `> `) for boxed tips or warnings.
   - Bold text (`**text**`) for definitions and emphasis.
   - LaTeX equations (`$...$` inline, `$$...$$` block).
   - Markdown tables (`| ... |`) for tabular data.
3. Save the formatted result into `styling_reference.md`.
4. Prompt the user:
   > *"I have created `styling_reference.md` based on page 1. Please review the style, and I will autonomously extract the remaining pages."*

### Step 3: Autonomous Bulk Extraction Loop (`/goal`)
Use `/goal` to loop through all sequential pages in `page_cache/` without stopping:

```bash
/goal Resume OCR extraction from page_cache into document.md
```

**Strict Agent Execution Rules for Step 3:**
1. **Never load `document.md` into context**. Inspect only the last few lines to know which page was processed:
   ```powershell
   Get-Content document.md -Tail 20
   ```
2. **Read ONLY the next sequential image** (e.g., `page_cache/page_002.jpg`).
3. **Extract & Format**: Adhere strictly to the conventions established in `styling_reference.md`.
4. **Buffer Output**: Save the current page to `temp_page.md`.
5. **Append cleanly via Python**:
   ```bash
   python -c "with open('temp_page.md', 'r', encoding='utf-8') as src, open('document.md', 'a', encoding='utf-8') as dst: dst.write('\n\n' + src.read().strip() + '\n')"
   ```
6. **Wipe buffer**: Clear `temp_page.md` and proceed immediately to the next image.

### Step 4: One-Click Compile to Academic Word (.docx) & PDF
Once extraction is complete, compile the resulting Markdown file using the built-in `md2docx` engine:

```bash
# Academic Mode: Includes cover page, dynamic TOC, and vector headers
python -m md2docx document.md -o "final_report.docx" --academic --pdf
```

The resulting `.docx` and `.pdf` files will have:
- Perfect right-to-left Persian alignment (`B Nazanin`, `B Titr`).
- True vector OMML equations (no pixelated images).
- Clean Booktabs tables with repeating headers.
- Automatic Table of Contents (`w:sdt`).

---

## 📋 Quality Verification Checklist
- [ ] No flipped parentheses in chemical/mathematical terms (e.g. `M(acid)` or `(Cl⁻)`).
- [ ] English references isolated from Persian `bidi` tags.
- [ ] Document validated via `officecli validate final_report.docx`.
- [ ] Temporary files (`temp_page.md`, `page_cache/`) cleaned up when finished.
