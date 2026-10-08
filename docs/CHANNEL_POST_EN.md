# 📢 Ready-to-Publish Channel / Social Media Post (English Version)

---

🚀 **Tired of AI crashing when trying to convert long PDFs/books to Word? Here is the ultimate open-source solution!**

Extracting 50+ page PDFs with complex Persian/Arabic RTL text, chemical formulas, and math into clean Microsoft Word documents has always been broken:
❌ **Context Bloat:** After 15 pages, LLMs run out of context, hallucinate, or stop completely.
❌ **Flipped Parentheses:** Bidirectional text inverts parentheses, turning `(Cl⁻)` into `)Cl⁻(` and breaking variable names.
❌ **Blurry Formulas:** Math equations turn into low-res pixelated raster images.
❌ **Ugly Output:** No academic styles, no dynamic Table of Contents, broken tables.

We have open-sourced a complete, autonomous agentic pipeline that solves this end-to-end:

---

### 💡 How It Works

1. **Lightning-Fast Image Caching:** A multithreaded script (`scripts/pdf_to_cache.py`) converts entire PDF books into lightweight 1050px JPEGs in seconds (~35 pages/sec).
2. **Zero-Context-Bloat Agent Loop:** The AI agent extracts page 1 to create a `styling_reference.md` baseline. Then using `/goal`, it loops through pages one by one, writing to a temp buffer and appending directly via Python CLI. The agent's conversation window stays 100% fast and clean even for 500-page books!
3. **Dedicated Word & PDF Engine (`md2docx`):** Compiles Markdown directly into publication-quality DOCX & PDF:
   - ✅ Persian RTL Typography (`B Nazanin`, `B Titr`) with tokenized run generation (zero parenthesis flipping).
   - ✅ Native vector Office Math (OMML) converted directly from LaTeX.
   - ✅ Reference BIDI isolation for English citations (APA, IEEE) alongside Persian references.
   - ✅ Academic Booktabs tables with repeating headers on subsequent pages.
   - ✅ Academic cover page with universal crest and dynamic Word Table of Contents (`w:sdt`).

---

### 🛠️ Quick Start (in Google Antigravity IDE)

This repository includes a native workspace skill (`agentic-pdf-to-docx`):
1. Clone the repo and drop your PDF into the folder.
2. Open **Google Antigravity IDE**.
3. In chat, simply type:
   > *"/goal Convert textbook.pdf to Word and PDF"*
4. The agent handles caching, autonomous OCR extraction, and compiles the final `.docx` and `.pdf` files automatically!

Works with any frontier AI agent & model (Claude Sonnet 5.5, GPT-6 Astra / Sol, Gemini 4 Argon, DeepSeek-V4-Pro, Llama 4, Grok 4.7) via the included step-by-step bilingual prompt guide.

---

🔗 **GitHub Repository (Code, Templates, GUI & Guide):**
👉 [https://github.com/faithsaly5-stack/Antigravity-PDF-to-Docx](https://github.com/faithsaly5-stack/Antigravity-PDF-to-Docx)

⭐ Star the repo if you find it helpful!
