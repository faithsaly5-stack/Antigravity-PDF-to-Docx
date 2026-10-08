# 🚀 Agentic Bulk PDF-to-DOCX & PDF Pipeline (Bilingual Guide)
## راهنمای جامع خط لوله استخراج خودکار کتاب و جزوه PDF به Word و PDF با هوش مصنوعی

> **English Summary:** This repository provides an end-to-end autonomous AI agent pipeline to extract large, multi-page PDFs (books, lecture notes, lab manuals) into Markdown with **zero context bloat**, and compile them into publication-quality Microsoft Word (`.docx`) and PDF documents with native Persian/RTL typography, OMML math formulas, and dynamic Table of Contents.
> 
> **خلاصه فارسی:** این پروژه یک راهکار کامل و هوشمند برای تبدیل خودکار اسناد حجیم و طولانی PDF (کتاب، جزوه، اسناد دست‌نویس و مقالات) به مارک‌داون بدون اشباع حافظه هوش مصنوعی (Zero Context Bloat) و سپس تولید خودکار فایل نهایی Word (`.docx`) و `PDF` دانشگاهی با فونت‌های استاندارد فارسی (`B Nazanin`، `B Titr`)، فرمول‌های بومی Word و فهرست مطالب پویا است.

---

## 🌟 چرا این روش بهترین راهکار موجود است؟ (Why This Method?)

| مشکل متداول مدل‌های عادی بدون خط لوله ایجنتی (Claude, ChatGPT, Gemini, DeepSeek) | راهکار هوشمند این خط لوله (Our Agentic Pipeline) |
| :--- | :--- |
| **اشباع حافظه (Context Bloat):** پس از ۱۵-۲۰ صفحه چت طولانی شده، هوش مصنوعی کند می‌شود یا متوقف می‌گردد. | **استخراج تک‌صفحه‌ای با الحاق پایتون:** متن در فایل موقت نوشته شده و با اسکریپت الحاق می‌شود؛ پنجره حافظه مدل همیشه ۱۰۰٪ سبک و سریع می‌ماند. |
| **بهم‌ریختگی پرانتزها و فرمول‌ها:** در متن دوزبانه پرانتزها وارونه شده (`(Cl⁻)` -> `)Cl⁻(`) و علامت‌های منفی جابجا می‌شوند. | **موتور قطعه‌بندی هوشمند (`md2docx`):** فرمول‌ها، نمادهای شیمیایی و متون انگلیسی بدون `w:rtl` پردازش شده و هرگز وارونه نمی‌شوند. |
| **فرمول‌های ریاضی بی‌کیفیت:** فرمول‌ها تبدیل به عکس‌های تار و شطرنجی می‌شوند. | **معادلات بومی آفیس (OMML):** تبدیل مستقیم LaTeX به معادلات برداری و قابل ویرایش Word. |
| **سبک‌بندی ناپیوسته:** صفحات آغازین با تیتر و صفحات بعدی با فونت و فرمت متفاوت استخراج می‌شوند. | **الگوی مبنا (`styling_reference.md`):** صفحه نمونه استخراج شده و تمام صفحات بعدی موظف به پیروی دقیق از همان فرمت هستند. |

---

# 🇮🇷 راهنمای فارسی (Persian Guide)

## روش‌های استفاده (۳ مسیر ساده و کاربردی)

### 🟢 مسیر اول: خودکار با عامل هوشمند Antigravity (ساده‌ترین روش)
پروژه مجهز به یک **مهارت هوشمند (Agent Skill)** به نام `agentic-pdf-to-docx` است:
1. فایل PDF خود را در پوشه پروژه کپی کنید.
2. نرم‌افزار **Google Antigravity IDE** را در این پوشه باز کنید.
3. در پنجره چت کافی است بنویسید:
   ```text
   فایل book.pdf را به صورت کامل به ورد تبدیل کن
   ```
   یا به انگلیسی:
   ```text
   /goal Convert book.pdf into Word and PDF
   ```
4. هوش مصنوعی به صورت کاملاً خودکار:
   - تصاویر سبک ۱۰۵۰ پیکسلی را در پوشه `page_cache` می‌سازد.
   - صفحه اول را تبدیل به `styling_reference.md` می‌کند و تاییدیه می‌گیرد.
   - با فرمان `/goal` تمام صفحات را بدون توقف استخراج کرده و در `document.md` تجمیع می‌کند.
   - در پایان با موتور `md2docx` خروجی‌های شکیل `.docx` و `.pdf` را تحویل می‌دهد!

---

### 🟡 مسیر دوم: اجرای اسکریپت آماده کش + پرامپت
اگر می‌خواهید تبدیل صفحات پی‌دی‌اف به تصاویر را خودتان با سرعت فوق‌العاده بالا انجام دهید:

1. اجرای اسکریپت با ترمینال (۱۰۰ صفحه در ۳ ثانیه):
   ```bash
   python scripts/pdf_to_cache.py "my_book.pdf"
   ```
2. ارسال **پرامپت شماره ۲** به هوش مصنوعی برای شروع استخراج خودکار تمام صفحات.

---

### 🟠 مسیر سوم: پرامپت‌های مستقیم برای هر هوش مصنوعی (Claude 3.7 Sonnet, GPT-4o / o1, Gemini 2.0 Flash / Pro, DeepSeek-R1, Cursor)

#### 📝 پرامپت مرحله ۱ (آماده‌سازی کش و صفحه الگو - انگلیسی جهت درک بهتر ایجنت):
```text
Prepare the environment by converting a PDF into optimized JPEGs, and then extract a single sample page to establish the ground-truth Markdown styling reference for a bulk OCR pipeline.

### 📝 Phase 1: PDF to Image Conversion
1. Identify the target .pdf file in the current working directory.
2. Delete the folders page_cache and any previous test images if they exist.
3. Run python scripts/pdf_to_cache.py (or create a script using pymupdf that renders each page to exactly 1050px width, quality=70, saving to page_cache/page_001.jpg, etc.).
4. Wait until all images are successfully saved.

### 📝 Phase 2: Create Styling Reference
5. Open the first representative page (e.g. page_cache/page_001.jpg).
6. Transcribe the text with maximum accuracy, proper RTL/Persian alignment, LaTeX equations ($...$), and tables.
7. Apply strict Markdown rules:
   - Headings with proper levels (# , ## ) and a space after #.
   - Callout blocks (> [!NOTE] or > ) for tips/warnings.
   - Bold text (**text**) for terms and emphasis.
8. Save this output into styling_reference.md and request confirmation from the user.
```

#### 🚀 پرامپت مرحله ۲ (استخراج خودکار و بدون وقفه کل صفحات - انگلیسی):
```text
/goal Resume the strict OCR text extraction of the pictures in the `page_cache` folder into `document.md` efficiently, without ANY image cropping or context bloat.

### 📝 Execution Steps
1. Use a terminal command (e.g., Get-Content document.md -Tail 20 or tail -n 20 document.md) to inspect ONLY the end of document.md to identify the last processed page.
2. Read styling_reference.md to follow exact styling conventions.
3. DO NOT load or read the full document.md into your context to prevent memory bloat.
4. Select the VERY NEXT sequential image from the page_cache folder.
5. Extract all text, equations, and tables accurately, formatting according to styling_reference.md.
6. Write this output into temp_page.md.
7. Run this Python terminal command to append the content:
   python -c "with open('temp_page.md', 'r', encoding='utf-8') as src, open('document.md', 'a', encoding='utf-8') as dst: dst.write('\n\n' + src.read().strip() + '\n'); print('Appended successfully')"
8. Clear temp_page.md, reset contextual memory of the current page, and loop back to Step 4 until all pages are extracted.
```

---

### 🏁 مرحله نهایی: تبدیل ۱-کلیکی مارک‌داون به Word و PDF

پس از پایان استخراج، فایل `document.md` را با یکی از روش‌های زیر به فایل ورد شکیل تبدیل کنید:

1. **با خط فرمان (CLI):**
   ```bash
   # حالت رسمی دانشگاهی (همراه با طرح جلد، فهرست مطالب و صدور مستقیم PDF)
   python -m md2docx document.md -o "Final_Document.docx" --academic --pdf
   ```

2. **با محیط گرافیکی ویندوز (GUI):**
   روی [`Start_Markdown_Studio.bat`](Start_Markdown_Studio.bat) دابل کلیک کنید، فایل `document.md` را انتخاب کرده و روی **شروع تبدیل** کلیک کنید!

---

# 🇬🇧 English Guide

## 3 Flexible Usage Methods

### 🟢 Method 1: Autonomous Antigravity Agent Skill (Recommended)
This repository includes a native workspace skill (`.agents/skills/agentic-pdf-to-docx/`):
1. Copy your `.pdf` file into the repository folder.
2. Open **Google Antigravity IDE** in this folder.
3. Prompt the agent in the chat:
   ```text
   /goal Convert my_book.pdf to Word and PDF
   ```
4. The agent handles everything autonomously:
   - Renders 1050px optimized JPEGs using `scripts/pdf_to_cache.py`.
   - Generates `styling_reference.md` from page 1.
   - Runs the autonomous `/goal` loop without memory bloat.
   - Compiles the final `.docx` and `.pdf` with academic layout and native formulas.

---

### 🟡 Method 2: Fast Pre-caching Script + Prompts
Run the multithreaded caching utility:
```bash
python scripts/pdf_to_cache.py "my_document.pdf"
```
Then execute **Prompt 2** in your AI agent to run the extraction loop.

---

### 🟠 Method 3: Copy-Paste Prompts (For Any AI Agent: Claude 3.7 Sonnet, GPT-4o / o1, Gemini 2.0 Flash / Pro, DeepSeek-R1, Cursor)

#### 📝 Prompt 1: Preparation & Styling Setup
```text
Prepare the environment by converting a PDF into optimized JPEGs, and then extract a single sample page to establish the ground-truth Markdown styling reference for a bulk OCR pipeline.

### 📝 Phase 1: PDF to Image Conversion
1. Identify the target .pdf file in the current working directory.
2. Delete the folders page_cache and any previous test images if they exist.
3. Run python scripts/pdf_to_cache.py (or create a script using pymupdf that renders each page to exactly 1050px width, quality=70, saving to page_cache/page_001.jpg, etc.).
4. Wait until all images are successfully saved.

### 📝 Phase 2: Create Styling Reference
5. Open the first representative page (e.g. page_cache/page_001.jpg).
6. Transcribe the text with maximum accuracy, proper RTL/Persian alignment, LaTeX equations ($...$), and tables.
7. Apply strict Markdown rules:
   - Headings with proper levels (# , ## ) and a space after #.
   - Callout blocks (> [!NOTE] or > ) for tips/warnings.
   - Bold text (**text**) for terms and emphasis.
8. Save this output into styling_reference.md and request confirmation from the user.
```

#### 🚀 Prompt 2: Autonomous Bulk Extraction Loop
```text
/goal Resume the strict OCR text extraction of the pictures in the `page_cache` folder into `document.md` efficiently, without ANY image cropping or context bloat.

### 📝 Execution Steps
1. Use a terminal command (e.g., Get-Content document.md -Tail 20 or tail -n 20 document.md) to inspect ONLY the end of document.md to identify the last processed page.
2. Read styling_reference.md to follow exact styling conventions.
3. DO NOT load or read the full document.md into your context to prevent memory bloat.
4. Select the VERY NEXT sequential image from the page_cache folder.
5. Extract all text, equations, and tables accurately, formatting according to styling_reference.md.
6. Write this output into temp_page.md.
7. Run this Python terminal command to append the content:
   python -c "with open('temp_page.md', 'r', encoding='utf-8') as src, open('document.md', 'a', encoding='utf-8') as dst: dst.write('\n\n' + src.read().strip() + '\n'); print('Appended successfully')"
8. Clear temp_page.md, reset contextual memory of the current page, and loop back to Step 4 until all pages are extracted.
```

---

### 🏁 Final Compilation Step: Markdown to DOCX & PDF

Compile `document.md` into an academic Word document:
```bash
python -m md2docx document.md -o "Final_Output.docx" --academic --pdf
```
Or launch the Windows GUI Studio via [`Start_Markdown_Studio.bat`](Start_Markdown_Studio.bat).
