---
name: persian-word-edit
description: >-
  Use this skill whenever the user asks to modify, edit, or replace text in a Microsoft Word (.docx) document that contains Persian (Farsi), Arabic, or any RTL complex script text with specific fonts (e.g., B Nazanin). This skill ensures that complex script font properties and RTL layouts are strictly preserved during automated edits.
---

# Editing Persian / RTL Word Documents

When you are tasked with automating edits to a Microsoft Word document containing Persian or RTL text, it is critical that you preserve the document's original formatting (like complex script fonts such as B Nazanin, B Lotus, and right-to-left layout constraints).

## The Problem
Using standard Python libraries (like `python-docx`) to replace whole paragraphs usually overrides or wipes out the underlying OpenXML Run Properties (`w:rPr`) and Paragraph Properties (`w:pPr`). This strips away the Complex Script (CS) font settings and RTL alignment, permanently destroying the document's native Persian formatting.

## The Solution
**ALWAYS use the `officecli` tool** to perform targeted string replacements instead of rewriting whole paragraphs using Python scripts. `officecli` operates safely on the DOM level and leaves the styling nodes completely intact.

## Workflow

1. **Dump and Inspect the Document**:
   Use `officecli view <filename> text > dump.txt` to dump the text alongside their stable `paraId` identifiers.
   *Example:* `[/body/p[@paraId=71747998]] بررسی تنوع ژنتیکی...`

2. **Perform Targeted String Replacements**:
   Use the `officecli set` command combined with `--find` and `--replace`.
   ```bash
   officecli set document.docx "/body/p[@paraId=71747998]" --find "Old Text" --replace "New Text"
   ```
   *Note: Because `officecli` replaces the text node directly within its run, the Persian font, size, and RTL direction are 100% preserved.*

3. **Restructuring (Bullets to Text, Adding Table Rows)**:
   - To remove unwanted bullet points or paragraphs, use `officecli remove document.docx "/body/p[@paraId=...]"`.
   - To merge text, apply the full combined text replacement to one paragraph via `--find/--replace`, then use `officecli remove` on the remaining redundant paragraphs.
   - To duplicate a row in a table while keeping its exact styling, use `officecli add`:
     ```bash
     officecli add document.docx "/body/tbl[1]" --from "/body/tbl[1]/tr[4]"
     ```
     Then, use `officecli set` to modify the newly cloned row.

4. **Scripting the Edits**:
   If there are many edits, write them sequentially in a PowerShell (`.ps1`) script and execute it to apply all changes efficiently and accurately in one go.
