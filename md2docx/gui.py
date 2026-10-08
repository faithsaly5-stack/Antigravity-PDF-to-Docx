"""
Graphical User Interface (GUI) for Perfect Markdown to DOCX Converter.
Built with CustomTkinter for a modern, sleek Windows native experience.
"""

import os
import sys
import time
import threading
import subprocess
import shutil
from typing import List, Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox

from .config import TypographyConfig
from .converter import MarkdownToDocx

# Set appearance and theme
ctk.set_appearance_mode("System")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("blue")


class MarkdownToDocxGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Markdown to DOCX Studio — موتور پیشرفته تبدیل مارک‌داون")
        self.geometry("860, 720")
        self.minsize(760, 640)

        # State variables
        self.selected_files: List[str] = []
        self.output_directory: str = ""
        self.last_docx_path: Optional[str] = None
        self.last_pdf_path: Optional[str] = None
        self.is_converting = False

        self._build_ui()

    def _build_ui(self):
        # Grid layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, corner_radius=12, fg_color=("gray90", "gray17"))
        header_frame.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="🚀 Markdown to DOCX & PDF Studio",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
        )
        title_lbl.grid(row=0, column=0, padx=15, pady=(10, 2), sticky="w")

        subtitle_lbl = ctk.CTkLabel(
            header_frame,
            text="تایپوگرافی دوزبانه استاندارد فارسی، فرمول‌های بومی آفیس (OMML)، جداول راست‌به‌چپ و اعتبارسنجی OpenXML",
            font=ctk.CTkFont(family="Vazirmatn", size=13),
            text_color=("gray40", "gray70"),
        )
        subtitle_lbl.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        # Dark/Light Mode switch
        theme_switch = ctk.CTkSwitch(
            header_frame,
            text="حالت تیره (Dark)",
            command=self._toggle_theme,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        theme_switch.grid(row=0, column=1, rowspan=2, padx=15, pady=10, sticky="e")
        if ctk.get_appearance_mode() == "Dark":
            theme_switch.select()

        # 2. File Selection Card
        file_card = ctk.CTkFrame(self, corner_radius=12)
        file_card.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        file_card.grid_columnconfigure(1, weight=1)

        file_icon_lbl = ctk.CTkLabel(
            file_card, text="📄 فایل‌های ورودی:", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        )
        file_icon_lbl.grid(row=0, column=0, padx=15, pady=12, sticky="w")

        self.file_entry = ctk.CTkEntry(
            file_card,
            placeholder_text="یک یا چند فایل Markdown انتخاب کنید...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.file_entry.grid(row=0, column=1, padx=(0, 10), pady=12, sticky="ew")

        browse_btn = ctk.CTkButton(
            file_card,
            text="انتخاب فایل‌ها (Browse)",
            command=self._select_files,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            width=140,
        )
        browse_btn.grid(row=0, column=2, padx=(0, 15), pady=12)

        # Output Dir Row
        out_lbl = ctk.CTkLabel(
            file_card, text="📁 پوشه خروجی:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        )
        out_lbl.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="w")

        self.out_dir_entry = ctk.CTkEntry(
            file_card,
            placeholder_text="همان پوشه فایل ورودی (پیش‌فرض)",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.out_dir_entry.grid(row=1, column=1, padx=(0, 10), pady=(0, 12), sticky="ew")

        browse_out_btn = ctk.CTkButton(
            file_card,
            text="تغییر پوشه...",
            command=self._select_output_dir,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            width=140,
        )
        browse_out_btn.grid(row=1, column=2, padx=(0, 15), pady=(0, 12))

        # 3. Settings Card
        settings_card = ctk.CTkFrame(self, corner_radius=12)
        settings_card.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        settings_card.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Switches
        self.academic_var = ctk.BooleanVar(value=False)
        self.academic_switch = ctk.CTkSwitch(
            settings_card,
            text="قالب رسمی دانشگاهی (طرح جلد + فهرست)",
            variable=self.academic_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.academic_switch.grid(row=0, column=0, columnspan=2, padx=15, pady=10, sticky="w")

        self.pdf_var = ctk.BooleanVar(value=True)
        self.pdf_switch = ctk.CTkSwitch(
            settings_card,
            text="صدور مستقیم PDF با موتور Microsoft Word",
            variable=self.pdf_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.pdf_switch.grid(row=0, column=2, columnspan=2, padx=15, pady=10, sticky="w")

        self.validate_var = ctk.BooleanVar(value=False)
        self.validate_switch = ctk.CTkSwitch(
            settings_card,
            text="اعتبارسنجی ساختار با OfficeCLI (اختیاری)",
            variable=self.validate_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.validate_switch.grid(row=1, column=0, columnspan=2, padx=15, pady=(0, 12), sticky="w")

        # Fonts Dropdowns
        font_lbl = ctk.CTkLabel(
            settings_card, text="فونت متن فارسی:", font=ctk.CTkFont(family="Segoe UI", size=12)
        )
        font_lbl.grid(row=1, column=2, padx=(15, 5), pady=(0, 12), sticky="w")

        self.font_combo = ctk.CTkComboBox(
            settings_card,
            values=["B Nazanin", "Vazirmatn", "IRANYekan", "B Mitra", "Sahel"],
            width=130,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.font_combo.set("B Nazanin")
        self.font_combo.grid(row=1, column=3, padx=(0, 15), pady=(0, 12), sticky="w")

        # 4. Action & Log Area
        action_frame = ctk.CTkFrame(self, corner_radius=12)
        action_frame.grid(row=3, column=0, padx=20, pady=(10, 15), sticky="nsew")
        action_frame.grid_columnconfigure(0, weight=1)
        action_frame.grid_rowconfigure(2, weight=1)

        # Action Buttons Row
        btn_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_row.grid(row=0, column=0, padx=15, pady=(12, 8), sticky="ew")
        btn_row.grid_columnconfigure(0, weight=1)

        self.convert_btn = ctk.CTkButton(
            btn_row,
            text="⚡ شروع تبدیل به Word (Convert to DOCX)",
            command=self._start_conversion_thread,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            height=42,
            fg_color="#1D4ED8",
            hover_color="#1E40AF",
        )
        self.convert_btn.grid(row=0, column=0, sticky="ew")

        # Progress bar & Status
        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.grid(row=1, column=0, padx=15, pady=4, sticky="ew")
        self.progress_bar.set(0.0)

        # Console Text Box
        self.log_box = ctk.CTkTextbox(
            action_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
            corner_radius=8,
        )
        self.log_box.grid(row=2, column=0, padx=15, pady=(6, 10), sticky="nsew")
        self._log("آماده پردازش. لطفاً فایل(های) Markdown را انتخاب و روی دکمه شروع کلیک کنید.")

        # Quick action buttons row (Open Docx, Open PDF, Open Folder)
        self.bottom_bar = ctk.CTkFrame(action_frame, fg_color="transparent")
        self.bottom_bar.grid(row=3, column=0, padx=15, pady=(0, 12), sticky="ew")
        self.bottom_bar.grid_columnconfigure((0, 1, 2), weight=1)

        self.open_docx_btn = ctk.CTkButton(
            self.bottom_bar,
            text="📝 باز کردن در Word",
            command=self._open_docx,
            state="disabled",
            fg_color="#059669",
            hover_color="#047857",
        )
        self.open_docx_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.open_pdf_btn = ctk.CTkButton(
            self.bottom_bar,
            text="📕 مشاهده PDF نهایی",
            command=self._open_pdf,
            state="disabled",
            fg_color="#DC2626",
            hover_color="#B91C1C",
        )
        self.open_pdf_btn.grid(row=0, column=1, padx=5, sticky="ew")

        self.open_folder_btn = ctk.CTkButton(
            self.bottom_bar,
            text="📂 پوشه خروجی (Open Folder)",
            command=self._open_folder,
            state="disabled",
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
        )
        self.open_folder_btn.grid(row=0, column=2, padx=(5, 0), sticky="ew")

    def _toggle_theme(self):
        current = ctk.get_appearance_mode()
        ctk.set_appearance_mode("Light" if current == "Dark" else "Dark")

    def _select_files(self):
        files = filedialog.askopenfilenames(
            title="انتخاب فایل‌های Markdown",
            filetypes=[
                ("Markdown Files", "*.md *.markdown *.mdown *.mkd"),
                ("Text Files", "*.txt"),
                ("All Files", "*.*"),
            ],
        )
        if files:
            self.selected_files = list(files)
            if len(self.selected_files) == 1:
                self.file_entry.delete(0, "end")
                self.file_entry.insert(0, self.selected_files[0])
            else:
                self.file_entry.delete(0, "end")
                self.file_entry.insert(0, f"({len(self.selected_files)} فایل انتخاب شد)")

            # Auto-suggest output dir from first file
            if not self.output_directory:
                first_dir = os.path.dirname(os.path.abspath(self.selected_files[0]))
                self.out_dir_entry.delete(0, "end")
                self.out_dir_entry.insert(0, first_dir)
                self.output_directory = first_dir

            self._log(f"انتخاب شد: {len(self.selected_files)} فایل")

    def _select_output_dir(self):
        out_dir = filedialog.askdirectory(title="انتخاب پوشه خروجی")
        if out_dir:
            self.output_directory = out_dir
            self.out_dir_entry.delete(0, "end")
            self.out_dir_entry.insert(0, out_dir)

    def _log(self, message: str):
        t_str = time.strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{t_str}] {message}\n")
        self.log_box.see("end")

    def _start_conversion_thread(self):
        if not self.selected_files:
            raw_input = self.file_entry.get().strip()
            if raw_input and os.path.exists(raw_input):
                self.selected_files = [raw_input]
            else:
                messagebox.showwarning("خطا", "لطفاً ابتدا حداقل یک فایل Markdown را انتخاب کنید.")
                return

        if self.is_converting:
            return

        self.is_converting = True
        self.convert_btn.configure(state="disabled", text="⏳ در حال تبدیل...")
        self.progress_bar.start()

        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        try:
            total_files = len(self.selected_files)
            target_out_dir = self.out_dir_entry.get().strip()
            use_academic = self.academic_var.get()
            export_pdf = self.pdf_var.get()
            validate = self.validate_var.get()
            font_fa = self.font_combo.get()

            # Resolve template if requested
            template_path = None
            if use_academic:
                default_tmpl = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "templates", "default_academic.docx"
                )
                if os.path.exists(default_tmpl):
                    template_path = default_tmpl

            cfg = TypographyConfig(
                font_fa_body=font_fa,
                font_fa_heading="B Titr",
            )
            converter = MarkdownToDocx(config=cfg)

            for idx, file_path in enumerate(self.selected_files):
                abs_file = os.path.abspath(file_path)
                file_name = os.path.basename(abs_file)
                base_name = os.path.splitext(file_name)[0]

                out_folder = target_out_dir if target_out_dir else os.path.dirname(abs_file)
                os.makedirs(out_folder, exist_ok=True)

                out_docx = os.path.join(out_folder, f"{base_name}.docx")
                out_pdf = os.path.join(out_folder, f"{base_name}.pdf") if export_pdf else None

                self._log(f"({idx+1}/{total_files}) در حال پردازش: {file_name} ...")
                t0 = time.time()

                res = converter.convert_file(
                    input_file=abs_file,
                    output_docx=out_docx,
                    output_pdf=out_pdf,
                    template=template_path,
                    export_pdf=export_pdf,
                )
                t_elapsed = time.time() - t0

                stats = res["stats"]
                self._log(
                    f"✔ {file_name} در {t_elapsed:.2f} ثانیه تبدیل شد! "
                    f"(تیترها: {stats['headings']} | جداول: {stats['tables']} | فرمول‌ها: {stats['math_blocks']})"
                )

                self.last_docx_path = res["docx_path"]
                self.last_pdf_path = res["pdf_path"]

                # Optional validation
                if validate:
                    if shutil.which("officecli"):
                        self._log("در حال اعتبارسنجی ساختار با OfficeCLI...")
                        try:
                            v_res = subprocess.run(
                                ["officecli", "validate", out_docx],
                                capture_output=True,
                                text=True,
                                check=False,
                            )
                            if v_res.returncode == 0:
                                self._log("✔ اعتبارسنجی موفق: سند ۱۰۰٪ با استاندارد OpenXML سازگار است.")
                            else:
                                self._log(f"نکته اعتبارسنجی: {v_res.stdout[:150]}")
                        except Exception:
                            pass
                    else:
                        self._log("توجه: ابزار اختیاری OfficeCLI نصب نیست (اعتبارسنجی رد شد).")

            self._log("🎉 کلیه عملیات تبدیل با موفقیت به پایان رسید!")
            self.after(0, self._on_conversion_complete)

        except Exception as e:
            self._log(f"❌ خطا در فرآیند تبدیل: {str(e)}")
            self.after(0, lambda: messagebox.showerror("خطا", str(e)))
        finally:
            self.is_converting = False
            self.after(0, self._reset_ui_after_conversion)

    def _on_conversion_complete(self):
        if self.last_docx_path and os.path.exists(self.last_docx_path):
            self.open_docx_btn.configure(state="normal")
            self.open_folder_btn.configure(state="normal")
        if self.last_pdf_path and os.path.exists(self.last_pdf_path):
            self.open_pdf_btn.configure(state="normal")

    def _reset_ui_after_conversion(self):
        self.progress_bar.stop()
        self.progress_bar.set(1.0)
        self.convert_btn.configure(state="normal", text="⚡ شروع تبدیل به Word (Convert to DOCX)")

    def _open_docx(self):
        if self.last_docx_path and os.path.exists(self.last_docx_path):
            os.startfile(self.last_docx_path)

    def _open_pdf(self):
        if self.last_pdf_path and os.path.exists(self.last_pdf_path):
            os.startfile(self.last_pdf_path)

    def _open_folder(self):
        folder = self.out_dir_entry.get().strip()
        if not folder and self.last_docx_path:
            folder = os.path.dirname(self.last_docx_path)
        if folder and os.path.exists(folder):
            os.startfile(folder)


def launch_gui():
    app = MarkdownToDocxGUI()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
