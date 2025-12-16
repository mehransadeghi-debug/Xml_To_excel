# Xml_To_excel
Transform Excel or both 
--------------------------------------------------------------------------------------------------------------------------
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
from PIL import Image, ImageTk

# === فعال‌سازی High DPI ===
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass


class ExcelToXMLConverter:
    def __init__(self, root):
        self.logo_photo = None
        self.root = root
        self.root.title("⚡ Excel to XML Pro — نسخه اصلی")
        self.root.geometry("780x750")
        self.root.minsize(650, 550)
        self.root.configure(bg="#f0f4ff")

        self.colors = {
            "bg": "#f0f4ff",
            "card_bg": "#ffffff",
            "card_border": "#e2e8f0",
            "primary": "#4361ee",
            "primary_dark": "#3a56d4",
            "success": "#06d6a0",
            "success_dark": "#05b88f",
            "danger": "#ef476f",
            "danger_dark": "#d63a5a",
            "gray": "#7289da",
            "text": "#1a1a2e",
            "text_secondary": "#4e5d8f",
            "stat_bg": "#f8fafc",
            "stat_border": "#e2e8f0"
        }

        self.excel_path = tk.StringVar()
        self.original_logo = None
        self.logo_label = None
        self.loaded_df = None  # ✅ برای پیش‌نمایش داده

        self.rows_label = None
        self.cols_label = None
        self.file_label = None

        self.create_shadow_card()
        self.load_logo()

        # === عنوان ===
        title = tk.Label(
            self.content_frame,
            text="XML سامانه تبدیل فایل اکسل به ",
            font=("Segoe UI", 22, "bold"),
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        )
        title.pack(pady=(20, 5))

        subtitle = tk.Label(
            self.content_frame,
            text="پشتیبانی از فارسی، انگلیسی و تمام نسخه‌های اکسل (2007–2025)",
            font=("Segoe UI", 11),
            bg=self.colors["card_bg"],
            fg=self.colors["text_secondary"]
        )
        subtitle.pack(pady=(0, 20))

        # === فیلد فایل ===
        file_frame = tk.Frame(self.content_frame, bg=self.colors["card_bg"])
        file_frame.pack(padx=40, fill="x")

        tk.Label(file_frame, text="📁 فایل اکسل:", font=("Segoe UI", 11, "bold"), bg=self.colors["card_bg"]).pack(anchor="w")

        entry_bg = tk.Frame(file_frame, bg="#f1f5ff", highlightbackground="#cbd5e1", highlightthickness=1, bd=0)
        entry_bg.pack(fill="x", pady=8, ipady=3)

        self.path_entry = tk.Entry(
            entry_bg,
            textvariable=self.excel_path,
            state="readonly",
            font=("Segoe UI", 10),
            bg="#f1f5ff",
            fg=self.colors["text"],
            relief="flat",
            bd=0
        )
        self.path_entry.pack(fill="x", padx=10, pady=3)

        # دکمه‌ها
        btn_frame = tk.Frame(self.content_frame, bg=self.colors["card_bg"])
        btn_frame.pack(pady=15)

        self.browse_btn = self.create_modern_button(
            btn_frame, "🔍 انتخاب فایل", self.select_file,
            bg=self.colors["primary"], hover=self.colors["primary_dark"]
        )

        self.clear_btn = self.create_modern_button(
            btn_frame, "🗑️ پاک کردن", self.clear_file,
            bg=self.colors["gray"], hover=self.colors["text_secondary"]
        )

        # === بخش آمار ===
        self.create_stats_section()

        # === دکمه‌های اصلی ===
        action_frame = tk.Frame(self.content_frame, bg=self.colors["card_bg"])
        action_frame.pack(pady=20)

        self.convert_btn = self.create_modern_button(
            action_frame, "🚀  XML تبدیل به ", self.convert_to_xml,
            bg=self.colors["success"], hover=self.colors["success_dark"],
            font_size=13, bold=True, width=18
        )

        self.preview_btn = self.create_modern_button(  # ✅ دکمه پیش‌نمایش
            action_frame, "👁️ پیش‌نمایش داده", self.preview_data,
            bg="#7289da", hover="#5a6bc4",
            font_size=12, bold=False, width=18
        )

        self.exit_btn = self.create_modern_button(
            action_frame, "🚪 خروج", self.exit_app,
            bg=self.colors["danger"], hover=self.colors["danger_dark"],
            font_size=13, bold=True, width=12
        )

        footer = tk.Label(
            self.content_frame,
            text="کاربر ارشد سازمان 👩‍💼👨‍💼",
            font=("Segoe UI", 16),
            bg=self.colors["card_bg"],
            fg=self.colors["text_secondary"],
            justify="center"
        )
        footer.pack(side=tk.BOTTOM, pady=15)

    def create_stats_section(self):
        stats_frame = tk.Frame(self.content_frame, bg=self.colors["card_bg"])
        stats_frame.pack(pady=15, padx=40, fill="x")

        title = tk.Label(
            stats_frame,
            text="📊 آمار فایل اکسل",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        )
        title.pack(anchor="w", pady=(0, 8))

        stat_bg = tk.Frame(stats_frame, bg=self.colors["stat_bg"], bd=1, relief="solid",
                           highlightbackground=self.colors["stat_border"])
        stat_bg.pack(fill="x", padx=5)

        self.rows_label = tk.Label(stat_bg, text="• تعداد ردیف‌ها: —", font=("Segoe UI", 10), bg=self.colors["stat_bg"],
                                   fg=self.colors["text_secondary"], anchor="w", justify="left")
        self.rows_label.pack(fill="x", padx=15, pady=5)

        self.cols_label = tk.Label(stat_bg, text="• تعداد ستون‌ها: —", font=("Segoe UI", 10), bg=self.colors["stat_bg"],
                                   fg=self.colors["text_secondary"], anchor="w", justify="left")
        self.cols_label.pack(fill="x", padx=15, pady=5)

        self.file_label = tk.Label(stat_bg, text="• فایل: —", font=("Segoe UI", 10), bg=self.colors["stat_bg"],
                                   fg=self.colors["text_secondary"], anchor="w", justify="left")
        self.file_label.pack(fill="x", padx=15, pady=5)

    def update_stats(self, file_path, df):
        try:
            rows, cols = df.shape
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            size_mb = file_size / (1024 * 1024)
            size_str = f"{size_mb:.2f} MB" if size_mb > 1 else f"{file_size / 1024:.1f} KB"

            self.rows_label.config(text=f"• تعداد ردیف‌ها: {rows}")
            self.cols_label.config(text=f"• تعداد ستون‌ها: {cols}")
            self.file_label.config(text=f"• فایل: {file_name} ({size_str})")
        except:
            self.rows_label.config(text="• تعداد ردیف‌ها: —")
            self.cols_label.config(text="• تعداد ستون‌ها: —")
            self.file_label.config(text="• فایل: —")

    def clear_stats(self):
        if self.rows_label:
            self.rows_label.config(text="• تعداد ردیف‌ها: —")
        if self.cols_label:
            self.cols_label.config(text="• تعداد ستون‌ها: —")
        if self.file_label:
            self.file_label.config(text="• فایل: —")

    def create_shadow_card(self):
        shadow = tk.Frame(self.root, bg="#e0e7ff", bd=0)
        shadow.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.96, relheight=0.96)
        outer = tk.Frame(shadow, bg=self.colors["card_border"], bd=0)
        outer.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.99, relheight=0.99)
        self.content_frame = tk.Frame(outer, bg=self.colors["card_bg"], bd=0)
        self.content_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.99, relheight=0.98)

    def create_modern_button(self, parent, text, command, bg, hover, font_size=11, bold=False, width=None):
        font_style = ("Segoe UI", font_size, "bold" if bold else "normal")
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=font_style,
            bg=bg,
            fg="white",
            relief="flat",
            bd=0,
            padx=24,
            pady=10,
            cursor="hand2",
            width=width
        )
        btn.bind("<Enter>", lambda e: btn.config(bg=hover))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        btn.pack(side=tk.LEFT, padx=12)
        return btn

    def load_logo(self):
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            for name in ["logo.png", "logo.jpg"]:
                path = os.path.join(base_dir, name)
                if os.path.isfile(path):
                    img = Image.open(path)
                    img = img.convert("RGBA")
                    img.thumbnail((200, 100), Image.Resampling.LANCZOS)
                    self.original_logo = img
                    self.update_logo()
                    return
            spacer = tk.Label(self.content_frame, text="", bg=self.colors["card_bg"], height=2)
            spacer.pack(pady=10)
        except Exception as e:
            print(f"⚠️ خطا در لوگو: {e}")
            spacer = tk.Label(self.content_frame, text="", bg=self.colors["card_bg"], height=2)
            spacer.pack(pady=10)

    # noinspection PyTypeChecker
    def update_logo(self):
        if not self.original_logo:
            return
        width = min(220, max(140, self.root.winfo_width() // 5))
        ratio = width / self.original_logo.width
        height = int(self.original_logo.height * ratio)
        resized = self.original_logo.resize((width, height), Image.Resampling.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(resized)
        if self.logo_label:
            self.logo_label.config(image=self.logo_photo)
        else:
            self.logo_label = tk.Label(self.content_frame, image=self.logo_photo , bg=self.colors["card_bg"])
            self.logo_label.pack(pady=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="انتخاب فایل اکسل",
            filetypes=[
                ("فایل‌های اکسل", "*.xlsx *.xls"),
                ("اکسل نوین (.xlsx)", "*.xlsx"),
                ("اکسل قدیم (.xls)", "*.xls"),
                ("همه فایل‌ها", "*.*")
            ]
        )
        if file_path:
            self.excel_path.set(file_path)
            try:
                df = pd.read_excel(file_path, dtype=str, keep_default_na=False, na_values=[""])
                self.loaded_df = df  # ✅ ذخیره برای پیش‌نمایش
                self.update_stats(file_path, df)
            except Exception as e:
                messagebox.showerror("خطا", f"خطا در خواندن فایل:\n{str(e)}")
                self.clear_stats()
                self.loaded_df = None

    def clear_file(self):
        self.excel_path.set("")
        self.loaded_df = None  # ✅ پاک کردن داده
        self.clear_stats()
        messagebox.showinfo("پاک‌سازی", "فایل انتخاب‌شده حذف شد.")

    def preview_data(self):
        if self.loaded_df is None or self.loaded_df.empty:
            messagebox.showwarning("⚠️ هشدار", "ابتدا یک فایل اکسل معتبر انتخاب کنید.")
            return

        preview_win = tk.Toplevel(self.root)
        preview_win.title("پیش‌نمایش داده (10 ردیف اول)")
        preview_win.geometry("800x300")
        preview_win.minsize(600, 200)
        preview_win.configure(bg=self.colors["card_bg"])

        text_area = tk.Text(
            preview_win,
            font=("Consolas", 10),
            wrap="none",
            bg="#f8fafc",
            fg=self.colors["text"],
            relief="flat",
            padx=10,
            pady=10
        )
        text_area.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar_y = tk.Scrollbar(preview_win, orient="vertical", command=text_area.yview)
        scrollbar_y.pack(side="right", fill="y")
        text_area.configure(yscrollcommand=scrollbar_y.set)

        df_preview = self.loaded_df.head(10).fillna("[خالی]")
        cols = list(df_preview.columns)

        col_widths = []
        for col in cols:
            header_len = len(str(col))
            data_max = df_preview[col].astype(str).str.len().max() if not df_preview[col].empty else 0
            width = max(header_len, data_max, 5)
            col_widths.append(int(width))

        header = "ردیف".ljust(6) + " ".join(str(col).ljust(w) for col, w in zip(cols, col_widths))
        lines = [header, "-" * len(header)]

        for idx, (_, row) in enumerate(df_preview.iterrows()):
            row_str = f"{idx + 1:<5} " + " ".join(str(row[col]).ljust(w) for col, w in zip(cols, col_widths))
            lines.append(row_str)

        text_area.insert("1.0", "\n".join(lines))
        text_area.config(state="disabled")

    def sanitize_xml_tag(self, name):
        if name is None or (isinstance(name, float) and pd.isna(name)):
            return "Column"
        name = str(name).strip()
        if not name:
            return "Column"
        sanitized = re.sub(r'[^a-zA-Z\u0600-\u06FF0-9_.\-]', '_', name)
        sanitized = re.sub(r'_+', '_', sanitized).strip('_')
        if not sanitized:
            return "Column"
        if sanitized[0].isdigit() or sanitized[0] in "-.":
            sanitized = "col_" + sanitized
        return sanitized

    def convert_to_xml(self):
        path = self.excel_path.get()
        if not path or not os.path.isfile(path):
            messagebox.showerror("❌ خطا", "لطفاً یک فایل اکسل معتبر انتخاب کنید.")
            return

        try:
            df = pd.read_excel(path, dtype=str, keep_default_na=False, na_values=[""])
            if df.empty:
                messagebox.showwarning("⚠️ هشدار", "فایل اکسل خالی است!")
                return

            df.columns = [self.sanitize_xml_tag(col) for col in df.columns]
            root_elem = ET.Element("Worksheet")

            for idx, row in df.iterrows():
                row_elem = ET.SubElement(root_elem, "Row", id=str(idx + 1))
                for col in df.columns:
                    cell_elem = ET.SubElement(row_elem, col)
                    value = row[col]
                    cell_elem.text = str(value) if pd.notna(value) else ""

            rough = ET.tostring(root_elem, encoding='unicode')
            reparsed = minidom.parseString(rough)
            pretty_xml = "\n".join(line for line in reparsed.toprettyxml(indent="  ").splitlines() if line.strip())

            xml_path = os.path.splitext(path)[0] + ".xml"
            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(pretty_xml)

            # ✅ نمایش انیمیشن 3 ثانیه‌ای
            self.show_success_animation(xml_path)

        except Exception as e:
            messagebox.showerror("❌ خطا", f"خطا در تبدیل:\n{str(e)}")

    def show_success_animation(self, xml_path):
        success_win = tk.Toplevel(self.root)
        success_win.geometry("400x200")
        success_win.resizable(False, False)
        success_win.configure(bg=self.colors["card_bg"])
        success_win.overrideredirect(True)

        # مرکز‌گیری
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 100
        success_win.geometry(f"400x200+{x}+{y}")

        label = tk.Label(
            success_win,
            text="✅ تبدیل با موفقیت انجام شد!\n\n|",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors["card_bg"],
            fg=self.colors["success"]  # رنگ اولیه
        )
        label.pack(expand=True)

        # ✅ پالت رنگ زیبا برای چرخ‌دنده
        spinner_colors = [
            "#4361ee",  # آبی
            "#3a0ca3",  # بنفش تیره
            "#4cc9f0",  # آبی روشن
            "#4361ee",  # آبی
            "#f72585",  # صورتی انرژی
            "#ff9e00",
            "#ff5400",
            "#ff0058",
            "#c700ff",
            "#7d00ff",
            "#0077ff"
        ]
        spinner_frames = ["|", "/", "-", "\\", "|"]
        frame_index = 0

        def animate():
            nonlocal frame_index
            if success_win.winfo_exists():
                frame = spinner_frames[frame_index % len(spinner_frames)]
                color = spinner_colors[frame_index % len(spinner_colors)]
                label.config(text=f"✅ تبدیل با موفقیت انجام شد!\n\n{frame}", fg=color)
                frame_index += 1
                success_win.after(120, animate)  # سرعت کمی افزایش یافت

        animate()
        success_win.after(3000, lambda: self.close_success_animation(success_win, xml_path))

    def close_success_animation(self, win, xml_path):
        if win.winfo_exists():
            win.destroy()
        messagebox.showinfo("✅ موفقیت", f"فایل XML با موفقیت ایجاد شد!\n📁 {xml_path}")

    def exit_app(self):
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    root.tk.call('tk', 'scaling', 1.2)
    app = ExcelToXMLConverter(root)
    root.mainloop()
