import hashlib
import json
import logging
import os
import re
import threading
import time
from datetime import datetime
from tkinter import filedialog, messagebox
from typing import Any, Dict, List, Optional

import customtkinter as ctk
import pandas as pd
from dotenv import load_dotenv, set_key
from excel_parser import ExcelParser

# DND_FILES may not be available in all tkinter installations
try:
    from tkinter import DND_FILES
except ImportError:
    DND_FILES = None
from groq import Groq

load_dotenv()
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Constants
CACHE_EXPIRY_SECONDS = 86400 * 7  # 7 days
API_REQUEST_TIMEOUT = 30  # seconds
API_MAX_RETRIES = 3
GROQ_KEY_PREFIX = "gsk_"

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
CACHE_DIR = os.path.join(DATA_DIR, "cache")
LOGS_DIR = os.path.join(DATA_DIR, "logs")

for d in [DATA_DIR, CACHE_DIR, LOGS_DIR]:
    os.makedirs(d, exist_ok=True)

LOG_FILE = os.path.join(LOGS_DIR, f"app_{datetime.now().strftime('%Y%m')}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


def validate_groq_api_key(api_key: str) -> bool:
    """Перевірити чи дійсний API ключ Groq.

    Groq API ключі починаються з префіксу 'gsk_'.

    Args:
        api_key: Ключ для перевірки

    Returns:
        True якщо ключ валідний, False інакше
    """
    if not api_key or not isinstance(api_key, str):
        return False
    api_key = api_key.strip()
    if len(api_key) < 20:
        return False
    return api_key.startswith(GROQ_KEY_PREFIX)


class AICache:
    """Cache for AI analysis results to avoid repeated API calls."""

    def __init__(self, cache_dir: str):
        self.cache_dir = cache_dir
        os.makedirs(cache_dir, exist_ok=True)

    def _get_cache_key(self, data: List[List], model: str) -> str:
        data_str = str(data[:20])
        return hashlib.md5((data_str + model).encode()).hexdigest()

    def get(self, data: List[List], model: str) -> Optional[Dict]:
        key = self._get_cache_key(data, model)
        cache_file = os.path.join(self.cache_dir, f"{key}.json")
        if os.path.exists(cache_file):
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cached = json.load(f)
                    if time.time() - cached.get("timestamp", 0) < CACHE_EXPIRY_SECONDS:
                        logger.info(f"Cache hit for key {key[:8]}...")
                        return cached.get("result")
            except Exception:
                pass
        return None

    def set(self, data: List[List], model: str, result: Dict):
        key = self._get_cache_key(data, model)
        cache_file = os.path.join(self.cache_dir, f"{key}.json")
        try:
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(
                    {"timestamp": time.time(), "result": result}, f, ensure_ascii=False
                )
            logger.info(f"Cached result for key {key[:8]}...")
        except Exception as e:
            logger.error(f"Failed to cache: {e}")


class DropFrame(ctk.CTkFrame):
    """Frame that supports drag & drop files."""

    def __init__(self, master, callback, **kwargs):
        super().__init__(master, **kwargs)
        self.callback = callback
        self.configure(border_width=2, border_color="#3B8ED0")

        self.drop_label = ctk.CTkLabel(
            self,
            text="📂 Перетягніть файли сюди\nабо натисніть кнопку нижче",
            font=("Arial", 14),
            text_color="gray",
        )
        self.drop_label.pack(pady=30, padx=20)

        self.bind("<Button-1>", lambda e: self._on_click())
        self.drop_label.bind("<Button-1>", lambda e: self._on_click())

        try:
            self.drop_target_register = lambda: None
        except:
            pass

    def _on_click(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if files:
            self.callback(list(files))


class AIPassboltApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI Passbolt Migration Tool")
        self.geometry("900x650")

        self.groq_key = os.getenv("GROQ_API_KEY", "")
        self.source_file = ""
        self.parsed_data = []
        self.preview_data = None
        self.use_ai_detection = True  # Enable AI by default

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self, width=870)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tab_converter = self.tabview.add("З файлу")
        self.tab_clipboard = self.tabview.add("З буфера")
        self.tab_settings = self.tabview.add("Налаштування")
        self.tab_preview = self.tabview.add("Попередній перегляд")

        self.setup_converter_tab()
        self.setup_clipboard_tab()
        self.setup_settings_tab()
        self.setup_preview_tab()

    def setup_converter_tab(self):
        self.tab_converter.grid_columnconfigure(0, weight=1)

        # File selection frame
        self.file_frame = ctk.CTkFrame(self.tab_converter)
        self.file_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        self.lbl_file = ctk.CTkLabel(
            self.file_frame, text="Виберіть Excel файл:", font=("Arial", 12, "bold")
        )
        self.lbl_file.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.btn_browse = ctk.CTkButton(
            self.file_frame, text="Оглянути...", command=self.browse_file, width=150
        )
        self.btn_browse.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.lbl_selected_file = ctk.CTkLabel(
            self.file_frame, text="Файл не вибрано", text_color="gray"
        )
        self.lbl_selected_file.grid(row=1, column=2, padx=20, pady=10, sticky="w")

        # Action Buttons
        self.btn_analyze = ctk.CTkButton(
            self.tab_converter,
            text="Обробити",
            command=self.smart_parse,
            state="disabled",
            height=45,
            font=("Arial", 14, "bold"),
            fg_color="#2D8C3C",
            hover_color="#1E5E28",
        )
        self.btn_analyze.grid(row=1, column=0, padx=20, pady=15, sticky="ew")

        # Progress bar and label
        self.progress_frame = ctk.CTkFrame(self.tab_converter, fg_color="transparent")
        self.progress_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        self.lbl_progress = ctk.CTkLabel(
            self.progress_frame, text="", text_color="gray", width=150
        )
        self.lbl_progress.pack(side="left", fill="x", expand=True)

        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame, mode="indeterminate", width=200
        )
        self.progress_bar.pack(side="right", padx=(10, 0))

        # Result frame
        self.result_frame = ctk.CTkScrollableFrame(
            self.tab_converter, label_text="Результат конвертації", height=250
        )
        self.result_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.tab_converter.grid_rowconfigure(3, weight=1)

        # Export button
        self.btn_export = ctk.CTkButton(
            self.tab_converter,
            text="Зберегти CSV для Passbolt",
            command=self.export_csv,
            state="disabled",
            fg_color="#2D8C3C",
            hover_color="#1E5E28",
            height=40,
            font=("Arial", 13, "bold"),
        )
        self.btn_export.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

    def setup_clipboard_tab(self):
        """Setup tab for pasting data from clipboard."""
        self.tab_clipboard.grid_columnconfigure(0, weight=1)
        self.tab_clipboard.grid_rowconfigure(1, weight=1)

        # Instructions
        self.lbl_clipboard_instr = ctk.CTkLabel(
            self.tab_clipboard,
            text="1. Скопіюйте таблицю з Excel (Ctrl+C)\n2. Вставте в поле нижче (Ctrl+V)\n3. Натисніть 'Обробити'\n4. Експортуйте CSV для Passbolt",
            font=("Arial", 12),
            justify="left",
        )
        self.lbl_clipboard_instr.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        # Text area for pasted data
        self.clipboard_text = ctk.CTkTextbox(self.tab_clipboard, font=("Consolas", 11))
        self.clipboard_text.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # Buttons frame
        self.btn_frame_clipboard = ctk.CTkFrame(self.tab_clipboard)
        self.btn_frame_clipboard.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        self.btn_process_clipboard = ctk.CTkButton(
            self.btn_frame_clipboard,
            text="Обробити",
            command=self.process_clipboard_data,
            height=40,
            font=("Arial", 14, "bold"),
            fg_color="#2D8C3C",
            hover_color="#1E5E28",
        )
        self.btn_process_clipboard.grid(row=0, column=0, padx=10, pady=5)

        self.btn_clear_clipboard = ctk.CTkButton(
            self.btn_frame_clipboard,
            text="Очистити",
            command=self.clear_clipboard,
            height=40,
            fg_color="gray",
            hover_color="darkgray",
        )
        self.btn_clear_clipboard.grid(row=0, column=1, padx=10, pady=5)

        # Result label
        self.lbl_clipboard_result = ctk.CTkLabel(
            self.tab_clipboard, text="", text_color="green"
        )
        self.lbl_clipboard_result.grid(row=3, column=0, padx=20, pady=5)

        # Export button
        self.btn_export_clipboard = ctk.CTkButton(
            self.tab_clipboard,
            text="Експортувати CSV для Passbolt",
            command=self.export_clipboard_csv,
            state="disabled",
            height=40,
            font=("Arial", 12, "bold"),
        )
        self.btn_export_clipboard.grid(row=4, column=0, padx=20, pady=10)

        # Store parsed data
        self.clipboard_parsed_data = []

    def process_clipboard_data(self):
        """Process pasted table data."""
        text = self.clipboard_text.get("1.0", "end-1c")

        if not text.strip():
            messagebox.showwarning("Помилка", "Вставте таблицю в поле!")
            return

        try:
            # Parse tab-separated or comma-separated data
            lines = text.strip().split("\n")
            self.clipboard_parsed_data = []

            # Detect delimiter (tab or comma)
            delimiter = "\t" if "\t" in lines[0] else ","

            # Get headers from first line
            headers = [h.strip().strip('"') for h in lines[0].split(delimiter)]

            # Find column indices
            col_map = {}
            for i, h in enumerate(headers):
                h_lower = h.lower()
                if "title" in h_lower or "name" in h_lower or "назва" in h_lower:
                    col_map["Title"] = i
                elif "user" in h_lower or "login" in h_lower or "ім" in h_lower:
                    col_map["Username"] = i
                elif "pass" in h_lower or "пароль" in h_lower:
                    col_map["Password"] = i
                elif "url" in h_lower or "ip" in h_lower or "uri" in h_lower:
                    col_map["URL"] = i
                elif "group" in h_lower or "група" in h_lower:
                    col_map["Group"] = i
                elif "note" in h_lower or "приміт" in h_lower:
                    col_map["Notes"] = i

            # Parse data rows
            for line in lines[1:]:
                if not line.strip():
                    continue

                values = [v.strip().strip('"') for v in line.split(delimiter)]

                record = {
                    "Group": values[col_map["Group"]]
                    if "Group" in col_map and col_map["Group"] < len(values)
                    else "Imported",
                    "Title": values[col_map["Title"]]
                    if "Title" in col_map and col_map["Title"] < len(values)
                    else "",
                    "Username": values[col_map["Username"]]
                    if "Username" in col_map and col_map["Username"] < len(values)
                    else "",
                    "Password": values[col_map["Password"]]
                    if "Password" in col_map and col_map["Password"] < len(values)
                    else "",
                    "URL": values[col_map["URL"]]
                    if "URL" in col_map and col_map["URL"] < len(values)
                    else "",
                    "Notes": values[col_map["Notes"]]
                    if "Notes" in col_map and col_map["Notes"] < len(values)
                    else "",
                }

                # Skip if no title
                if not record["Title"]:
                    continue

                # Normalize URL
                if record["URL"] and not record["URL"].startswith("http"):
                    if re.match(r"^\d{1,3}\.", record["URL"]):
                        record["URL"] = f"http://{record['URL']}"

                self.clipboard_parsed_data.append(record)

            if self.clipboard_parsed_data:
                self.lbl_clipboard_result.configure(
                    text=f"✅ Знайдено {len(self.clipboard_parsed_data)} записів!",
                    text_color="green",
                )
                self.btn_export_clipboard.configure(state="normal")
            else:
                messagebox.showwarning("Помилка", "Не знайдено даних для обробки!")

        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося обробити дані:\n{e}")

    def clear_clipboard(self):
        """Clear clipboard text area."""
        self.clipboard_text.delete("1.0", "end")
        self.lbl_clipboard_result.configure(text="")
        self.btn_export_clipboard.configure(state="disabled")
        self.clipboard_parsed_data = []

    def export_clipboard_csv(self):
        """Export clipboard data to CSV."""
        if not self.clipboard_parsed_data:
            messagebox.showwarning("Помилка", "Немає даних для експорту!")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="passbolt_import.csv",
        )

        if file_path:
            try:
                from excel_parser import ExcelParser

                parser = ExcelParser.__new__(ExcelParser)
                parser.PASSBOLT_COLUMNS = [
                    "Group",
                    "Title",
                    "Username",
                    "Password",
                    "URL",
                    "Notes",
                ]
                parser.export_to_csv(self.clipboard_parsed_data, file_path)

                messagebox.showinfo(
                    "Успіх",
                    f"Експортовано {len(self.clipboard_parsed_data)} записів!\n\nТепер імпортуйте цей файл в Passbolt.",
                )
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося експортувати:\n{e}")

    def setup_settings_tab(self):
        self.lbl_groq = ctk.CTkLabel(
            self.tab_settings, text="Groq API Key:", font=("Arial", 14, "bold")
        )
        self.lbl_groq.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")

        self.entry_groq = ctk.CTkEntry(self.tab_settings, width=600, show="*")
        self.entry_groq.insert(0, self.groq_key)
        self.entry_groq.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.btn_save_keys = ctk.CTkButton(
            self.tab_settings, text="Зберегти", command=self.save_keys, width=150
        )
        self.btn_save_keys.grid(row=2, column=0, padx=20, pady=20, sticky="w")

        # AI Detection toggle
        self.ai_detection_var = ctk.BooleanVar(value=True)

        self.chk_ai_detection = ctk.CTkCheckBox(
            self.tab_settings,
            text="Використовувати AI для визначення структури таблиці",
            variable=self.ai_detection_var,
            command=self.toggle_ai_detection,
            font=("Arial", 12),
        )
        self.chk_ai_detection.grid(row=3, column=0, padx=20, pady=10, sticky="w")

        # Model selection
        self.model_var = ctk.StringVar(value="llama-3.3-70b-versatile")

        self.lbl_model = ctk.CTkLabel(
            self.tab_settings, text="Модель AI:", font=("Arial", 14, "bold")
        )
        self.lbl_model.grid(row=4, column=0, padx=20, pady=(20, 5), sticky="w")

        self.model_menu = ctk.CTkOptionMenu(
            self.tab_settings,
            values=[
                "llama-3.3-70b-versatile",
                "qwen-2.5-coder-32b",
            ],
            variable=self.model_var,
            width=300,
        )
        self.model_menu.grid(row=5, column=0, padx=20, pady=5, sticky="w")

        # Info label
        info_text = """
Інструкція:
1. Отримайте API ключ на https://console.groq.com
2. Введіть ключ у поле вище
3. Натисніть "Зберегти"
4. Увімкніть "Використовувати AI" для кращого розпізнавання

Рекомендована модель (максимальна якість): llama-3.3-70b-versatile
"""
        self.lbl_info = ctk.CTkLabel(
            self.tab_settings, text=info_text, justify="left", font=("Arial", 11)
        )
        self.lbl_info.grid(row=6, column=0, padx=20, pady=20, sticky="w")

    def setup_preview_tab(self):
        self.tab_preview.grid_columnconfigure(0, weight=1)
        self.tab_preview.grid_rowconfigure(0, weight=1)

        self.preview_text = ctk.CTkTextbox(self.tab_preview, font=("Consolas", 11))
        self.preview_text.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.btn_refresh_preview = ctk.CTkButton(
            self.tab_preview, text="Оновити", command=self.refresh_preview, width=150
        )
        self.btn_refresh_preview.grid(row=1, column=0, padx=20, pady=10)

    def save_keys(self):
        self.groq_key = self.entry_groq.get().strip()
        self.use_ai_detection = self.ai_detection_var.get()

        # Validate API key format
        if self.groq_key and not validate_groq_api_key(self.groq_key):
            messagebox.showwarning(
                "Попередження",
                "API ключ має сумнівний формат.\n"
                "Ключі Groq починаються з 'gsk_'.\n\n"
                "Перевірте ключ на https://console.groq.com",
            )

        # Create .env file if it doesn't exist
        if not os.path.exists(".env"):
            with open(".env", "w") as f:
                pass

        set_key(".env", "GROQ_API_KEY", self.groq_key)
        messagebox.showinfo("Успіх", "API ключ збережено!")

    def toggle_ai_detection(self):
        """Toggle AI detection on/off."""
        self.use_ai_detection = self.ai_detection_var.get()

    def browse_file(self):
        file = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ]
        )
        if file:
            self.source_file = file
            try:
                # Check if file exists
                if not os.path.exists(file):
                    messagebox.showerror("Помилка", f"Файл не знайдено:\n{file}")
                    return

                parser = ExcelParser(
                    file,
                    groq_api_key=self.groq_key if self.use_ai_detection else None,
                    model=self.model_var.get(),
                )
                sheet_names = parser.get_sheet_names()

                # Load data for preview
                parser.load_excel()

                self.lbl_selected_file.configure(
                    text=f"{os.path.basename(file)}\nАркушів: {len(sheet_names)} | Рядків: {len(parser.raw_data)}",
                    text_color="#3B8ED0",
                )
                self.btn_analyze.configure(state="normal")

                # Store parser for later use
                self.preview_data = parser

                # Show preview
                self.show_preview(parser)

            except PermissionError:
                messagebox.showerror(
                    "Помилка",
                    f"Файл заблоковано!\n\nЗакрийте файл в Excel і спробуйте знову.",
                )
            except Exception as e:
                import traceback

                error_details = traceback.format_exc()
                print(f"Error loading file: {error_details}")
                messagebox.showerror(
                    "Помилка",
                    f"Не вдалося прочитати файл:\n{e}\n\n"
                    f"Переконайтесь що файл НЕ відкритий в Excel.",
                )

    def show_preview(self, parser: ExcelParser):
        """Show raw data preview."""
        self.preview_text.delete("1.0", "end")

        preview_lines = []
        preview_lines.append(f"Файл: {os.path.basename(parser.file_path)}\n")
        preview_lines.append(f"Аркуші: {', '.join(parser.get_sheet_names())}\n")
        preview_lines.append("=" * 80 + "\n\n")

        # Show first 10 rows of raw data
        preview_lines.append("Перші рядки даних:\n")
        for i, row in enumerate(parser.raw_data[:10]):
            preview_lines.append(f"Рядок {i}: {row}\n")

        self.preview_text.insert("1.0", "".join(preview_lines))

    def refresh_preview(self):
        if self.preview_data:
            self.show_preview(self.preview_data)

    def smart_parse(self):
        """Smart parsing - automatically chooses best method."""
        if not self.source_file:
            return

        self.btn_analyze.configure(state="disabled", text="Обробка...")
        self.lbl_progress.configure(text="Завантаження файлу...")
        self.progress_bar.start()
        threading.Thread(target=self._run_smart_parse, daemon=True).start()

    def _run_smart_parse(self):
        """Run smart parsing in background thread."""
        try:
            parser = ExcelParser(
                self.source_file,
                groq_api_key=self.groq_key if self.use_ai_detection else None,
                model=self.model_var.get(),
            )
            parser.load_excel()

            # Try all fast preprocessing formats first
            processed_data = None
            method_used = ""

            # Try Format 1: Server list with Login/Password column
            processed_data = self._preprocess_server_list_format(parser.raw_data)
            if processed_data:
                method_used = "Server List"

            # Try Format 2: Kubernetes cluster format
            if not processed_data:
                processed_data = self._preprocess_k8s_format(parser.raw_data)
                if processed_data:
                    method_used = "K8s Cluster"

            # Try Format 3: Complex table with separate credentials
            if not processed_data:
                processed_data = self._preprocess_complex_table_format(parser.raw_data)
                if processed_data:
                    method_used = "Complex Table"

            # Try Format 4: Column-based
            if not processed_data:
                processed_data = self._preprocess_vertical_format(parser.raw_data)
                if processed_data:
                    method_used = "Vertical"

            # Try Format 5: Scattered
            if not processed_data:
                processed_data = self._preprocess_scattered_format(parser.raw_data)
                if processed_data:
                    method_used = "Scattered"

            # Try Format 6: Table format
            if not processed_data:
                processed_data = self._preprocess_table_format(parser.raw_data)
                if processed_data:
                    method_used = "Table"

            if processed_data:
                # Fast method worked - use result
                self.parsed_data = []
                for row in processed_data[1:]:
                    record = {
                        "Group": "Imported",
                        "Title": row[0] if len(row) > 0 else "",
                        "Username": row[1] if len(row) > 1 else "",
                        "Password": row[2] if len(row) > 2 else "",
                        "URL": row[3] if len(row) > 3 else "",
                        "Notes": row[4] if len(row) > 4 else "",
                    }
                    if record["Title"]:
                        self.parsed_data.append(record)

                self.after(0, self.update_result_ui)
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Готово",
                        f"Оброблено {len(self.parsed_data)} записів\n"
                        f"Метод: {method_used} (швидкий парсинг)",
                    ),
                )
            else:
                if not self.use_ai_detection or not self.groq_key:
                    # Fast methods failed - fall back to rule-based parsing
                    self.after(
                        0, lambda: self.lbl_progress.configure(text="Парсинг без AI...")
                    )
                    try:
                        self.parsed_data = parser.parse(use_ai=False)
                        self.after(0, self.update_result_ui)
                        self.after(
                            0,
                            lambda: messagebox.showinfo(
                                "Готово",
                                f"Оброблено {len(self.parsed_data)} записів\n"
                                f"Метод: Rule-based (без AI)",
                            ),
                        )
                    except Exception as e:
                        self.after(0, self.stop_progress)
                        self.after(
                            0,
                            lambda: messagebox.showerror(
                                "Помилка", f"Помилка парсингу: {e}"
                            ),
                        )
                else:
                    # Fast methods failed - use AI automatically
                    self.after(
                        0, lambda: self.lbl_progress.configure(text="AI аналізує...")
                    )

                    # Run AI analysis
                    raw_data_str = ""
                    for i, row in enumerate(parser.raw_data[:50]):
                        raw_data_str += f"Row {i}: {row}\n"

                    prompt = f"""
You are an expert at parsing Excel tables for Passbolt password manager import.

Passbolt CSV format has these fields:
- Group: group/category/environment
- Title: resource name (REQUIRED)
- Username: login/username
- Password: password/secret (can be empty)
- URL: resource address (IP, domain, with http:// or https://)
- Notes: additional info

DATA ({len(parser.raw_data)} rows):
{raw_data_str}

TASK:
1. Identify server/resource rows
2. Extract Title, Username, Password, URL for EACH resource
3. Keep ALL records - even if Password is empty
4. Return JSON object with a records array

Return ONLY JSON:
{{"records": [
  {{"Group": "Imported", "Title": "Server", "Username": "admin", "Password": "pass123", "URL": "http://192.168.1.1", "Notes": ""}}
]}}

IMPORTANT: Create ONE record for EACH server/resource row. Do NOT skip any!
"""

                    try:
                        client = Groq(api_key=self.groq_key)
                        chat = client.chat.completions.create(
                            messages=[{"role": "user", "content": prompt}],
                            model=self.model_var.get(),
                            response_format={"type": "json_object"},
                            temperature=0.1,
                        )

                        result = json.loads(chat.choices[0].message.content)

                        # Handle different response formats
                        if isinstance(result, list):
                            self.parsed_data = result
                        elif isinstance(result, dict):
                            if "resources" in result:
                                self.parsed_data = result["resources"]
                            elif "data" in result:
                                self.parsed_data = result["data"]
                            elif "records" in result:
                                self.parsed_data = result["records"]
                            else:
                                self.parsed_data = [result]

                        self.after(0, self.update_result_ui)
                        self.after(
                            0,
                            lambda: messagebox.showinfo(
                                "Готово",
                                f"Оброблено {len(self.parsed_data)} записів\n"
                                f"Метод: AI Analysis (розумний парсинг)",
                            ),
                        )
                    except Exception as e:
                        self.after(0, self.stop_progress)
                        self.after(
                            0,
                            lambda: messagebox.showerror(
                                "Помилка", f"Помилка AI парсингу: {e}"
                            ),
                        )

        except Exception as e:
            import traceback

            error_details = traceback.format_exc()
            self.after(0, self.stop_progress)
            self.after(
                0,
                lambda: messagebox.showerror(
                    "Помилка", f"Помилка: {e}\n\nДеталі: {error_details}"
                ),
            )

        self.after(0, self.stop_progress)
        self.after(
            0, lambda: self.btn_analyze.configure(state="normal", text="Обробити")
        )

    def _preprocess_vertical_format(self, raw_data):
        """
        Detect and transform vertical/block format to standard table.

        Example input:
        Row 0: ['Server-01', 'Server-02']
        Row 1: ['login', 'pu']
        Row 2: ['pas', '@912Lkmo']

        Output:
        Row 0: ['Title', 'Username', 'Password', 'URL', 'Notes']
        Row 1: ['Server-01', 'pu', '@912Lkmo', '', '']
        Row 2: ['Server-02', 'pu', '@912Lkmo', '', '']
        """
        if len(raw_data) < 3:
            return None

        # Detect credential rows (rows containing login/pas/user/pass keywords)
        cred_keywords = [
            "login",
            "user",
            "username",
            "pass",
            "pas",
            "pwd",
            "password",
            "url",
            "ip",
            "host",
        ]

        cred_type_map = {
            "login": "Username",
            "user": "Username",
            "username": "Username",
            "pass": "Password",
            "pas": "Password",
            "pwd": "Password",
            "password": "Password",
            "url": "URL",
            "ip": "URL",
            "host": "URL",
        }

        cred_rows = {}  # Map Passbolt field to row values
        title_rows = []  # Rows that look like server/resource names

        for i, row in enumerate(raw_data):
            if not row or not any(cell for cell in row):
                continue

            row_str = " ".join(str(cell).lower() for cell in row if cell)

            # Check if this row contains credential labels
            is_cred_row = any(kw in row_str for kw in cred_keywords)

            if is_cred_row:
                detected_fields = set()
                for cell in row:
                    if not cell:
                        continue
                    cell_lower = str(cell).lower().strip()
                    for kw in cred_keywords:
                        if kw in cell_lower:
                            field = cred_type_map.get(kw)
                            if field:
                                detected_fields.add(field)
                for field in detected_fields:
                    if field not in cred_rows:
                        cred_rows[field] = {"row": i, "values": row[:]}
            else:
                # This might be a title row (server names)
                # Check if all cells look like names (not credentials)
                looks_like_titles = all(
                    cell and not str(cell).lower().strip() in cred_keywords
                    for cell in row
                    if cell
                )
                if looks_like_titles and len([c for c in row if c]) >= 1:
                    title_rows.append((i, row))

        # If we found both titles and credentials, transform
        if title_rows and len(cred_rows) >= 1:
            # Build standard table with Passbolt columns
            result = [["Title", "Username", "Password", "URL", "Notes"]]

            def _pick_value(values, col_idx):
                # Try same column first
                if col_idx < len(values):
                    val = values[col_idx]
                    if val and not any(kw in str(val).lower() for kw in cred_keywords):
                        return str(val).strip()
                # Fallback: first non-label value
                for v in values:
                    if v and not any(kw in str(v).lower() for kw in cred_keywords):
                        return str(v).strip()
                return ""

            # Data rows - one per title
            for title_idx, title_row in title_rows:
                for title_col, title_value in enumerate(title_row):
                    if title_value and str(title_value).strip():
                        title = str(title_value).strip()
                        username = (
                            _pick_value(cred_rows["Username"]["values"], title_col)
                            if "Username" in cred_rows
                            else ""
                        )
                        password = (
                            _pick_value(cred_rows["Password"]["values"], title_col)
                            if "Password" in cred_rows
                            else ""
                        )
                        url = (
                            _pick_value(cred_rows["URL"]["values"], title_col)
                            if "URL" in cred_rows
                            else ""
                        )

                        # Normalize URL/IP if possible
                        if url and not url.startswith(
                            ("http://", "https://", "ftp://")
                        ):
                            if re.match(r"^\d{1,3}\.", url):
                                url = f"http://{url}"

                        result.append([title, username, password, url, ""])

            if len(result) > 1:
                return result

        return None

    def _preprocess_scattered_format(self, raw_data):
        """
        Handle scattered row-based format where data is spread across many rows.

        Example (FREGAT_DEV.xlsx style):
        Row 0: [URL, None, None, None, Description]
        Row 1: [None, None, None, 'Server:rdp: 172.16.33.36']
        Row 2: ['k8s-fregate', '172.16.33.27', None, None, 'Логін: ...']
        Row 3: [None, None, None, None, 'Пароль: Qq15guteE2@']

        Returns: Consolidated record with all data extracted
        """
        import re

        if len(raw_data) < 2:
            return None

        # Collect all non-None values with positions
        all_values = []
        for i, row in enumerate(raw_data):
            if not row:
                continue
            for j, cell in enumerate(row):
                if cell is not None and str(cell).strip():
                    all_values.append((i, j, str(cell).strip()))

        if not all_values:
            return None

        # Keywords for detection
        cred_keywords = [
            "login",
            "user",
            "pass",
            "pas",
            "pwd",
            "password",
            "логін",
            "пароль",
            "користувач",
            "сервер",
            "база",
            "database",
            "db",
            "url",
            "ip",
            "host",
        ]

        # Extract fields
        title = ""
        username = ""
        password = ""
        url = ""
        notes_parts = []

        # First pass: find title (resource name)
        for i, j, val in all_values:
            val_lower = val.lower()
            # Skip URLs, credentials, metadata
            is_url = val.startswith("http") or "://" in val
            is_cred_label = any(kw in val_lower for kw in cred_keywords)
            is_metadata = any(
                kw in val_lower
                for kw in [
                    "комплект",
                    "находится",
                    "dmz",
                    "доступ",
                    "база",
                    "описание",
                    "description",
                    "сервер:",
                    "логін:",
                    "пароль:",
                ]
            )

            if not is_url and not is_cred_label and not is_metadata:
                # Check if it looks like a name (not IP, not too short)
                if not re.match(r"^\d{1,3}\.", val) and len(val) > 2:
                    title = val
                    break

        # If no title found, use first non-URL value
        if not title:
            for i, j, val in all_values:
                if not val.startswith("http") and ":" not in val:
                    title = val
                    break

        # Second pass: extract all credentials and data
        for i, j, val in all_values:
            val_lower = val.lower()

            # URL
            if val.startswith("http") or "://" in val:
                if not url:
                    url = val

            # Username patterns
            elif any(
                kw in val_lower for kw in ["логін", "login:", "user:", "користувач:"]
            ):
                if ":" in val:
                    extracted = val.split(":", 1)[1].strip()
                    if extracted and not username:
                        username = extracted

            # Password patterns
            elif any(
                kw in val_lower for kw in ["пароль", "pass:", "password:", "pwd:"]
            ):
                if ":" in val:
                    extracted = val.split(":", 1)[1].strip()
                    if extracted and not password:
                        password = extracted

            # IP addresses
            elif re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", val):
                if not url:
                    url = f"http://{val}"
                else:
                    notes_parts.append(f"IP: {val}")

            # Server info
            elif "сервер" in val_lower or "server" in val_lower:
                if ":" in val:
                    parts = val.split(":", 1)
                    if len(parts) > 1:
                        server_info = parts[1].strip()
                        # Check if it contains IP
                        ip_match = re.search(
                            r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", server_info
                        )
                        if ip_match and not url:
                            url = f"http://{ip_match.group()}"
                        else:
                            notes_parts.append(f"Server: {server_info}")

            # Database info
            elif "база даних" in val_lower or "database" in val_lower:
                notes_parts.append(val)

            # Description/metadata
            elif (
                "комплект" in val_lower
                or "находится" in val_lower
                or "dmz" in val_lower
            ):
                notes_parts.append(val)

            # Access info
            elif "доступ до" in val_lower:
                notes_parts.append(val)

        # Build result
        if title or username or password or url:
            result = [["Title", "Username", "Password", "URL", "Notes"]]
            notes_str = " | ".join(notes_parts) if notes_parts else ""
            result.append([title, username, password, url, notes_str])
            return result

        return None

    def _preprocess_table_format(self, raw_data):
        """
        Handle standard table format with multiple rows of server data.

        Example (Master-MedStar.xlsx style):
        Row 0: ['Ingress', '192.168.0.100', '94.131.242.33', 'root', None, None, 'URL']
        Row 1: ['k8s-master', '192.168.0.101', '94.131.242.33:23', 'AV7K•kW55Y', None, None, None]
        Row 2: ['k8s-worker01', '192.168.0.102', '94.131.242.33:24', None, None, None, 'admin']
        Row 3: ['k8s-worker02', '192.168.0.104', '94.131.242.33:25', None, None, None, None]

        Returns: Table with Title, Username, Password, URL, Notes for each server
        """
        import re

        if len(raw_data) < 2:
            return None

        result = [["Title", "Username", "Password", "URL", "Notes"]]

        # Skip header row if first row looks like column names (A, B, C or Col_0, etc.)
        start_row = 0
        if raw_data and len(raw_data) > 0:
            first_row = raw_data[0]
            if first_row and len(first_row) > 0:
                first_val = str(first_row[0]).strip() if first_row[0] else ""
                # Skip if it's a column header like 'A', 'B', 'C', 'Col_0', etc.
                if first_val in [
                    "A",
                    "B",
                    "C",
                    "D",
                    "E",
                    "F",
                    "G",
                ] or first_val.startswith("Col_"):
                    start_row = 1

        # Process all data rows
        for row_idx in range(start_row, len(raw_data)):
            row = raw_data[row_idx]
            if not row or not any(cell for cell in row):
                continue

            # Extract values from ALL columns
            title = ""
            username = ""
            password = ""
            url = ""
            notes_parts = []
            all_ips = []

            for col_idx, cell in enumerate(row):
                if cell is None or not str(cell).strip():
                    continue

                val = str(cell).strip()
                val_lower = val.lower()

                # Column 0: Title/Name
                if col_idx == 0 and val:
                    # Skip if it looks like a header
                    if val in ["A", "B", "C", "D", "E", "F", "G"] or val.startswith(
                        "Col_"
                    ):
                        continue
                    title = val

                # IP addresses (any column)
                elif re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", val):
                    all_ips.append(val)

                # URLs
                elif val.startswith("http") or "://" in val:
                    if not url:
                        url = val
                    else:
                        notes_parts.append(f"URL: {val}")

                # Potential passwords or usernames in other columns (any length >= 2)
                elif len(val) >= 2:
                    # Check for password patterns
                    has_upper = any(c.isupper() for c in val)
                    has_lower = any(c.islower() for c in val)
                    has_digit = any(c.isdigit() for c in val)
                    has_special = any(c in "!@#$%^&*•·?/" for c in val)

                    is_likely_password = (
                        (has_upper and has_digit)
                        or (has_upper and has_lower and len(val) >= 6)
                        or has_special
                    )

                    # Check for common usernames
                    is_username = val_lower in [
                        "admin",
                        "root",
                        "administrator",
                        "user",
                        "test",
                        "sa",
                        "postgres",
                        "mysql",
                        "oracle",
                        "nginx",
                        "apache",
                    ]

                    if is_likely_password and not password:
                        password = val
                    elif is_username and not username:
                        username = val
                    elif not username and not password and len(val) >= 3:
                        # Could be either - check characteristics
                        if val_lower.isalpha() or val_lower in [
                            "admin",
                            "root",
                            "user",
                        ]:
                            username = val
                        else:
                            password = val

            # Build record if we have a title
            if title:
                # Set URL from first IP if no URL found
                if not url and all_ips:
                    first_ip = all_ips[0].split(":")[0]
                    first_port = all_ips[0].split(":")[1] if ":" in all_ips[0] else ""
                    url = f"http://{first_ip}"
                    if first_port:
                        url += f":{first_port}"

                # Add remaining IPs to notes
                for ip in all_ips[1:]:
                    notes_parts.append(f"IP: {ip}")

                notes_str = " | ".join(notes_parts) if notes_parts else ""
                result.append([title, username, password, url, notes_str])

        return result if len(result) > 1 else None

    def _preprocess_k8s_format(self, raw_data):
        """
        Handle Kubernetes cluster format with shared credentials.

        Example (k8s-emp-sp.xlsx style):
        Row 0: ['ver 1.34', 'Ubuntu 24.04']  ← cluster info
        Row 1: ['k8s-emp-sp', 'sa / 291263']  ← base name + shared credentials
        Row 2: ['k8s-emp-sp-master', '192.168.71.137']  ← server + IP
        Row 3: ['k8s-emp-sp-worker', '192.168.71.138']  ← server + IP
        Row 4: ['ingress', '192.168.71.143']  ← server + IP

        Returns: One record per server with shared credentials
        """
        import re

        if len(raw_data) < 3:
            return None

        # Find credentials row - look for " / " pattern in second column
        cred_row_idx = -1
        shared_username = ""
        shared_password = ""

        for i, row in enumerate(raw_data):
            if not row or len(row) < 2:
                continue

            # Check second column for credentials pattern "user / pass" or "user/pass"
            second_cell = str(row[1]).strip() if len(row) > 1 else ""

            # Look for pattern with spaces: "sa / 291263"
            if " / " in second_cell:
                cred_row_idx = i
                parts = second_cell.split(" / ")
                if len(parts) >= 2:
                    shared_username = parts[0].strip()
                    shared_password = parts[1].strip()
                break

            # Also check without spaces: "user/pass"
            elif "/" in second_cell and not second_cell.startswith("http"):
                cred_row_idx = i
                parts = second_cell.split("/")
                if len(parts) >= 2:
                    shared_username = parts[0].strip()
                    shared_password = parts[1].strip()
                break

        # If no credentials found, can't process this format
        if cred_row_idx == -1:
            return None

        # Find server rows (name + IP pattern)
        servers = []
        for i, row in enumerate(raw_data):
            if not row or len(row) < 1:
                continue

            name = str(row[0]).strip() if row[0] else ""
            ip_or_val = str(row[1]).strip() if len(row) > 1 and row[1] else ""

            # Skip version/info rows
            if name.lower().startswith("ver") or name.lower().startswith("ubuntu"):
                continue

            # Skip credential row itself
            if i == cred_row_idx:
                continue

            # Check if it's a server row (has IP or looks like server name)
            is_ip = re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", ip_or_val)

            if name and len(name) > 1:
                # Include all rows that aren't version info
                servers.append(
                    {
                        "title": name,
                        "ip": ip_or_val if is_ip else "",
                        "username": shared_username,
                        "password": shared_password,
                    }
                )

        if not servers:
            return None

        # Build result table
        result = [["Title", "Username", "Password", "URL", "Notes"]]

        for server in servers:
            url = f"http://{server['ip']}" if server["ip"] else ""
            result.append(
                [server["title"], server["username"], server["password"], url, ""]
            )

        return result if len(result) > 1 else None

    def _preprocess_complex_table_format(self, raw_data):
        """
        Handle complex table format with separate credential section.

        Example:
        Row 0: ['K8S-EMPX', 'ver 1.34', 'Ubuntu 24.04']  ← header
        Row 1: ['', '(old)', 'new ip(host-vm11)', 'linux login']  ← labels
        Row 2: ['Empx-SQL', '192.168.50.200', '172.16.36.41', 'sa', '291263']  ← servers
        Row 3: ['Empx-K8S', '192.168.50.201', '172.16.36.40']  ← servers
        Row 4: ['ingress', '192.168.50.202', '172.16.36.40']  ← servers
        Row 5: ['sql', 'sa', '1Qazxcvb']  ← separate credentials

        Returns: Records for each server with credentials
        """
        import re

        if len(raw_data) < 3:
            return None

        # Collect all server rows (have IP pattern)
        servers = []
        credentials = {}  # Map service names to credentials

        for i, row in enumerate(raw_data):
            if not row:
                continue

            # Check if this row has credentials section (short row with name + user + pass)
            if len(row) >= 3:
                first = str(row[0]).strip() if row[0] else ""
                second = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                third = str(row[2]).strip() if len(row) > 2 and row[2] else ""

                # Credential row: short name + username + password (no IPs)
                is_ip_1 = (
                    re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", second)
                    if second
                    else False
                )
                is_ip_2 = (
                    re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", third)
                    if third
                    else False
                )

                if first and second and third and not is_ip_1 and not is_ip_2:
                    # This looks like a credential row
                    if first.lower() not in ["ver", "ubuntu", "k8s", "old", "new"]:
                        credentials[first.lower()] = {
                            "username": second,
                            "password": third,
                        }
                        continue

            # Server row: name + IP(s)
            if len(row) >= 2:
                name = str(row[0]).strip() if row[0] else ""
                ip1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                ip2 = str(row[2]).strip() if len(row) > 2 and row[2] else ""

                # Check if it's a server row (has IP)
                is_ip = re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", ip1)

                if name and is_ip and len(name) > 2:
                    # Skip header rows
                    if name.lower() in ["ver", "ubuntu", "k8s", "old", "new"]:
                        continue

                    # Find matching credentials
                    username = ""
                    password = ""
                    name_lower = name.lower()

                    # Try to match with credentials
                    for cred_name, cred_data in credentials.items():
                        if cred_name in name_lower or name_lower in cred_name:
                            username = cred_data["username"]
                            password = cred_data["password"]
                            break

                    # Also check column D for username
                    if len(row) > 3 and row[3]:
                        col_d = str(row[3]).strip()
                        if col_d and col_d.lower() not in [
                            "linux login",
                            "login",
                            "user",
                        ]:
                            if not username:
                                username = col_d

                    # Check column E for password
                    if len(row) > 4 and row[4]:
                        col_e = str(row[4]).strip()
                        if col_e and not password:
                            password = col_e

                    servers.append(
                        {
                            "title": name,
                            "ip1": ip1,
                            "ip2": ip2
                            if re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", ip2)
                            else "",
                            "username": username,
                            "password": password,
                        }
                    )

        if not servers:
            return None

        # Build result table
        result = [["Title", "Username", "Password", "URL", "Notes"]]

        for server in servers:
            url = f"http://{server['ip1']}" if server["ip1"] else ""
            notes = ""
            if server["ip2"]:
                notes = f"IP: {server['ip2']}"

            result.append(
                [server["title"], server["username"], server["password"], url, notes]
            )

        return result if len(result) > 1 else None

    def _preprocess_server_list_format(self, raw_data):
        """
        Handle server list format with Login/Password column.

        Example:
        Row 0-2: headers and metadata
        Row 3: ['NAME', 'HOST', 'IP', 'MAC', 'CPU', 'RAM', 'SSD', 'Login/Password']
        Row 4: ['ingress', '', '172.16.37.2', ...]
        Row 5: ['k8s-dev-hp1', 'host-vm7', '172.16.37.3', ..., 'sa / xx7by4a5']
        ...

        Returns: Records for each server with credentials from Login/Password column
        """
        import re

        if len(raw_data) < 4:
            return None

        # Find header row (contains 'NAME' or 'Login')
        header_row_idx = -1
        for i, row in enumerate(raw_data[:10]):
            if not row:
                continue
            row_str = " ".join(str(cell).lower() for cell in row if cell)
            if "name" in row_str and (
                "login" in row_str or "password" in row_str or "ip" in row_str
            ):
                header_row_idx = i
                break

        if header_row_idx == -1:
            return None

        # Find Login/Password column index
        header_row = raw_data[header_row_idx]
        login_pass_col = -1
        name_col = 0
        ip_col = -1
        host_col = -1

        for j, cell in enumerate(header_row):
            if cell:
                cell_lower = str(cell).lower()
                if (
                    "login" in cell_lower
                    or "password" in cell_lower
                    or cell_lower == "login/password"
                ):
                    login_pass_col = j
                elif cell_lower == "name":
                    name_col = j
                elif cell_lower == "ip":
                    ip_col = j
                elif cell_lower == "host":
                    host_col = j

        # Must have NAME and Login/Password columns
        if name_col == -1 or login_pass_col == -1:
            return None

        # Parse server rows
        servers = []
        for i in range(header_row_idx + 1, len(raw_data)):
            row = raw_data[i]
            if not row:
                continue

            # Get values from appropriate columns
            name = (
                str(row[name_col]).strip()
                if len(row) > name_col and row[name_col]
                else ""
            )

            # Skip empty or header-like rows
            if not name or name.lower() in [
                "name",
                "важливі",
                "vulnerability",
                "діапазон",
                "my-jenkins",
            ]:
                continue

            # Get IP
            ip = ""
            if ip_col >= 0 and len(row) > ip_col and row[ip_col]:
                ip_val = str(row[ip_col]).strip()
                if re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", ip_val):
                    ip = ip_val

            # Get Host
            host = ""
            if host_col >= 0 and len(row) > host_col and row[host_col]:
                host = str(row[host_col]).strip()

            # Get Login/Password
            username = ""
            password = ""
            if (
                login_pass_col >= 0
                and len(row) > login_pass_col
                and row[login_pass_col]
            ):
                login_pass = str(row[login_pass_col]).strip()

                # Parse "user / pass" or "user/pass" pattern
                if " / " in login_pass:
                    parts = login_pass.split(" / ")
                    if len(parts) >= 2:
                        username = parts[0].strip()
                        password = parts[1].strip()
                elif "/" in login_pass:
                    parts = login_pass.split("/")
                    if len(parts) >= 2:
                        username = parts[0].strip()
                        password = parts[1].strip()

            # Only add if we have a valid server name
            if name and len(name) > 2:
                servers.append(
                    {
                        "title": name,
                        "host": host,
                        "ip": ip,
                        "username": username,
                        "password": password,
                    }
                )

        if not servers:
            return None

        # Build result table
        result = [["Title", "Username", "Password", "URL", "Notes"]]

        for server in servers:
            url = f"http://{server['ip']}" if server["ip"] else ""
            notes = ""
            if server["host"]:
                notes = f"Host: {server['host']}"

            result.append(
                [server["title"], server["username"], server["password"], url, notes]
            )

        return result if len(result) > 1 else None

    def parse_without_ai(self):
        """Parse Excel without AI assistance - use preprocessing functions directly."""
        if not self.source_file:
            return

        try:
            parser = ExcelParser(
                self.source_file, groq_api_key=None, model=self.model_var.get()
            )
            parser.load_excel()

            # Try all preprocessing formats in order of specificity
            processed_data = None

            # Try Format 1: Server list with Login/Password column (MOST SPECIFIC)
            processed_data = self._preprocess_server_list_format(parser.raw_data)

            # Try Format 2: Kubernetes cluster format
            if not processed_data:
                processed_data = self._preprocess_k8s_format(parser.raw_data)

            # Try Format 3: Complex table with separate credentials
            if not processed_data:
                processed_data = self._preprocess_complex_table_format(parser.raw_data)

            # Try Format 4: Column-based (multiple servers in rows)
            if not processed_data:
                processed_data = self._preprocess_vertical_format(parser.raw_data)

            # Try Format 5: Scattered row-based data
            if not processed_data:
                processed_data = self._preprocess_scattered_format(parser.raw_data)

            # Try Format 6: Table format with multiple server rows (LEAST SPECIFIC)
            if not processed_data:
                processed_data = self._preprocess_table_format(parser.raw_data)

            if processed_data:
                # Convert to Passbolt format
                header = processed_data[
                    0
                ]  # ['Title', 'Username', 'Password', 'URL', 'Notes']

                self.parsed_data = []
                for row in processed_data[1:]:
                    record = {
                        "Group": "Imported",
                        "Title": row[0] if len(row) > 0 else "",
                        "Username": row[1] if len(row) > 1 else "",
                        "Password": row[2] if len(row) > 2 else "",
                        "URL": row[3] if len(row) > 3 else "",
                        "Notes": row[4] if len(row) > 4 else "",
                    }
                    if record["Title"]:  # Only add if has title
                        self.parsed_data.append(record)

                self.update_result_ui()
                messagebox.showinfo(
                    "Готово", f"Оброблено {len(self.parsed_data)} записів"
                )
            else:
                # Fallback to original parser with AI
                self.parsed_data = parser.parse(use_ai=self.use_ai_detection)
                self.update_result_ui()
                messagebox.showinfo(
                    "Готово", f"Оброблено {len(self.parsed_data)} записів"
                )

        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка парсингу: {e}")

    def start_ai_analysis(self):
        if not self.use_ai_detection:
            messagebox.showwarning(
                "AI Вимкнено", "Увімкніть AI в налаштуваннях для цього режиму."
            )
            return
        if not self.groq_key:
            messagebox.showwarning("API Ключ", "Вкажіть Groq API Key у налаштуваннях.")
            return

        self.btn_analyze.configure(state="disabled", text="Обробка...")
        self.lbl_progress.configure(text="Триває AI аналіз структури таблиці...")
        threading.Thread(target=self.run_analysis, daemon=True).start()

    def run_analysis(self):
        try:
            parser = ExcelParser(
                self.source_file,
                groq_api_key=self.groq_key if self.use_ai_detection else None,
                model=self.model_var.get(),
            )
            parser.load_excel()

            # Debug: check if data was loaded
            if not parser.raw_data:
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Помилка", "Не вдалося прочитати дані з файлу"
                    ),
                )
                self.after(
                    0,
                    lambda: self.btn_analyze.configure(
                        state="normal", text="AI Аналіз та Конвертація"
                    ),
                )
                return

            # Pre-process: detect and handle multiple format types
            processed_data = None

            # Try Format 1: Column-based (multiple servers in rows)
            processed_data = self._preprocess_vertical_format(parser.raw_data)

            # Try Format 2: Scattered row-based data
            if not processed_data:
                processed_data = self._preprocess_scattered_format(parser.raw_data)

            # Try Format 3: Standard table with multiple server rows
            if not processed_data:
                processed_data = self._preprocess_table_format(parser.raw_data)

            # Use processed data if transformation happened
            if processed_data:
                raw_data_str = ""
                for i, row in enumerate(processed_data):
                    raw_data_str += f"Row {i}: {row}\n"
                total_rows = len(processed_data)
            else:
                # No transformation needed, use original
                raw_data_str = ""
                rows_to_show = min(80, len(parser.raw_data))
                for i, row in enumerate(parser.raw_data[:rows_to_show]):
                    raw_data_str += f"Row {i}: {row}\n"
                total_rows = len(parser.raw_data)

            # AI prompt in English for better understanding
            prompt = f"""
You are an expert at parsing Excel tables for Passbolt password manager import.

Passbolt CSV format has these fields:
- Group: group/category/environment (e.g., "Production", "Development", "Database")
- Title: resource name (REQUIRED - must have a value)
- Username: login/username/account name
- Password: password/secret (can be empty)
- URL: resource address (IP, domain, with http:// or https://)
- Notes: additional info, descriptions, comments, metadata

PREPROCESSED DATA ({total_rows} rows):
{raw_data_str}

The data has been pre-processed into a standard table format with columns:
[Title, Username, Password, URL, Notes]

TASK:
1. The first row is the header
2. For EACH subsequent row, create a Passbolt record
3. Keep ALL records - even if Password is empty (user can fill it later)
4. Preserve all data as-is from the preprocessed table

Return ONLY a JSON object:
{
                "records": [
  {
                    "Group": "Imported", "Title": "Server Name", "Username": "admin", "Password": "pass123", "URL": "http://192.168.1.1", "Notes": "Main server"}
]}

IMPORTANT RULES:
- Create ONE record for EACH data row (do NOT skip any)
- Title is REQUIRED - all rows have titles already extracted
- Password CAN BE EMPTY - still create the record
- Keep Group as "Imported"
- Keep URLs and IPs as they are (already formatted)
- Preserve all Notes content

DO NOT skip records with empty passwords - include ALL records!
"""

            client = Groq(api_key=self.groq_key)

            self.lbl_progress.configure(text="AI аналізує структуру даних...")

            model_name = self.model_var.get()

            chat = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model=model_name,
                response_format={"type": "json_object"},
                temperature=0.1,
            )

            self.lbl_progress.configure(text="Обробка результатів...")

            result = json.loads(chat.choices[0].message.content)

            # Debug output
            print(f"AI Response: {chat.choices[0].message.content[:500]}...")

            # Handle different response formats
            if isinstance(result, list):
                self.parsed_data = result
            elif isinstance(result, dict):
                if "resources" in result:
                    self.parsed_data = result["resources"]
                elif "data" in result:
                    self.parsed_data = result["data"]
                elif "records" in result:
                    self.parsed_data = result["records"]
                else:
                    self.parsed_data = [result]

            # Debug: show what we got
            print(f"Parsed {len(self.parsed_data)} records")
            if self.parsed_data:
                print(f"First record: {self.parsed_data[0]}")

            self.after(0, self.update_result_ui)

        except Exception as e:
            import traceback

            error_details = traceback.format_exc()
            print(f"Error: {error_details}")
            self.after(
                0,
                lambda: messagebox.showerror(
                    "Помилка", f"Помилка AI: {e}\n\nДеталі: {error_details}"
                ),
            )
            self.after(
                0,
                lambda: self.btn_analyze.configure(
                    state="normal", text="AI Аналіз та Конвертація"
                ),
            )
            self.after(0, self.stop_progress)

    def stop_progress(self):
        """Stop progress bar and clear label."""
        self.progress_bar.stop()
        self.lbl_progress.configure(text="")

    def update_result_ui(self):
        # Stop progress bar
        self.after(0, self.stop_progress)

        for widget in self.result_frame.winfo_children():
            widget.destroy()

        if self.parsed_data:
            # Show summary
            summary = ctk.CTkLabel(
                self.result_frame,
                text=f"✅ Знайдено {len(self.parsed_data)} записів",
                font=("Arial", 12, "bold"),
                text_color="#2D8C3C",
            )
            summary.pack(fill="x", padx=10, pady=5)

            # Show first 15 records
            for i, row in enumerate(self.parsed_data[:15]):
                title = row.get("Title", "N/A")
                username = row.get("Username", "")
                url = row.get("URL", "")
                group = row.get("Group", "")

                txt = f"{i + 1}. [{group}] {title}"
                if username:
                    txt += f" | 👤 {username}"
                if url:
                    txt += f" | 🔗 {url}"

                lbl = ctk.CTkLabel(
                    self.result_frame, text=txt, anchor="w", font=("Arial", 11)
                )
                lbl.pack(fill="x", padx=10, pady=2)

            if len(self.parsed_data) > 15:
                lbl_more = ctk.CTkLabel(
                    self.result_frame,
                    text=f"⋯ ще {len(self.parsed_data) - 15} записів (будуть у CSV)",
                    text_color="gray",
                    font=("Arial", 10),
                )
                lbl_more.pack(padx=10, pady=5)
        else:
            no_data = ctk.CTkLabel(
                self.result_frame,
                text="⚠️ Не знайдено даних для імпорту",
                text_color="orange",
                font=("Arial", 12),
            )
            no_data.pack(padx=10, pady=20)

        self.btn_analyze.configure(state="normal", text="AI Аналіз та Конвертація")
        self.lbl_progress.configure(text="")
        self.btn_export.configure(state="normal")

    def export_csv(self):
        if not self.parsed_data:
            messagebox.showwarning(
                "Помилка", "Немає даних для експорту.\nСпочатку проаналізуйте файл."
            )
            return

        import csv

        cleaned = []
        for row in self.parsed_data:
            if not isinstance(row, dict):
                continue

            title = row.get("Title", "")
            if not title:
                continue
            title = str(title).strip()

            # Get and clean other fields
            username = str(row.get("Username", "") or "").strip()
            password = str(row.get("Password", "") or "").strip()
            url = str(row.get("URL", "") or "").strip()
            group = str(row.get("Group", "Imported") or "Imported").strip()
            notes = str(row.get("Notes", "") or "").strip()

            # Normalize URL
            if url and not url.startswith(("http://", "https://", "ftp://")):
                url = "http://" + url

            cleaned.append(
                {
                    "Group": group or "Imported",
                    "Title": title,
                    "Username": username,
                    "Password": password or " ",
                    "URL": url,
                    "Notes": notes,
                }
            )

        if not cleaned:
            messagebox.showwarning(
                "Помилка", "Не знайдено жодного запису з назвою (Title)."
            )
            return

        final_df = pd.DataFrame(cleaned)[
            ["Group", "Title", "Username", "Password", "URL", "Notes"]
        ]

        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile="passbolt_import.csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if save_path:
            try:
                final_df.to_csv(
                    save_path,
                    index=False,
                    quoting=csv.QUOTE_ALL,
                    encoding="utf-8-sig",
                    lineterminator="\r\n",
                )
                messagebox.showinfo(
                    "Успіх",
                    f"Файл збережено: {save_path}\n\n"
                    f"Записів: {len(cleaned)}\n\n"
                    f"В Passbolt: Імпорт → KeePass (CSV)",
                )
            except Exception as e:
                messagebox.showerror(
                    "Помилка",
                    f"Не вдалося зберегти: {e}\n\nПеревірте чи файл не відкритий в іншій програмі.",
                )


if __name__ == "__main__":
    app = AIPassboltApp()
    app.mainloop()
