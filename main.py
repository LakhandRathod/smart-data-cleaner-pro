import json
import queue
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

# Optional drag-and-drop support:
# pip install tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD # type: ignore
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = None
    DND_FILES = None


APP_DIR = Path(__file__).resolve().parent
HISTORY_FILE = APP_DIR / "cleaning_history.json"


class DataCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Data Cleaner Pro")
        self.root.geometry("1250x820")
        self.root.minsize(1000, 680)

        self.input_path = None
        self.output_dir = None
        self.original_df = None
        self.cleaned_df = None
        self.worker_queue = queue.Queue()
        self.cleaning = False
        self.history = self.load_history()

        self.dark_mode = True
        self.apply_theme()
        self.build_ui()
        self.root.after(100, self.process_queue)

    # ---------------- Theme and layout ----------------

    def apply_theme(self):
        ctk.set_appearance_mode("dark" if self.dark_mode else "light")
        ctk.set_default_color_theme("blue")
        self.root.configure(bg="#0b1120" if self.dark_mode else "#eef2f7")

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        self.theme_button.configure(
            text="☀ Light mode" if self.dark_mode else "🌙 Dark mode"
        )
        self.set_status(
            "Dark theme enabled." if self.dark_mode else "Light theme enabled."
        )

    def build_ui(self):
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(
            self.root, width=215, corner_radius=0, fg_color=("#e4eaf2", "#111b2e")
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        ctk.CTkLabel(
            self.sidebar, text="⚡ SMART CLEANER",
            font=("Arial", 21, "bold"), text_color="#38bdf8"
        ).pack(anchor="w", padx=20, pady=(30, 4))
        ctk.CTkLabel(
            self.sidebar, text="DATA AUTOMATION PRO",
            font=("Arial", 10, "bold"), text_color=("#64748b", "#94a3b8")
        ).pack(anchor="w", padx=21, pady=(0, 30))

        self.nav_buttons = {}
        for key, label in [
            ("Dashboard", "🏠  Dashboard"),
            ("Preview", "📄  Data Preview"),
            ("Compare", "↔  Before / After"),
            ("History", "🕘  Cleaning History"),
        ]:
            btn = ctk.CTkButton(
                self.sidebar, text=label, anchor="w", height=43,
                fg_color="transparent", hover_color=("#d4deeb", "#263750"),
                text_color=("#1e293b", "#dbeafe"),
                command=lambda page=key: self.show_page(page)
            )
            btn.pack(fill="x", padx=12, pady=4)
            self.nav_buttons[key] = btn

        self.theme_button = ctk.CTkButton(
            self.sidebar, text="☀ Light mode", height=38,
            fg_color="#2563eb", hover_color="#1d4ed8",
            command=self.toggle_theme
        )
        self.theme_button.pack(side="bottom", fill="x", padx=14, pady=(8, 12))

        ctk.CTkLabel(
            self.sidebar, text="Version 2.0 • Local processing",
            font=("Arial", 10), text_color=("#64748b", "#94a3b8")
        ).pack(side="bottom", pady=8)

        self.content = ctk.CTkFrame(self.root, fg_color="transparent")
        self.content.grid(row=0, column=1, sticky="nsew", padx=25, pady=22)
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        self.header = ctk.CTkFrame(self.content, fg_color="transparent")
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        self.header.grid_columnconfigure(0, weight=1)
        self.page_title = ctk.CTkLabel(
            self.header, text="Data Cleaning Dashboard",
            font=("Arial", 27, "bold"), anchor="w"
        )
        self.page_title.grid(row=0, column=0, sticky="w")
        self.page_subtitle = ctk.CTkLabel(
            self.header, text="Clean, compare and export your Excel data.",
            font=("Arial", 13), text_color=("#64748b", "#94a3b8"), anchor="w"
        )
        self.page_subtitle.grid(row=1, column=0, sticky="w", pady=(4, 0))

        self.pages = {}
        self.pages["Dashboard"] = self.build_dashboard()
        self.pages["Preview"] = self.build_table_page("Data Preview", "Original file preview")
        self.pages["Compare"] = self.build_compare_page()
        self.pages["History"] = self.build_history_page()

        self.show_page("Dashboard")

    def page_frame(self):
        frame = ctk.CTkScrollableFrame(
            self.content, fg_color="transparent", corner_radius=0
        )
        frame.grid(row=1, column=0, sticky="nsew")
        return frame

    def card(self, parent, title):
        frame = ctk.CTkFrame(parent, corner_radius=14, fg_color=("white", "#172338"))
        ctk.CTkLabel(frame, text=title, font=("Arial", 17, "bold")).pack(
            anchor="w", padx=18, pady=(16, 12)
        )
        return frame

    def build_dashboard(self):
        page = self.page_frame()

        upload = self.card(page, "📂  Input Excel / CSV File")
        upload.pack(fill="x", pady=(0, 13))
        row = ctk.CTkFrame(upload, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(0, 16))
        self.file_entry = ctk.CTkEntry(
            row, height=40, placeholder_text="Choose .xlsx, .xls or .csv file..."
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 9))
        ctk.CTkButton(row, text="Browse file", width=120, command=self.browse_file).pack(side="right")

        self.drop_label = ctk.CTkLabel(
            upload, text="Or drag and drop an Excel/CSV file here",
            height=52, corner_radius=10, fg_color=("#eaf4ff", "#203550"),
            text_color=("#1d4ed8", "#bfdbfe"), font=("Arial", 13, "bold")
        )
        self.drop_label.pack(fill="x", padx=18, pady=(0, 16))
        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self.drop_file)
        else:
            self.drop_label.configure(
                text="Drag & drop needs tkinterdnd2. You can use Browse file below."
            )

        output = self.card(page, "💾  Output Folder")
        output.pack(fill="x", pady=(0, 13))
        outrow = ctk.CTkFrame(output, fg_color="transparent")
        outrow.pack(fill="x", padx=18, pady=(0, 16))
        self.output_entry = ctk.CTkEntry(
            outrow, height=40, placeholder_text="Choose where cleaned files will be saved..."
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 9))
        ctk.CTkButton(outrow, text="Browse folder", width=125, command=self.browse_output).pack(side="right")

        options = self.card(page, "⚙  Cleaning Options")
        options.pack(fill="x", pady=(0, 13))
        self.duplicate_var = tk.BooleanVar(value=True)
        self.blank_var = tk.BooleanVar(value=True)
        self.space_var = tk.BooleanVar(value=True)
        self.title_var = tk.BooleanVar(value=False)
        self.na_var = tk.BooleanVar(value=False)

        checks = ctk.CTkFrame(options, fg_color="transparent")
        checks.pack(fill="x", padx=14, pady=(0, 14))
        options_list = [
            ("Remove duplicate rows", self.duplicate_var),
            ("Remove completely blank rows", self.blank_var),
            ("Trim leading/trailing spaces", self.space_var),
            ("Convert text to Title Case", self.title_var),
            ("Remove rows containing any missing value", self.na_var),
        ]
        for i, (label, variable) in enumerate(options_list):
            cb = ctk.CTkCheckBox(checks, text=label, variable=variable)
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=12, pady=9)

        run_card = self.card(page, "🚀  Run Cleaning")
        run_card.pack(fill="x", pady=(0, 13))
        self.progress = ctk.CTkProgressBar(run_card, height=12, progress_color="#0ea5e9")
        self.progress.pack(fill="x", padx=18, pady=(0, 7))
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(
            run_card, text="Ready", text_color=("#64748b", "#94a3b8")
        )
        self.progress_label.pack(anchor="w", padx=18, pady=(0, 12))

        buttons = ctk.CTkFrame(run_card, fg_color="transparent")
        buttons.pack(anchor="w", padx=14, pady=(0, 15))
        self.clean_button = ctk.CTkButton(
            buttons, text="✨  Clean Data", width=160, height=42,
            font=("Arial", 14, "bold"), command=self.start_cleaning
        )
        self.clean_button.pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="Preview Original", width=145, height=42,
            fg_color="#475569", hover_color="#334155",
            command=lambda: self.show_page("Preview")
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="Reset", width=90, height=42,
            fg_color="#475569", hover_color="#334155", command=self.reset
        ).pack(side="left", padx=4)

        self.summary_label = ctk.CTkLabel(
            page, text="No file loaded yet.", anchor="w",
            text_color=("#64748b", "#94a3b8")
        )
        self.summary_label.pack(fill="x", pady=8)
        return page

    def build_table_page(self, title, subtitle):
        page = self.page_frame()
        top = ctk.CTkFrame(page, fg_color="transparent")
        top.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(top, text=subtitle, font=("Arial", 13)).pack(side="left")
        ctk.CTkButton(
            top, text="Refresh preview", width=130,
            command=lambda: self.populate_tree(self.preview_tree, self.original_df)
        ).pack(side="right")
        holder, tree = self.make_tree(page)
        holder.pack(fill="both", expand=True)
        self.preview_tree = tree
        return page

    def build_compare_page(self):
        page = self.page_frame()
        self.compare_stats = ctk.CTkLabel(
            page, text="Clean a file to see the before/after comparison.",
            font=("Arial", 14), anchor="w"
        )
        self.compare_stats.pack(fill="x", pady=(0, 12))
        tabs = ctk.CTkTabview(page)
        tabs.pack(fill="both", expand=True)
        tabs.add("Before")
        tabs.add("After")
        before_holder, self.before_tree = self.make_tree(tabs.tab("Before"))
        after_holder, self.after_tree = self.make_tree(tabs.tab("After"))
        before_holder.pack(fill="both", expand=True, padx=5, pady=5)
        after_holder.pack(fill="both", expand=True, padx=5, pady=5)
        return page

    def build_history_page(self):
        page = self.page_frame()
        ctk.CTkLabel(
            page, text="Previous cleaning runs are recorded on this computer.",
            font=("Arial", 13), anchor="w"
        ).pack(fill="x", pady=(0, 12))
        self.history_box = ctk.CTkTextbox(page, height=430, corner_radius=12)
        self.history_box.pack(fill="both", expand=True)
        row = ctk.CTkFrame(page, fg_color="transparent")
        row.pack(fill="x", pady=10)
        ctk.CTkButton(row, text="Refresh history", command=self.refresh_history).pack(side="left", padx=4)
        ctk.CTkButton(
            row, text="Clear history", fg_color="#b91c1c",
            hover_color="#991b1b", command=self.clear_history
        ).pack(side="left", padx=4)
        self.refresh_history()
        return page

    def make_tree(self, parent):
        holder = ctk.CTkFrame(parent, corner_radius=10)
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(holder, show="headings")
        ybar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        xbar = ttk.Scrollbar(holder, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview", rowheight=28, font=("Arial", 10),
            background="#172338", foreground="#e2e8f0",
            fieldbackground="#172338"
        )
        style.configure(
            "Treeview.Heading", font=("Arial", 10, "bold"),
            background="#263750", foreground="#ffffff"
        )
        style.map("Treeview", background=[("selected", "#1d4ed8")])
        return holder, tree

    # ---------------- Navigation ----------------

    def show_page(self, name):
        for page in self.pages.values():
            page.grid_remove()
        self.pages[name].grid()
        titles = {
            "Dashboard": ("Data Cleaning Dashboard", "Clean, compare and export your data."),
            "Preview": ("Original Data Preview", "Inspect the source data before cleaning."),
            "Compare": ("Before / After Comparison", "Compare the original and cleaned datasets."),
            "History": ("Cleaning History", "Review previous cleaning runs."),
        }
        self.page_title.configure(text=titles[name][0])
        self.page_subtitle.configure(text=titles[name][1])
        for key, btn in self.nav_buttons.items():
            btn.configure(fg_color="#263750" if key == name else "transparent")
        if name == "Preview":
            self.populate_tree(self.preview_tree, self.original_df)
        elif name == "Compare":
            self.populate_tree(self.before_tree, self.original_df)
            self.populate_tree(self.after_tree, self.cleaned_df)

    # ---------------- File loading ----------------

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            self.load_file(path)

    def drop_file(self, event):
        # tkinterdnd2 may wrap paths with braces when names contain spaces.
        path = event.data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        self.load_file(path)

    def load_file(self, path):
        try:
            path = Path(path)
            if not path.is_file():
                raise FileNotFoundError("Selected file does not exist.")
            if path.suffix.lower() == ".csv":
                df = pd.read_csv(path)
            elif path.suffix.lower() in (".xlsx", ".xls"):
                df = pd.read_excel(path)
            else:
                raise ValueError("Choose an Excel (.xlsx/.xls) or CSV (.csv) file.")

            self.input_path = path
            self.original_df = df.copy()
            self.cleaned_df = None
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, str(path))
            self.summary_label.configure(
                text=f"Loaded: {path.name}   •   {len(df):,} rows   •   {len(df.columns)} columns"
            )
            self.set_status(f"Loaded {path.name}")
            self.populate_tree(self.preview_tree, self.original_df)
            self.populate_tree(self.before_tree, self.original_df)
            self.populate_tree(self.after_tree, None)
            self.toast("File loaded successfully")
        except Exception as exc:
            messagebox.showerror("File loading error", str(exc))

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = Path(folder)
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)

    # ---------------- Preview and comparison ----------------

    def populate_tree(self, tree, df):
        tree.delete(*tree.get_children())
        if df is None:
            tree["columns"] = ()
            return
        columns = [str(c) for c in df.columns]
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=145, minwidth=90, stretch=True)
        # Limit display rows so very large workbooks remain responsive.
        for _, row in df.head(300).iterrows():
            values = [self.display_value(v) for v in row.tolist()]
            tree.insert("", "end", values=values)
        if len(df) > 300:
            self.set_status("Preview shows the first 300 rows for performance.")

    @staticmethod
    def display_value(value):
        if pd.isna(value):
            return ""
        return str(value)

    def update_comparison(self):
        if self.original_df is None or self.cleaned_df is None:
            return
        before_rows, after_rows = len(self.original_df), len(self.cleaned_df)
        removed = before_rows - after_rows
        self.compare_stats.configure(
            text=f"Rows before: {before_rows:,}     Rows after: {after_rows:,}     "
                 f"Rows removed: {removed:,}     Columns: {len(self.cleaned_df.columns)}"
        )
        self.populate_tree(self.before_tree, self.original_df)
        self.populate_tree(self.after_tree, self.cleaned_df)

    # ---------------- Cleaning worker ----------------

    def start_cleaning(self):
        if self.cleaning:
            return
        if self.original_df is None or self.input_path is None:
            messagebox.showwarning("Select a file", "Please load an Excel or CSV file first.")
            return
        if not self.output_entry.get().strip():
            messagebox.showwarning("Output folder", "Please choose an output folder.")
            return
        output_dir = Path(self.output_entry.get().strip())
        if not output_dir.is_dir():
            messagebox.showerror("Invalid folder", "The output folder does not exist.")
            return

        # Read Tk variables on the main thread, then pass plain Python values to worker.
        options = {
            "duplicates": bool(self.duplicate_var.get()),
            "blank": bool(self.blank_var.get()),
            "spaces": bool(self.space_var.get()),
            "title": bool(self.title_var.get()),
            "missing": bool(self.na_var.get()),
        }
        source = self.original_df.copy()
        self.cleaning = True
        self.clean_button.configure(state="disabled", text="Cleaning...")
        self.progress.set(0)
        self.progress_label.configure(text="Starting...")
        threading.Thread(
            target=self.clean_worker,
            args=(source, options, output_dir, self.input_path),
            daemon=True
        ).start()

    def clean_worker(self, df, options, output_dir, input_path):
        try:
            original_rows = len(df)
            details = []

            self.worker_queue.put(("progress", 0.12, "Checking data..."))
            time.sleep(0.15)

            if options["duplicates"]:
                old = len(df)
                df = df.drop_duplicates()
                details.append(f"Duplicate rows removed: {old - len(df)}")
            self.worker_queue.put(("progress", 0.30, "Removing duplicates..."))

            if options["blank"]:
                old = len(df)
                df = df.dropna(how="all")
                details.append(f"Completely blank rows removed: {old - len(df)}")
            self.worker_queue.put(("progress", 0.48, "Removing blank rows..."))

            if options["spaces"]:
                text_cols = df.select_dtypes(include=["object", "string"]).columns
                for col in text_cols:
                    df[col] = df[col].map(
                        lambda x: x.strip() if isinstance(x, str) else x
                    )
                details.append("Leading and trailing spaces trimmed.")
            self.worker_queue.put(("progress", 0.65, "Trimming text spaces..."))

            if options["title"]:
                text_cols = df.select_dtypes(include=["object", "string"]).columns
                for col in text_cols:
                    df[col] = df[col].map(
                        lambda x: x.title() if isinstance(x, str) else x
                    )
                details.append("Title Case applied to text columns.")
            self.worker_queue.put(("progress", 0.80, "Applying text formatting..."))

            if options["missing"]:
                old = len(df)
                df = df.dropna()
                details.append(f"Rows with any missing value removed: {old - len(df)}")

            output_dir.mkdir(parents=True, exist_ok=True)
            output_xlsx = output_dir / f"{input_path.stem}_cleaned.xlsx"
            output_csv = output_dir / f"{input_path.stem}_cleaned.csv"
            df.to_excel(output_xlsx, index=False)
            df.to_csv(output_csv, index=False, encoding="utf-8-sig")

            details.insert(0, f"Original rows: {original_rows:,}")
            details.append(f"Final rows: {len(df):,}")
            details.append(f"Total rows removed: {original_rows - len(df):,}")
            details.append(f"Excel saved: {output_xlsx}")
            details.append(f"CSV saved: {output_csv}")

            record = {
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "source": str(input_path),
                "excel_output": str(output_xlsx),
                "csv_output": str(output_csv),
                "rows_before": original_rows,
                "rows_after": len(df),
                "rows_removed": original_rows - len(df),
            }
            self.worker_queue.put(("done", df, details, record))
        except Exception as exc:
            self.worker_queue.put(("error", str(exc)))

    def process_queue(self):
        try:
            while True:
                item = self.worker_queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, value, message = item
                    self.progress.set(value)
                    self.progress_label.configure(text=message)
                elif kind == "done":
                    _, df, details, record = item
                    self.cleaned_df = df
                    self.cleaning = False
                    self.clean_button.configure(state="normal", text="✨  Clean Data")
                    self.progress.set(1)
                    self.progress_label.configure(text="✓ Cleaning completed!")
                    self.set_status("Cleaning completed successfully.")
                    self.update_comparison()
                    self.refresh_history()
                    self.history.insert(0, record)
                    self.save_history()
                    self.refresh_history()
                    self.toast("Cleaning completed • Excel and CSV exported")
                    messagebox.showinfo(
                        "Cleaning complete",
                        "Your cleaned files have been saved:\n\n"
                        + record["excel_output"] + "\n" + record["csv_output"]
                    )
                    self.show_page("Compare")
                elif kind == "error":
                    self.cleaning = False
                    self.clean_button.configure(state="normal", text="✨  Clean Data")
                    self.progress.set(0)
                    self.progress_label.configure(text="Cleaning failed.")
                    messagebox.showerror("Cleaning error", item[1])
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    # ---------------- History ----------------

    def load_history(self):
        try:
            if HISTORY_FILE.exists():
                return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
        return []

    def save_history(self):
        try:
            HISTORY_FILE.write_text(
                json.dumps(self.history[:100], indent=2),
                encoding="utf-8"
            )
        except Exception as exc:
            self.set_status(f"Could not save history: {exc}")

    def refresh_history(self):
        if not hasattr(self, "history_box"):
            return
        self.history_box.delete("1.0", "end")
        if not self.history:
            self.history_box.insert("end", "No cleaning history yet.\n")
            return
        for i, item in enumerate(self.history, start=1):
            self.history_box.insert(
                "end",
                f"{i}. {item.get('date', 'Unknown date')}\n"
                f"   Source: {item.get('source', '')}\n"
                f"   Rows: {item.get('rows_before', 0):,} → "
                f"{item.get('rows_after', 0):,} "
                f"(removed {item.get('rows_removed', 0):,})\n"
                f"   Excel: {item.get('excel_output', '')}\n"
                f"   CSV: {item.get('csv_output', '')}\n\n"
            )

    def clear_history(self):
        if messagebox.askyesno("Clear history", "Delete all saved cleaning history?"):
            self.history = []
            self.save_history()
            self.refresh_history()

    # ---------------- Helpers ----------------

    def set_status(self, message):
        self.progress_label.configure(text=message)

    def toast(self, message):
        # Lightweight temporary in-app notification.
        popup = ctk.CTkToplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(fg_color="#166534")
        ctk.CTkLabel(
            popup, text="✓  " + message, text_color="white",
            font=("Arial", 12, "bold"), padx=18, pady=12
        ).pack()
        self.root.update_idletasks()
        x = self.root.winfo_rootx() + self.root.winfo_width() - 390
        y = self.root.winfo_rooty() + 35
        popup.geometry(f"+{x}+{y}")
        popup.after(2400, popup.destroy)

    def reset(self):
        if self.cleaning:
            messagebox.showwarning("Busy", "Wait until the current cleaning task finishes.")
            return
        self.input_path = None
        self.output_dir = None
        self.original_df = None
        self.cleaned_df = None
        self.file_entry.delete(0, "end")
        self.output_entry.delete(0, "end")
        self.progress.set(0)
        self.progress_label.configure(text="Ready")
        self.summary_label.configure(text="No file loaded yet.")
        self.populate_tree(self.preview_tree, None)
        self.populate_tree(self.before_tree, None)
        self.populate_tree(self.after_tree, None)
        self.compare_stats.configure(text="Clean a file to see the before/after comparison.")
        self.set_status("Application reset.")


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
        # CTk widgets can be hosted in a TkinterDnD root.
        ctk.set_appearance_mode("dark")
        app = DataCleanerApp(root)
    else:
        root = ctk.CTk()
        app = DataCleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
