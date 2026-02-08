import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk
import aggregator
import os
import threading
import datetime

# --- Translation Dictionary ---
TRANSLATIONS = {
    "JP": {
        "title": "コメントシート集計ツール",
        "step1": "1. 入力データ",
        "select_files": "コメントシートを選択...",
        "no_files": "ファイルが選択されていません",
        "select_att": "出席表を選択 (任意)...",
        "no_att": "出席表が選択されていません",
        "att_selected": "出席表: ",
        "step2": "2. 設定",
        "target_year": "対象年度 (任意):",
        "year_desc": "(この年度で始まる科目を抽出)",
        "step3": "3. 実行",
        "run": "集計開始",
        "ready": "準備完了。",
        "processing": "処理中... 保存先: ",
        "success": "成功: ",
        "success_title": "完了",
        "success_msg": "集計が完了しました！\nファイルが保存されました。",
        "error": "エラー",
        "warning": "警告",
        "select_files_warn": "ファイルを選択してください。",
        "year_empty_title": "警告: 対象年度が空です",
        "year_empty_msg": "対象年度が入力されていません。\n\nヘッダー行や関係ない科目を含む「すべての行」が処理されます。\n\n本当に続行しますか？",
        "files_selected": " 個のファイルを選択",
        "file_selection_cancelled": "ファイルの選択がキャンセルされました。",
        "att_selection_cancelled": "出席表の選択がキャンセルされました。",
        "critical_error": "致命的なエラー: ",
        "save_as": "保存先を指定"
    },
    "EN": {
        "title": "Comment Sheet Aggregator",
        "step1": "1. Input Data",
        "select_files": "Select Comment Sheets...",
        "no_files": "No files selected",
        "select_att": "Select Attendance Sheet (Optional)...",
        "no_att": "No attendance sheet selected",
        "att_selected": "Attendance: ",
        "step2": "2. Settings",
        "target_year": "Target Year (Optional):",
        "year_desc": "(Filters courses starting with this year)",
        "step3": "3. Execution",
        "run": "Run Aggregation",
        "ready": "Ready.",
        "processing": "Processing... saving to ",
        "success": "SUCCESS: ",
        "success_title": "Success",
        "success_msg": "Aggregation complete!\nFile saved.",
        "error": "Error",
        "warning": "Warning",
        "select_files_warn": "Please select files first.",
        "year_empty_title": "Warning: Target Year Empty",
        "year_empty_msg": "Target Year is empty.\n\nThis will process ALL rows in the input files, which might include header rows or unrelated courses.\n\nAre you sure you want to proceed?",
        "files_selected": " files selected",
        "file_selection_cancelled": "File selection cancelled.",
        "att_selection_cancelled": "Attendance selection cancelled.",
        "critical_error": "CRITICAL ERROR: ",
        "save_as": "Save Output As"
    }
}

import platform

class CommentAggregatorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x480") # Compact size
        
        # OS Detection
        self.is_windows = platform.system() == "Windows"
        
        # Style Configuration
        self.style = ttk.Style()
        
        if self.is_windows:
            try:
                self.style.theme_use('vista')
            except:
                self.style.theme_use('clam')
            self.default_font = ("Segoe UI", 9)
            self.header_font = ("Segoe UI", 10, "bold")
        else:
            # Mac / Linux
            self.style.theme_use('clam') # 'clam' looks decent on Mac/Linux usually
            self.default_font = ("Helvetica", 11) # System font
            self.header_font = ("Helvetica", 12, "bold")

        self.style.configure('.', font=self.default_font)
        self.style.configure('TButton', font=self.default_font, padding=3)
        self.style.configure('TLabel', font=self.default_font)
        self.style.configure('TLabelframe.Label', font=self.header_font, foreground="#003399")
        
        # Current Language (Default JP)
        self.lang = "JP"
        
        # Main Container with Padding
        main_frame = ttk.Frame(root, padding="10 10 10 10") # Reduced padding
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Header (Title & Language Switcher) ---
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Language Switcher (Right aligned)
        frame_lang = ttk.Frame(header_frame)
        frame_lang.pack(side=tk.RIGHT)
        ttk.Label(frame_lang, text="Language: ").pack(side=tk.LEFT)
        self.combo_lang = ttk.Combobox(frame_lang, values=["日本語 (JP)", "English (EN)"], state="readonly", width=10)
        self.combo_lang.current(0) # Select JP
        self.combo_lang.pack(side=tk.LEFT)
        self.combo_lang.bind("<<ComboboxSelected>>", self.change_language)

        # Variables
        self.selected_files = []
        self.attendance_file = None
        
        # --- Section 1: Input Files ---
        self.frame_input = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_input.pack(fill=tk.X, pady=(0, 10))
        
        # File Selection
        self.btn_select = ttk.Button(self.frame_input, command=self.select_files)
        self.btn_select.pack(anchor=tk.W)
        
        self.label_file_count = ttk.Label(self.frame_input, foreground="gray")
        self.label_file_count.pack(anchor=tk.W, pady=(2, 0))
        
        # Attendance Selection
        self.btn_attendance = ttk.Button(self.frame_input, command=self.select_attendance)
        self.btn_attendance.pack(anchor=tk.W, pady=(5, 0))
        
        self.label_attendance = ttk.Label(self.frame_input, foreground="gray")
        self.label_attendance.pack(anchor=tk.W, pady=(2, 0))

        # --- Section 2: Settings ---
        self.frame_settings = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_settings.pack(fill=tk.X, pady=(0, 10))
        
        # Target Year
        frame_year = ttk.Frame(self.frame_settings)
        frame_year.pack(fill=tk.X)
        
        self.label_year = ttk.Label(frame_year)
        self.label_year.pack(side=tk.LEFT)
        self.entry_year = ttk.Entry(frame_year, width=8)
        self.entry_year.pack(side=tk.LEFT, padx=5)
        
        # Default Year
        current_year = str(datetime.datetime.now().year)
        self.entry_year.insert(0, current_year)
        
        self.label_year_desc = ttk.Label(frame_year, foreground="gray", font=("Segoe UI", 8))
        self.label_year_desc.pack(side=tk.LEFT)

        # --- Section 3: Action ---
        self.frame_action = ttk.LabelFrame(main_frame, padding="5 5 5 5")
        self.frame_action.pack(fill=tk.BOTH, expand=True)
        
        self.btn_run = ttk.Button(self.frame_action, command=self.run_aggregation, state=tk.DISABLED)
        self.btn_run.pack(fill=tk.X, pady=(0, 5))
        
        # Log Area
        self.log_area = scrolledtext.ScrolledText(self.frame_action, height=6, state='disabled', font=("Consolas", 9), bg="#f8f9fa", relief=tk.FLAT)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        # Apply Text
        self.update_ui_text()
        self.log(TRANSLATIONS[self.lang]["ready"])

    def create_resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def change_language(self, event):
        val = self.combo_lang.get()
        if "JP" in val:
            self.lang = "JP"
        else:
            self.lang = "EN"
        self.update_ui_text()

    def update_ui_text(self):
        t = TRANSLATIONS[self.lang]
        
        self.root.title(t["title"])
        self.frame_input.config(text=t["step1"])
        self.btn_select.config(text=t["select_files"])
        
        if not self.selected_files:
            self.label_file_count.config(text=t["no_files"])
        else:
             self.label_file_count.config(text=f"{len(self.selected_files)}{t['files_selected']}")

        self.btn_attendance.config(text=t["select_att"])
        if not self.attendance_file:
            self.label_attendance.config(text=t["no_att"])
        else:
            self.label_attendance.config(text=f"{t['att_selected']}{os.path.basename(self.attendance_file)}")
            
        self.frame_settings.config(text=t["step2"])
        self.label_year.config(text=t["target_year"])
        self.label_year_desc.config(text=t["year_desc"])
        
        self.frame_action.config(text=t["step3"])
        self.btn_run.config(text=t["run"])

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def select_files(self):
        t = TRANSLATIONS[self.lang]
        files = filedialog.askopenfilenames(
            title=t["select_files"],
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if files:
            self.selected_files = list(files)
            self.label_file_count.config(text=f"{len(self.selected_files)}{t['files_selected']}", foreground="#008800")
            self.btn_run.config(state=tk.NORMAL)
            self.log(f"Selected {len(files)} files.")
        else:
            self.log(t["file_selection_cancelled"])

    def select_attendance(self):
        t = TRANSLATIONS[self.lang]
        file = filedialog.askopenfilename(
            title=t["select_att"],
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.attendance_file = file
            self.label_attendance.config(text=f"{t['att_selected']}{os.path.basename(file)}", foreground="#000088")
            self.log(f"Selected attendance sheet: {os.path.basename(file)}")
        else:
            self.log(t["att_selection_cancelled"])

    def run_aggregation(self):
        t = TRANSLATIONS[self.lang]
        if not self.selected_files:
            messagebox.showwarning(t["warning"], t["select_files_warn"])
            return

        target_year = self.entry_year.get().strip()
        
        # Confirmation for empty year
        if not target_year:
            confirm = messagebox.askyesno(
                t["year_empty_title"], 
                t["year_empty_msg"]
            )
            if not confirm:
                return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="summary_output.xlsx",
            title=t["save_as"]
        )

        if not output_file:
            return

        self.btn_run.config(state=tk.DISABLED)
        self.log(f"{t['processing']}{os.path.basename(output_file)}")
        
        # Run in thread
        threading.Thread(target=self.process_thread, args=(output_file, target_year, self.attendance_file), daemon=True).start()

    def process_thread(self, output_file, target_year, attendance_file):
        t = TRANSLATIONS[self.lang]
        try:
            success, message = aggregator.process_files(self.selected_files, output_file, target_year, attendance_file)
            if success:
                self.log(t["success"] + message)
                messagebox.showinfo(t["success_title"], t["success_msg"])
            else:
                self.log("FAILED: " + message)
                messagebox.showerror(t["error"], message)
        except Exception as e:
            self.log(t["critical_error"] + str(e))
            messagebox.showerror(t["error"], str(e))
        finally:
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))

if __name__ == "__main__":
    root = tk.Tk()
    app = CommentAggregatorApp(root)
    root.mainloop()
