import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import math
import os
import time


# ── College Map (acronym → college_id from CollegeSeeder) ───────
COLLEGE_MAP = {
    'CON':  1,
    'COT':  2,
    'CAS':  3,
    'CPAG': 4,
    'COB':  5,
    'COE':  6,
}
EMAIL_DOMAIN = "@student.buksu.edu.ph"

# ── Theme Colors ────────────────────────────────────────────────
BG        = "#0f1117"
BG2       = "#1a1d27"
BG3       = "#22263a"
ACCENT    = "#4f8ef7"
ACCENT2   = "#7c3aed"
SUCCESS   = "#22c55e"
WARNING   = "#f59e0b"
DANGER    = "#ef4444"
TEXT      = "#e2e8f0"
TEXT_DIM  = "#64748b"
BORDER    = "#2d3352"
FONT_MAIN = ("Segoe UI", 10)
FONT_H1   = ("Segoe UI", 18, "bold")
FONT_H2   = ("Segoe UI", 13, "bold")
FONT_MONO = ("Consolas", 9)


class SQLConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("COMELEC · SQL Converter")
        self.geometry("820x680")
        self.minsize(700, 580)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.input_file  = tk.StringVar()
        self.output_dir  = tk.StringVar(value=os.path.join(os.getcwd(), "sql_chunks"))
        self.table_name  = tk.StringVar(value="users")
        self.chunk_size  = tk.IntVar(value=1000)
        self.role_val    = tk.StringVar(value="voter")
        self.status_val  = tk.StringVar(value="active")
        self.progress    = tk.DoubleVar(value=0)
        self._running    = False

        self._build_ui()

    # ── UI BUILD ────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=BG, pady=0)
        header.pack(fill="x", padx=0, pady=0)

        gradient_bar = tk.Frame(header, bg=ACCENT, height=3)
        gradient_bar.pack(fill="x")

        title_row = tk.Frame(header, bg=BG, padx=28, pady=16)
        title_row.pack(fill="x")

        tk.Label(title_row, text="⚡", font=("Segoe UI", 22), bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(title_row, text=" SQL Converter", font=FONT_H1, bg=BG, fg=TEXT).pack(side="left")
        tk.Label(title_row, text="  COMELEC 2026", font=("Segoe UI", 10), bg=BG, fg=TEXT_DIM).pack(side="left", pady=(6, 0))

        # Main body
        body = tk.Frame(self, bg=BG, padx=24, pady=4)
        body.pack(fill="both", expand=True)

        # ── File Selection Card
        self._card(body, "📂  File Selection", self._build_file_section).pack(fill="x", pady=(0, 12))

        # ── Settings Card
        self._card(body, "⚙️  Settings", self._build_settings_section).pack(fill="x", pady=(0, 12))

        # ── Preview Card
        self._card(body, "👁  Preview", self._build_preview_section).pack(fill="both", expand=True, pady=(0, 12))

        # ── Action Bar
        self._build_action_bar(body)

    def _card(self, parent, title, builder):
        frame = tk.Frame(parent, bg=BG2, bd=0, relief="flat",
                         highlightthickness=1, highlightbackground=BORDER)
        header = tk.Frame(frame, bg=BG3, padx=14, pady=9)
        header.pack(fill="x")
        tk.Label(header, text=title, font=FONT_H2, bg=BG3, fg=TEXT).pack(side="left")
        body = tk.Frame(frame, bg=BG2, padx=14, pady=12)
        body.pack(fill="both", expand=True)
        builder(body)
        return frame

    def _build_file_section(self, parent):
        # Input file
        row = tk.Frame(parent, bg=BG2)
        row.pack(fill="x", pady=(0, 8))
        tk.Label(row, text="Excel File (.xlsx)", font=FONT_MAIN, bg=BG2,
                 fg=TEXT_DIM, width=18, anchor="w").pack(side="left")
        entry = tk.Entry(row, textvariable=self.input_file, font=FONT_MONO,
                         bg=BG3, fg=TEXT, insertbackground=TEXT,
                         relief="flat", bd=6, highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=ACCENT)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn(row, "Browse", self._browse_input, small=True).pack(side="left")

        # Output dir
        row2 = tk.Frame(parent, bg=BG2)
        row2.pack(fill="x")
        tk.Label(row2, text="Output Folder", font=FONT_MAIN, bg=BG2,
                 fg=TEXT_DIM, width=18, anchor="w").pack(side="left")
        tk.Entry(row2, textvariable=self.output_dir, font=FONT_MONO,
                 bg=BG3, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=6, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT).pack(
                     side="left", fill="x", expand=True, padx=(0, 8))
        self._btn(row2, "Browse", self._browse_output, small=True).pack(side="left")

    def _build_settings_section(self, parent):
        grid = tk.Frame(parent, bg=BG2)
        grid.pack(fill="x")

        fields = [
            ("Table Name",   self.table_name,  None),
            ("Rows / File",  self.chunk_size,  None),
            ("Default Role", self.role_val,    ["voter", "admin"]),
            ("Default Status", self.status_val, ["active", "inactive"]),
        ]

        for col, (label, var, choices) in enumerate(fields):
            f = tk.Frame(grid, bg=BG2)
            f.grid(row=0, column=col, padx=(0, 16), sticky="ew")
            grid.columnconfigure(col, weight=1)
            tk.Label(f, text=label, font=("Segoe UI", 9), bg=BG2,
                     fg=TEXT_DIM).pack(anchor="w", pady=(0, 3))
            if choices:
                cb = ttk.Combobox(f, textvariable=var, values=choices,
                                  state="readonly", font=FONT_MAIN)
                cb.pack(fill="x")
                self._style_combobox(cb)
            else:
                tk.Entry(f, textvariable=var, font=FONT_MONO, bg=BG3, fg=TEXT,
                         insertbackground=TEXT, relief="flat", bd=6,
                         highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=ACCENT).pack(fill="x")

    def _build_preview_section(self, parent):
        self.preview_label = tk.Label(parent, text="Load a file to see a preview.",
                                      font=FONT_MAIN, bg=BG2, fg=TEXT_DIM)
        self.preview_label.pack(anchor="w", pady=(0, 6))

        frame = tk.Frame(parent, bg=BG3, highlightthickness=1, highlightbackground=BORDER)
        frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(frame, font=FONT_MONO, bg=BG3, fg=TEXT,
                                relief="flat", bd=8, wrap="word",
                                insertbackground=TEXT, state="disabled",
                                height=7)
        scroll = tk.Scrollbar(frame, command=self.log_text.yview, bg=BG3,
                              troughcolor=BG3, relief="flat")
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_config("green",   foreground=SUCCESS)
        self.log_text.tag_config("yellow",  foreground=WARNING)
        self.log_text.tag_config("red",     foreground=DANGER)
        self.log_text.tag_config("blue",    foreground=ACCENT)
        self.log_text.tag_config("dim",     foreground=TEXT_DIM)

    def _build_action_bar(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", pady=(0, 8))

        # Progress
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=BG3, background=ACCENT,
                        darkcolor=ACCENT, lightcolor=ACCENT,
                        bordercolor=BG3, thickness=6)
        self.pbar = ttk.Progressbar(bar, variable=self.progress,
                                    maximum=100, length=400,
                                    style="Custom.Horizontal.TProgressbar")
        self.pbar.pack(side="left", fill="x", expand=True, padx=(0, 16))

        self.status_lbl = tk.Label(bar, text="Ready", font=("Segoe UI", 9),
                                   bg=BG, fg=TEXT_DIM)
        self.status_lbl.pack(side="left", padx=(0, 16))

        self.convert_btn = self._btn(bar, "⚡  Convert to SQL", self._start_convert,
                                     accent=True)
        self.convert_btn.pack(side="right")
        self._btn(bar, "🔍 Preview", self._load_preview, small=True).pack(side="right", padx=(0, 8))

    # ── HELPERS ─────────────────────────────────────────────────

    def _btn(self, parent, text, cmd, accent=False, small=False):
        bg     = ACCENT if accent else BG3
        fg     = "#fff"  if accent else TEXT
        pad_x  = 10 if small else 18
        pad_y  = 5  if small else 9
        font   = ("Segoe UI", 9) if small else ("Segoe UI", 10, "bold")
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                      activebackground=ACCENT2 if accent else BORDER,
                      activeforeground="#fff", relief="flat", cursor="hand2",
                      padx=pad_x, pady=pad_y, font=font, bd=0)
        b.bind("<Enter>", lambda e: b.configure(bg=ACCENT2 if accent else BORDER))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _style_combobox(self, cb):
        s = ttk.Style()
        s.configure("TCombobox", fieldbackground=BG3, background=BG3,
                    foreground=TEXT, selectbackground=ACCENT,
                    selectforeground="#fff", relief="flat")
        cb.configure(style="TCombobox")

    def _log(self, msg, tag=""):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_status(self, txt, color=TEXT_DIM):
        self.status_lbl.configure(text=txt, fg=color)

    # ── FILE BROWSE ─────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if path:
            self.input_file.set(path)
            self._load_preview()

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir.set(path)

    # ── PREVIEW ─────────────────────────────────────────────────

    def _load_preview(self):
        path = self.input_file.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("No File", "Please select a valid Excel file first.")
            return
        self._clear_log()
        try:
            df = pd.read_excel(path, dtype=str, nrows=5)
            total_df = pd.read_excel(path, dtype=str)
            total = len(total_df)
            chunks = math.ceil(total / self.chunk_size.get())

            self.preview_label.configure(
                text=f"✅  {total:,} rows · {len(df.columns)} columns · "
                     f"→ {chunks} SQL files @ {self.chunk_size.get():,} rows each",
                fg=SUCCESS
            )
            self._log(f"Columns detected:", "blue")
            for col in df.columns:
                self._log(f"  • {col}", "dim")
            self._log("\nSample rows (first 5):", "blue")
            self._log(df.to_string(index=False), "")
        except Exception as e:
            self._log(f"Error reading file: {e}", "red")

    # ── CONVERT ─────────────────────────────────────────────────

    def _start_convert(self):
        if self._running:
            return
        path = self.input_file.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("No File", "Please select a valid Excel file first.")
            return
        self._running = True
        self.convert_btn.configure(state="disabled", text="Converting…")
        self._clear_log()
        self.progress.set(0)
        threading.Thread(target=self._run_convert, daemon=True).start()

    def _run_convert(self):
        try:
            start = time.time()
            self._log("📥  Reading Excel file…", "blue")
            df = pd.read_excel(self.input_file.get(), dtype={'student_number': str})
            total      = len(df)
            chunk_size = self.chunk_size.get()
            num_files  = math.ceil(total / chunk_size)
            out_dir    = self.output_dir.get()
            table      = self.table_name.get()
            role       = self.role_val.get()
            status     = self.status_val.get()

            os.makedirs(out_dir, exist_ok=True)
            self._log(f"✅  {total:,} rows loaded — generating {num_files} files\n", "green")

            def esc(val):
                if val is None or (isinstance(val, float) and math.isnan(val)) or str(val).strip() == '':
                    return 'NULL'
                return "'" + str(val).replace("'", "''") + "'"

            for i in range(num_files):
                chunk    = df.iloc[i * chunk_size: (i + 1) * chunk_size]
                fname    = f"voters_part_{i+1:02d}_of_{num_files:02d}.sql"
                filepath = os.path.join(out_dir, fname)

                with open(filepath, 'w', encoding='utf-8') as f:
                    s, e = i * chunk_size + 1, min((i + 1) * chunk_size, total)
                    f.write(f"-- Part {i+1} of {num_files} | Rows {s}–{e} of {total}\n")
                    f.write(
                        f"INSERT INTO `{table}` "
                        f"(`student_number`, `last_name`, `first_name`, `middle_name`, "
                        f"`email`, `sex`, `college_id`, `course`, `year_level`, `role`, `status`) VALUES\n"
                    )
                    rows = []
                    for _, row in chunk.iterrows():
                        yr         = str(int(row['year level'])) if pd.notna(row['year level']) else None
                        stud_num   = str(row['student_number']).strip()
                        email      = f"{stud_num}{EMAIL_DOMAIN}"
                        acr        = str(row['college']).strip().upper()
                        college_id = COLLEGE_MAP.get(acr, 'NULL')
                        vals = (
                            esc(stud_num),
                            esc(row['last name']),
                            esc(row['first name']),
                            esc(row['middle name']),
                            esc(email),
                            esc(row['sex']),
                            str(college_id),
                            esc(row['course']),
                            esc(yr),
                            f"'{role}'",
                            f"'{status}'",
                        )
                        rows.append(f"  ({', '.join(vals)})")
                    f.write(',\n'.join(rows) + ';\n')

                pct = ((i + 1) / num_files) * 100
                self.progress.set(pct)
                self._set_status(f"Part {i+1}/{num_files}", ACCENT)
                tag = "green" if (i + 1) == num_files else ""
                self._log(f"  [{i+1:02d}/{num_files}]  {fname}  ({len(chunk):,} rows)", tag)

            elapsed = time.time() - start
            self._log(f"\n🎉  Done! {num_files} files · {total:,} rows · {elapsed:.1f}s", "green")
            self._log(f"📁  Saved to: {out_dir}", "blue")
            self._set_status(f"Done — {num_files} files", SUCCESS)
            self.progress.set(100)
            messagebox.showinfo("✅ Complete",
                                f"Converted {total:,} rows into {num_files} SQL files.\n\nSaved to:\n{out_dir}")
        except Exception as ex:
            self._log(f"\n❌  Error: {ex}", "red")
            self._set_status("Error", DANGER)
            messagebox.showerror("Error", str(ex))
        finally:
            self._running = False
            self.convert_btn.configure(state="normal", text="⚡  Convert to SQL")


if __name__ == "__main__":
    app = SQLConverterApp()
    app.mainloop()