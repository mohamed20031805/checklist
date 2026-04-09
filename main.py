import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
import os

# ─── Color palette ───────────────────────────────────────────────
BG_MAIN      = "#F0F2F5"
BG_SIDEBAR   = "#1C2340"
BG_CARD      = "#FFFFFF"
BG_HEADER    = "#1C2340"
ACCENT       = "#2E6EE1"
ACCENT_LIGHT = "#EBF1FC"
TEXT_DARK    = "#1A1A2E"
TEXT_MID     = "#5A6478"
TEXT_LIGHT   = "#FFFFFF"
BORDER_CLR   = "#DDE2EC"
SECTION_COLORS = {
    "Listed Securities":  "#2E6EE1",
    "Fund of Funds":      "#7B3FE4",
    "FOREX":              "#E67E22",
    "Listed Derivatives": "#16A085",
    "OTC":                "#C0392B",
    "Collateral":         "#2980B9",
    "Securities Lending": "#8E44AD",
}

SECTIONS = {
    "Listed Securities": {
        "tasks": [
            {
                "name": "Sent target portfolio",
                "has_import": True,
                "subtasks": []
            },
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {
                        "name": "Swifts",
                        "fields": [
                            {"label": "Sent instructor BIC code", "type": "entry", "placeholder": "XXXXXXXXX"}
                        ]
                    },
                    {
                        "name": "SG Market",
                        "fields": [
                            {"label": "Commentary / Instructions", "type": "text", "placeholder": "Enter SG Market instructions..."}
                        ]
                    }
                ]
            }
        ]
    },
    "Fund of Funds": {
        "tasks": [
            {
                "name": "Type of Mutual Funds",
                "has_import": True,
                "type": "dropdown_conditional",
                "dropdown_label": "Fund type",
                "dropdown_values": ["French Fund", "Other Fund"],
                "subtasks": []
            },
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {
                        "name": "Swifts",
                        "fields": [
                            {"label": "BIC code", "type": "entry", "placeholder": "XXXXXXXXX"}
                        ]
                    },
                    {
                        "name": "SG Market",
                        "fields": [
                            {"label": "Commentary", "type": "text", "placeholder": "Enter commentary..."}
                        ]
                    }
                ]
            }
        ]
    },
    "FOREX": {
        "tasks": [
            {
                "name": "FX Counterparties",
                "has_import": False,
                "subtasks": [
                    {"name": "Spot",    "fields": []},
                    {"name": "Forward", "fields": []},
                    {
                        "name": "Share Class Hedging",
                        "fields": [
                            {"label": "Sent instructor BIC code",        "type": "entry", "placeholder": "XXXXXXXXX"},
                            {"label": "Access to SG Markets Forex tool", "type": "combo", "values": ["Yes", "No"]},
                            {"label": "ISDA agreement to be sent",       "type": "combo", "values": ["Yes", "No"]}
                        ]
                    },
                    {"name": "SGABLULL",               "fields": []},
                    {"name": "LMA",                    "fields": []},
                    {"name": "SGCIB",                  "fields": []},
                    {"name": "External counterparties","fields": []}
                ]
            }
        ]
    },
    "Listed Derivatives": {
        "tasks": [
            {"name": "Set up with SG Prime",                         "has_import": False, "subtasks": []},
            {
                "name": "Account opening confirmation with SG Prime",
                "has_import": False,
                "subtasks": [
                    {"name": "Account opening confirmation with SG Prime", "fields": []},
                    {"name": "Margin call payment process", "fields": [
                        {"label": "Process type", "type": "combo",
                         "values": ["0 treasury", "Swifts", "SG Markets"]}
                    ]},
                    {"name": "0 treasury", "fields": []},
                    {"name": "Swifts",     "fields": []},
                    {"name": "SG Market",  "fields": [
                        {"label": "POA (contract for 0 treasury process)", "type": "combo", "values": ["Yes", "No"]}
                    ]}
                ]
            },
            {"name": "Set up with external broker / clearer",        "has_import": False, "subtasks": []},
            {"name": "Clearer Name confirmation",                     "has_import": False, "subtasks": []},
            {"name": "Account opening confirmation with the clearer", "has_import": False, "subtasks": []},
            {
                "name": "Margin call payment process",
                "has_import": False,
                "subtasks": [
                    {"name": "Swifts",    "fields": []},
                    {"name": "SG Market", "fields": [
                        {"label": "Account number",   "type": "entry", "placeholder": "XXXXXXXX"},
                        {"label": "Creation request", "type": "text",  "placeholder": "Details..."}
                    ]}
                ]
            }
        ]
    },
    "OTC": {
        "tasks": [
            {
                "name": "Type of OTC products",
                "has_import": False,
                "subtasks": [
                    {"name": "Booking (record keeping only)", "fields": []},
                    {"name": "Cash payment",                 "fields": []},
                    {"name": "Valuation",                    "fields": []}
                ]
            },
            {"name": "Instruction method for record keeping & cash payment", "has_import": False, "subtasks": []},
            {"name": "Inform client of the OTC process",                     "has_import": False, "subtasks": []},
            {"name": "Account set up in Simcorp",                            "has_import": False, "subtasks": []}
        ]
    },
    "Collateral": {
        "tasks": [
            {"name": "Collateral type (Securities, cash)", "has_import": False, "subtasks": []},
            {"name": "Collateral Agent",                  "has_import": False, "subtasks": []},
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {"name": "Swifts",    "fields": []},
                    {"name": "SG Markets","fields": []}
                ]
            },
            {"name": "CSA agreement",      "has_import": False, "subtasks": []},
            {"name": "Lieu du collatéral", "has_import": False, "subtasks": []}
        ]
    },
    "Securities Lending": {
        "tasks": [
            {
                "name": "Type of lending",
                "has_import": False,
                "subtasks": [
                    {"name": "Agent lending",    "fields": []},
                    {"name": "Principal lending", "fields": []}
                ]
            },
            {"name": "Borrower identification", "has_import": False, "subtasks": []},
            {"name": "Collateral type",         "has_import": False, "subtasks": []},
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {"name": "Swifts",    "fields": [
                        {"label": "BIC code",   "type": "entry", "placeholder": "XXXXXXXXX"}
                    ]},
                    {"name": "SG Markets","fields": [
                        {"label": "Commentary", "type": "text",  "placeholder": "Enter commentary..."}
                    ]}
                ]
            },
            {"name": "Revenue split agreement", "has_import": False, "subtasks": []},
            {"name": "SLA / Agreement",         "has_import": True,  "subtasks": []}
        ]
    }
}

# ─── App ─────────────────────────────────────────────────────────
class ChecklistApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SGSS Onboarding Checklist")
        self.geometry("1280x800")
        self.minsize(1100, 700)
        self.configure(bg=BG_MAIN)

        self.client_name = tk.StringVar()
        self.client_ref  = tk.StringVar()
        self.client_date = tk.StringVar(value=datetime.today().strftime("%d/%m/%Y"))

        self.section_vars          = {}
        self.task_vars             = {}
        self.subtask_vars          = {}
        self.field_vars            = {}
        self.import_paths          = {}
        self.import_full_paths     = {}
        self.dropdown_vars         = {}
        self._section_frames       = {}
        self._task_sub_frames      = {}
        self._subtask_field_frames = {}

        self._build_ui()

    def _build_ui(self):
        topbar = tk.Frame(self, bg=BG_HEADER, height=60)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)
        tk.Label(topbar, text="SGSS  |  Onboarding Checklist",
                 font=("Arial", 16, "bold"), bg=BG_HEADER, fg=TEXT_LIGHT
                 ).pack(side="left", padx=24, pady=12)
        tk.Label(topbar, text="Confidential – Internal Use Only",
                 font=("Arial", 9), bg=BG_HEADER, fg="#8899BB"
                 ).pack(side="right", padx=24)

        main = tk.Frame(self, bg=BG_MAIN)
        main.pack(fill="both", expand=True)

        sidebar = tk.Frame(main, bg=BG_SIDEBAR, width=220)
        sidebar.pack(fill="y", side="left")
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        content = tk.Frame(main, bg=BG_MAIN)
        content.pack(fill="both", expand=True)

        info_bar = tk.Frame(content, bg=BG_CARD, pady=12)
        info_bar.pack(fill="x", padx=16, pady=(16, 0))
        self._build_info_bar(info_bar)

        wrapper = tk.Frame(content, bg=BG_MAIN)
        wrapper.pack(fill="both", expand=True, padx=16, pady=12)

        self._canvas = tk.Canvas(wrapper, bg=BG_MAIN, highlightthickness=0)
        sb = ttk.Scrollbar(wrapper, orient="vertical", command=self._canvas.yview)
        self.form_frame = tk.Frame(self._canvas, bg=BG_MAIN)
        self.form_frame.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_form()

        bottom = tk.Frame(content, bg=BG_CARD, pady=10)
        bottom.pack(fill="x", padx=16, pady=(0, 16))
        tk.Button(bottom, text="⬇  Export to Excel", font=("Arial", 11, "bold"),
                  bg=ACCENT, fg="white", relief="flat", cursor="hand2",
                  padx=24, pady=8, command=self._export_excel
                  ).pack(side="right", padx=16)
        tk.Button(bottom, text="↺  Reset Form", font=("Arial", 10),
                  bg=BG_MAIN, fg=TEXT_MID, relief="flat", cursor="hand2",
                  padx=16, pady=8, command=self._reset_form
                  ).pack(side="right")

    def _build_sidebar(self, parent):
        tk.Label(parent, text="SECTIONS", font=("Arial", 9, "bold"),
                 bg=BG_SIDEBAR, fg="#6677AA"
                 ).pack(anchor="w", padx=16, pady=(20, 8))
        for sec in SECTIONS:
            color = SECTION_COLORS.get(sec, ACCENT)
            btn = tk.Label(parent, text=f"  ● {sec}", font=("Arial", 10),
                           bg=BG_SIDEBAR, fg=TEXT_LIGHT, cursor="hand2", anchor="w", pady=6)
            btn.pack(fill="x", padx=8)
            btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(fg=c))
            btn.bind("<Leave>", lambda e, b=btn: b.config(fg=TEXT_LIGHT))

    def _build_info_bar(self, parent):
        tk.Label(parent, text="Client Information", font=("Arial", 11, "bold"),
                 bg=BG_CARD, fg=TEXT_DARK).grid(row=0, column=0, columnspan=6,
                 sticky="w", padx=16, pady=(0, 8))
        for i, (lbl, var, w) in enumerate([
            ("Client Name",    self.client_name, 30),
            ("Reference / ID", self.client_ref,  20),
            ("Date",           self.client_date, 14),
        ]):
            tk.Label(parent, text=lbl, font=("Arial", 9), bg=BG_CARD,
                     fg=TEXT_MID).grid(row=1, column=i*2, sticky="w", padx=(16, 4))
            tk.Entry(parent, textvariable=var, width=w, font=("Arial", 10),
                     relief="solid", bd=1).grid(row=1, column=i*2+1, sticky="w", padx=(0, 16))

    def _build_form(self):
        for sec_name, sec_data in SECTIONS.items():
            self._build_section(self.form_frame, sec_name, sec_data)

    def _build_section(self, parent, sec_name, sec_data):
        color = SECTION_COLORS.get(sec_name, ACCENT)
        hdr   = tk.Frame(parent, bg=color)
        hdr.pack(fill="x", pady=(12, 0))
        tk.Label(hdr, text=f"  {sec_name.upper()}", font=("Arial", 11, "bold"),
                 bg=color, fg="white", pady=10, anchor="w"
                 ).pack(side="left", fill="x", expand=True)

        var = tk.StringVar(value="No")
        self.section_vars[sec_name] = var
        frm = tk.Frame(hdr, bg=color)
        frm.pack(side="right", padx=12)
        tk.Label(frm, text="Expected:", font=("Arial", 9), bg=color, fg="white").pack(side="left")
        for v in ("Yes", "No"):
            tk.Radiobutton(frm, text=v, variable=var, value=v,
                           font=("Arial", 9, "bold"), bg=color, fg="white",
                           selectcolor=color, activebackground=color,
                           command=lambda s=sec_name: self._toggle_section(s)
                           ).pack(side="left", padx=4)

        tasks_frame = tk.Frame(parent, bg=BG_CARD,
                               highlightbackground=BORDER_CLR, highlightthickness=1)
        tasks_frame.pack(fill="x", pady=(0, 4))
        tasks_frame.pack_forget()
        self._section_frames[sec_name] = tasks_frame

        for task in sec_data["tasks"]:
            self._build_task(tasks_frame, sec_name, task, color)

    def _toggle_section(self, sec_name):
        if self.section_vars[sec_name].get() == "Yes":
            self._section_frames[sec_name].pack(fill="x", pady=(0, 4))
        else:
            self._section_frames[sec_name].pack_forget()

    def _build_task(self, parent, sec_name, task, color):
        task_name   = task["name"]
        key         = (sec_name, task_name)
        is_dropdown = task.get("type") == "dropdown_conditional"

        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=12, pady=4)

        var = tk.BooleanVar()
        self.task_vars[key] = var
        tk.Checkbutton(row, variable=var, bg=BG_CARD, activebackground=BG_CARD,
                       command=lambda k=key: self._toggle_task(k)).pack(side="left")
        tk.Label(row, text=task_name, font=("Arial", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_DARK).pack(side="left", padx=(4, 12))

        if is_dropdown:
            dd_var = tk.StringVar(value="")
            self.dropdown_vars[key] = dd_var
            tk.Label(row, text=task.get("dropdown_label", "Type") + ":",
                     font=("Arial", 9), bg=BG_CARD, fg=TEXT_MID).pack(side="left")
            ttk.Combobox(row, textvariable=dd_var,
                         values=task.get("dropdown_values", []),
                         width=16, font=("Arial", 9), state="readonly"
                         ).pack(side="left", padx=(4, 12))

        if task.get("has_import"):
            pv = tk.StringVar(value="")
            fp = tk.StringVar(value="")
            self.import_paths[key]      = pv
            self.import_full_paths[key] = fp
            tk.Button(row, text="📎 Import document", font=("Arial", 9),
                      bg=ACCENT_LIGHT, fg=ACCENT, relief="flat", cursor="hand2",
                      padx=8, pady=2,
                      command=lambda p=pv, f=fp: self._pick_file(p, f)
                      ).pack(side="left")
            tk.Label(row, textvariable=pv, font=("Arial", 8),
                     bg=BG_CARD, fg=TEXT_MID, wraplength=300).pack(side="left", padx=6)

        if task.get("subtasks"):
            sf = tk.Frame(parent, bg="#F7F9FC")
            sf.pack(fill="x", padx=24, pady=(0, 4))
            sf.pack_forget()
            self._task_sub_frames[key] = sf
            for sub in task["subtasks"]:
                self._build_subtask(sf, sec_name, task_name, sub, color)

        tk.Frame(parent, bg=BORDER_CLR, height=1).pack(fill="x", padx=12)

    def _toggle_task(self, key):
        if key in self._task_sub_frames:
            if self.task_vars[key].get():
                self._task_sub_frames[key].pack(fill="x", padx=24, pady=(0, 4))
            else:
                self._task_sub_frames[key].pack_forget()

    def _build_subtask(self, parent, sec_name, task_name, sub, color):
        sub_name = sub["name"]
        key      = (sec_name, task_name, sub_name)
        row      = tk.Frame(parent, bg="#F7F9FC")
        row.pack(fill="x", padx=8, pady=3)

        var = tk.BooleanVar()
        self.subtask_vars[key] = var
        tk.Checkbutton(row, variable=var, bg="#F7F9FC", activebackground="#F7F9FC",
                       command=lambda k=key: self._toggle_subtask(k)).pack(side="left")
        tk.Label(row, text="●", font=("Arial", 8), fg=color, bg="#F7F9FC").pack(side="left")
        tk.Label(row, text=f"  {sub_name}", font=("Arial", 9, "bold"),
                 bg="#F7F9FC", fg=TEXT_DARK).pack(side="left", padx=4)

        if sub.get("fields"):
            ff = tk.Frame(parent, bg="#EEF2FA")
            ff.pack(fill="x", padx=24, pady=(0, 4))
            ff.pack_forget()
            self._subtask_field_frames[key] = ff
            for fld in sub["fields"]:
                self._build_field(ff, sec_name, task_name, sub_name, fld)

    def _toggle_subtask(self, key):
        if key in self._subtask_field_frames:
            if self.subtask_vars[key].get():
                self._subtask_field_frames[key].pack(fill="x", padx=24, pady=(0, 4))
            else:
                self._subtask_field_frames[key].pack_forget()

    def _build_field(self, parent, sec_name, task_name, sub_name, fld):
        label = fld["label"]
        ftype = fld.get("type", "entry")
        key   = (sec_name, task_name, sub_name, label)
        row   = tk.Frame(parent, bg="#EEF2FA", pady=4)
        row.pack(fill="x", padx=8)
        tk.Label(row, text=label, font=("Arial", 9), bg="#EEF2FA",
                 fg=TEXT_MID, width=34, anchor="w").pack(side="left")

        if ftype == "entry":
            var = tk.StringVar()
            self.field_vars[key] = var
            ph  = fld.get("placeholder", "")
            e   = tk.Entry(row, textvariable=var, width=28, font=("Arial", 9),
                           relief="solid", bd=1, fg="#AAAAAA")
            e.insert(0, ph)
            e.bind("<FocusIn>",  lambda ev, w=e, p=ph:
                   (w.delete(0, "end"), w.config(fg=TEXT_DARK)) if w.get() == p else None)
            e.bind("<FocusOut>", lambda ev, w=e, p=ph, v=var:
                   (w.insert(0, p), w.config(fg="#AAAAAA")) if not v.get() else None)
            e.pack(side="left")

        elif ftype == "combo":
            var = tk.StringVar()
            self.field_vars[key] = var
            ttk.Combobox(row, textvariable=var, values=fld.get("values", []),
                         width=16, font=("Arial", 9), state="readonly").pack(side="left")

        elif ftype == "text":
            ph  = fld.get("placeholder", "")
            txt = tk.Text(row, width=34, height=3, font=("Arial", 9),
                          relief="solid", bd=1, fg="#AAAAAA")
            txt.insert("1.0", ph)
            txt.bind("<FocusIn>",  lambda ev, w=txt, p=ph:
                     (w.delete("1.0", "end"), w.config(fg=TEXT_DARK))
                     if w.get("1.0", "end-1c") == p else None)
            txt.bind("<FocusOut>", lambda ev, w=txt, p=ph:
                     (w.delete("1.0", "end"), w.insert("1.0", p), w.config(fg="#AAAAAA"))
                     if not w.get("1.0", "end-1c") else None)
            txt.pack(side="left")
            self.field_vars[key] = txt

    def _pick_file(self, path_var, full_path_var):
        path = filedialog.askopenfilename(
            title="Select document",
            filetypes=[("All files","*.*"),("PDF","*.pdf"),
                       ("Excel","*.xlsx *.xls"),("Word","*.docx"),
                       ("Images","*.png *.jpg *.jpeg")]
        )
        if path:
            path_var.set(os.path.basename(path))
            full_path_var.set(path)

    def _reset_form(self):
        if not messagebox.askyesno("Reset", "Reset all fields?"):
            return
        self.client_name.set("")
        self.client_ref.set("")
        self.client_date.set(datetime.today().strftime("%d/%m/%Y"))
        for v in self.section_vars.values():
            v.set("No")
        for v in self.task_vars.values():
            v.set(False)
        for v in self.subtask_vars.values():
            v.set(False)
        for sec in SECTIONS:
            self._section_frames[sec].pack_forget()

    def _get_field_value(self, key):
        widget = self.field_vars.get(key)
        if widget is None:
            return ""
        if isinstance(widget, tk.Text):
            val = widget.get("1.0", "end-1c")
            for sec, sd in SECTIONS.items():
                for task in sd["tasks"]:
                    for sub in task.get("subtasks", []):
                        for fld in sub.get("fields", []):
                            if (sec, task["name"], sub["name"], fld["label"]) == key:
                                return "" if val == fld.get("placeholder", "") else val
            return val
        return widget.get()

    # ── Excel Export ─────────────────────────────────────────────
    def _export_excel(self):
        if not self.client_name.get().strip():
            messagebox.showwarning("Missing info", "Please enter the client name.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx")],
            initialfile=f"Onboarding_{self.client_name.get().replace(' ','_')}_{datetime.today().strftime('%Y%m%d')}.xlsx"
        )
        if not save_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Onboarding Checklist"

        def hf(sz=11, bold=True, color="FFFFFF"):
            return Font(name="Arial", size=sz, bold=bold, color=color)
        def cf(sz=10, bold=False, color="1A1A2E"):
            return Font(name="Arial", size=sz, bold=bold, color=color)
        def xf(h):
            return PatternFill("solid", fgColor=h.lstrip("#"))
        thin = Side(style="thin", color="DDE2EC")
        def brd():
            return Border(left=thin, right=thin, top=thin, bottom=thin)
        ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
        lft = Alignment(horizontal="left",   vertical="center", wrap_text=True)

        for col, w in zip("ABCDEFG", [22, 32, 14, 14, 32, 26, 20]):
            ws.column_dimensions[col].width = w

        r = 1
        ws.merge_cells(f"A{r}:G{r}")
        ws[f"A{r}"] = "SGSS – Client Onboarding Checklist"
        ws[f"A{r}"].font = hf(14); ws[f"A{r}"].fill = xf("#1C2340")
        ws[f"A{r}"].alignment = ctr; ws.row_dimensions[r].height = 36
        r += 1

        for lbl, val in [("Client Name", self.client_name.get()),
                         ("Reference",   self.client_ref.get()),
                         ("Date",        self.client_date.get()),
                         ("Generated at",datetime.now().strftime("%d/%m/%Y %H:%M"))]:
            ws.merge_cells(f"A{r}:B{r}"); ws[f"A{r}"] = lbl
            ws[f"A{r}"].font = cf(10, True, "5A6478"); ws[f"A{r}"].fill = xf("#F0F2F5")
            ws[f"A{r}"].alignment = lft
            ws.merge_cells(f"C{r}:G{r}"); ws[f"C{r}"] = val
            ws[f"C{r}"].font = cf(10, True); ws[f"C{r}"].fill = xf("#F0F2F5")
            ws[f"C{r}"].alignment = lft; ws.row_dimensions[r].height = 18
            r += 1

        r += 1
        for col, h in enumerate(["Section","Task / Sub-task","Expected","Status",
                                  "Details / BIC / Commentary","Fund Type / Document","Notes"], 1):
            c = ws.cell(r, col, h)
            c.font = hf(10); c.fill = xf("#2E6EE1"); c.alignment = ctr; c.border = brd()
        ws.row_dimensions[r].height = 20; r += 1

        doc_idx = 1

        for sec_name, sec_data in SECTIONS.items():
            expected  = self.section_vars[sec_name].get()
            sec_hex   = SECTION_COLORS.get(sec_name, ACCENT).lstrip("#")

            for ti, task in enumerate(sec_data["tasks"]):
                task_name = task["name"]
                tkey      = (sec_name, task_name)
                done      = self.task_vars.get(tkey, tk.BooleanVar()).get()
                imp_name  = self.import_paths.get(tkey, tk.StringVar()).get()
                imp_full  = self.import_full_paths.get(tkey, tk.StringVar()).get()
                dd_val    = self.dropdown_vars.get(tkey, tk.StringVar()).get()
                sec_lbl   = sec_name if ti == 0 else ""

                # Colonne F : type de fond OU référence document
                extra = dd_val if dd_val else ""
                doc_ref = ""
                if imp_full and os.path.exists(imp_full):
                    sname   = f"Document_{doc_idx}"
                    doc_ref = f"→ Feuille '{sname}'"
                    self._add_doc_sheet(wb, sname, imp_full, sec_name, task_name, imp_name)
                    doc_idx += 1

                col_f = doc_ref if doc_ref else extra

                if task.get("subtasks"):
                    for col, v in enumerate([sec_lbl, task_name, expected,
                                             "✔" if done else "", "", col_f, ""], 1):
                        c = ws.cell(r, col, v)
                        c.font = cf(10, True); c.fill = xf("EBF1FC")
                        c.alignment = lft; c.border = brd()
                    ws.row_dimensions[r].height = 18; r += 1

                    for sub in task["subtasks"]:
                        skey  = (sec_name, task_name, sub["name"])
                        sdone = self.subtask_vars.get(skey, tk.BooleanVar()).get()
                        parts = []
                        for fld in sub.get("fields", []):
                            fk  = (sec_name, task_name, sub["name"], fld["label"])
                            val = self._get_field_value(fk)
                            if val:
                                parts.append(f"{fld['label']}: {val}")
                        det = "\n".join(parts)
                        for col, v in enumerate(["", f"  ▸ {sub['name']}", expected,
                                                  "✔" if sdone else "", det, "", ""], 1):
                            c = ws.cell(r, col, v)
                            c.font = cf(9); c.fill = xf("FAFBFD")
                            c.alignment = lft; c.border = brd()
                        ws.row_dimensions[r].height = max(16, det.count("\n")*14+16); r += 1
                else:
                    for col, v in enumerate([sec_lbl, task_name, expected,
                                             "✔" if done else "", "", col_f, ""], 1):
                        c = ws.cell(r, col, v)
                        c.font = cf(10); c.fill = xf("FFFFFF")
                        c.alignment = lft; c.border = brd()
                    ws.row_dimensions[r].height = 18; r += 1

            ws.merge_cells(f"A{r}:G{r}")
            ws[f"A{r}"].fill = xf(sec_hex)
            ws.row_dimensions[r].height = 4; r += 1

        ws.freeze_panes = "A8"
        ws.auto_filter.ref = "A7:G7"
        wb.save(save_path)
        docs_msg = f"\n\n{doc_idx-1} document(s) intégré(s) en feuille(s) séparée(s)." if doc_idx > 1 else ""
        messagebox.showinfo("Export réussi ✔", f"Fichier sauvegardé :\n{save_path}{docs_msg}")

    def _add_doc_sheet(self, wb, sheet_name, full_path, sec_name, task_name, basename):
        ws  = wb.create_sheet(title=sheet_name)
        col = SECTION_COLORS.get(sec_name, ACCENT)

        def xf(h):  return PatternFill("solid", fgColor=h.lstrip("#"))
        def hf(**kw): return Font(name="Arial", **kw)

        ws.merge_cells("A1:F1")
        ws["A1"] = f"Document  –  {sec_name}  /  {task_name}"
        ws["A1"].font = hf(size=12, bold=True, color="FFFFFF")
        ws["A1"].fill = xf(col)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        for i, (lbl, val) in enumerate([
            ("Fichier",     basename),
            ("Section",     sec_name),
            ("Tâche",       task_name),
            ("Importé le",  datetime.now().strftime("%d/%m/%Y %H:%M")),
        ], 2):
            ws[f"A{i}"] = lbl; ws[f"A{i}"].font = hf(size=10, bold=True, color="5A6478")
            ws[f"B{i}"] = val; ws[f"B{i}"].font = hf(size=10)
            ws.row_dimensions[i].height = 18

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 50

        ext = os.path.splitext(full_path)[1].lower()

        if ext in (".png", ".jpg", ".jpeg", ".bmp"):
            try:
                img = XLImage(full_path)
                img.anchor = "A7"
                max_w, max_h = 600, 700
                if img.width > max_w:
                    ratio = max_w / img.width
                    img.width = int(img.width * ratio); img.height = int(img.height * ratio)
                if img.height > max_h:
                    ratio = max_h / img.height
                    img.width = int(img.width * ratio); img.height = int(img.height * ratio)
                ws.add_image(img)
            except Exception as ex:
                ws["A7"] = f"⚠ Image non intégrée : {ex}"
                ws["A7"].font = hf(size=9, italic=True, color="CC0000")

        elif ext in (".xlsx", ".xls"):
            try:
                src = openpyxl.load_workbook(full_path, data_only=True)
                sw  = src.active
                ws["A7"] = f"Contenu du fichier : {basename}"
                ws["A7"].font = hf(size=10, bold=True)
                ro = 8
                for src_row in sw.iter_rows(values_only=True):
                    for ci, val in enumerate(src_row, 1):
                        if val is not None:
                            ws.cell(ro, ci, str(val)).font = hf(size=9)
                    ro += 1
                    if ro > 500:
                        ws.cell(ro, 1, "… (tronqué à 500 lignes)").font = hf(size=9, italic=True, color="888888")
                        break
            except Exception as ex:
                ws["A7"] = f"⚠ Impossible de lire le fichier Excel : {ex}"
                ws["A7"].font = hf(size=9, italic=True, color="CC0000")

        else:
            ws["A7"] = "📄 Chemin du document (ouvrir manuellement) :"
            ws["A7"].font = hf(size=10, bold=True)
            ws["B7"] = full_path
            ws["B7"].font = hf(size=9, color="2E6EE1")
            ws.merge_cells("A8:F8")
            ws["A8"] = "ℹ Les fichiers PDF et Word ne peuvent pas être intégrés directement dans Excel."
            ws["A8"].font = hf(size=9, italic=True, color="888888")


if __name__ == "__main__":
    app = ChecklistApp()
    app.mainloop()



#pip install openpyxl pyinstaller
#pyinstaller --onefile --windowed --name "SGSS_Onboarding_Checklist" main.py
