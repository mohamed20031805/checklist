import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
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
SUCCESS      = "#28A745"
WARNING      = "#FFC107"
DANGER       = "#DC3545"
SECTION_COLORS = {
    "Listed Securities":  "#2E6EE1",
    "Fund of Funds":      "#7B3FE4",
    "FOREX":              "#E67E22",
    "Listed Derivatives": "#16A085",
    "OTC":                "#C0392B",
    "Collateral":         "#2980B9",
    "Securities Lending": "#8E44AD",
}

# ─── Data model ──────────────────────────────────────────────────
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
                "has_import": False,
                "subtasks": [
                    {
                        "name": "Swifts",
                        "fields": [
                            {"label": "Sent instructor BIC code", "type": "entry", "placeholder": "XXXXXXXXX"},
                            {"label": "Emergency process?", "type": "combo", "values": ["Yes", "No"]},
                            {"label": "Forex standing instruction for Settlement + Corporate", "type": "entry", "placeholder": "Details..."}
                        ]
                    },
                    {
                        "name": "SG Market",
                        "fields": [
                            {"label": "Commentary", "type": "text", "placeholder": "Enter commentary..."}
                        ]
                    }
                ]
            },
            {"name": "French Fund",     "has_import": False, "subtasks": []},
            {"name": "Other Fund",      "has_import": False, "subtasks": []},
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {
                        "name": "Swifts",
                        "fields": [{"label": "BIC code", "type": "entry", "placeholder": "XXXXXXXXX"}]
                    },
                    {
                        "name": "SG Market",
                        "fields": [{"label": "Commentary", "type": "text", "placeholder": "Enter commentary..."}]
                    }
                ]
            },
            {
                "name": "Type of Forex",
                "has_import": False,
                "subtasks": [
                    {"name": "Spot",               "fields": []},
                    {"name": "Forward",            "fields": []},
                    {"name": "Share Class Hedging","fields": [
                        {"label": "Sent instructor BIC code", "type": "entry", "placeholder": "XXXXXXXXX"},
                        {"label": "Access to SG Markets Forex tool", "type": "combo", "values": ["Yes", "No"]},
                        {"label": "ISDA agreement to be sent",        "type": "combo", "values": ["Yes", "No"]}
                    ]},
                    {"name": "SGABLULL",            "fields": []},
                    {"name": "LMA",                 "fields": []},
                    {"name": "SGCIB",               "fields": []},
                    {"name": "External counterparties","fields": []}
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
                    {"name": "Share Class Hedging", "fields": [
                        {"label": "Sent instructor BIC code", "type": "entry", "placeholder": "XXXXXXXXX"},
                        {"label": "Access to SG Markets Forex tool", "type": "combo", "values": ["Yes", "No"]},
                        {"label": "ISDA agreement to be sent",        "type": "combo", "values": ["Yes", "No"]}
                    ]},
                    {"name": "SGABLULL",             "fields": []},
                    {"name": "LMA",                  "fields": []},
                    {"name": "SGCIB",                "fields": []},
                    {"name": "External counterparties", "fields": []}
                ]
            }
        ]
    },
    "Listed Derivatives": {
        "tasks": [
            {
                "name": "Set up with SG Prime",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Account opening confirmation with SG Prime",
                "has_import": False,
                "subtasks": [
                    {"name": "Account opening confirmation with SG Prime", "fields": []},
                    {"name": "Margin call payment process",                "fields": [
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
            {
                "name": "Set up with external broker / clearer",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Clearer Name confirmation",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Account opening confirmation with the clearer",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Margin call payment process",
                "has_import": False,
                "subtasks": [
                    {"name": "Swifts",    "fields": []},
                    {"name": "SG Market", "fields": [
                        {"label": "Account number", "type": "entry", "placeholder": "XXXXXXXX"},
                        {"label": "Creation request", "type": "text", "placeholder": "Details..."}
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
                    {"name": "Which service to be offered by SGSS", "fields": []},
                    {"name": "Booking (record keeping only)",        "fields": []},
                    {"name": "Cash payment",                        "fields": []},
                    {"name": "Valuation",                           "fields": []}
                ]
            },
            {
                "name": "Instruction method for record keeping & cash payment",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Inform client of the OTC process",
                "has_import": False,
                "subtasks": []
            },
            {
                "name": "Account set up in Simcorp",
                "has_import": False,
                "subtasks": []
            }
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
            {"name": "CSA agreement",         "has_import": False, "subtasks": []},
            {"name": "Lieu du collatéral",    "has_import": False, "subtasks": []},
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
            {"name": "Borrower identification",     "has_import": False, "subtasks": []},
            {"name": "Collateral type",             "has_import": False, "subtasks": []},
            {
                "name": "Instruction method",
                "has_import": False,
                "subtasks": [
                    {"name": "Swifts",    "fields": [
                        {"label": "BIC code", "type": "entry", "placeholder": "XXXXXXXXX"}
                    ]},
                    {"name": "SG Markets","fields": [
                        {"label": "Commentary", "type": "text", "placeholder": "Enter commentary..."}
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

        self.section_vars   = {}   # section -> BooleanVar (Yes/No)
        self.task_vars      = {}   # (section, task) -> BooleanVar
        self.subtask_vars   = {}   # (section, task, subtask) -> BooleanVar
        self.field_vars     = {}   # (section, task, subtask, field_label) -> StringVar
        self.import_paths   = {}   # (section, task) -> StringVar (file path)

        self._build_ui()

    # ── UI build ────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        topbar = tk.Frame(self, bg=BG_HEADER, height=60)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)
        tk.Label(topbar, text="SGSS  |  Onboarding Checklist",
                 font=("Arial", 16, "bold"), bg=BG_HEADER, fg=TEXT_LIGHT
                 ).pack(side="left", padx=24, pady=12)
        tk.Label(topbar, text="Confidential – Internal Use Only",
                 font=("Arial", 9), bg=BG_HEADER, fg="#8899BB"
                 ).pack(side="right", padx=24)

        # Main container
        main = tk.Frame(self, bg=BG_MAIN)
        main.pack(fill="both", expand=True)

        # Sidebar
        sidebar = tk.Frame(main, bg=BG_SIDEBAR, width=220)
        sidebar.pack(fill="y", side="left")
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        # Content area
        content = tk.Frame(main, bg=BG_MAIN)
        content.pack(fill="both", expand=True)

        # Client info bar
        info_bar = tk.Frame(content, bg=BG_CARD, pady=12)
        info_bar.pack(fill="x", padx=16, pady=(16, 0))
        self._build_info_bar(info_bar)

        # Scrollable form
        wrapper = tk.Frame(content, bg=BG_MAIN)
        wrapper.pack(fill="both", expand=True, padx=16, pady=12)

        canvas = tk.Canvas(wrapper, bg=BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        self.form_frame = tk.Frame(canvas, bg=BG_MAIN)

        self.form_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_form()

        # Bottom bar
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
                           bg=BG_SIDEBAR, fg=TEXT_LIGHT, cursor="hand2",
                           anchor="w", pady=6)
            btn.pack(fill="x", padx=8)
            btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(fg=c))
            btn.bind("<Leave>", lambda e, b=btn:  b.config(fg=TEXT_LIGHT))

    def _build_info_bar(self, parent):
        tk.Label(parent, text="Client Information", font=("Arial", 11, "bold"),
                 bg=BG_CARD, fg=TEXT_DARK).grid(row=0, column=0, columnspan=6,
                 sticky="w", padx=16, pady=(0, 8))
        fields = [
            ("Client Name",      self.client_name, 30),
            ("Reference / ID",   self.client_ref,  20),
            ("Date",             self.client_date,  14),
        ]
        for i, (lbl, var, w) in enumerate(fields):
            tk.Label(parent, text=lbl, font=("Arial", 9), bg=BG_CARD,
                     fg=TEXT_MID).grid(row=1, column=i*2, sticky="w", padx=(16,4))
            e = tk.Entry(parent, textvariable=var, width=w, font=("Arial", 10),
                         relief="solid", bd=1)
            e.grid(row=1, column=i*2+1, sticky="w", padx=(0,16))

    def _build_form(self):
        for sec_name, sec_data in SECTIONS.items():
            self._build_section(self.form_frame, sec_name, sec_data)

    def _build_section(self, parent, sec_name, sec_data):
        color = SECTION_COLORS.get(sec_name, ACCENT)

        # Section header card
        hdr = tk.Frame(parent, bg=color, pady=0)
        hdr.pack(fill="x", pady=(12, 0))
        tk.Label(hdr, text=f"  {sec_name.upper()}", font=("Arial", 11, "bold"),
                 bg=color, fg="white", pady=10, anchor="w"
                 ).pack(side="left", fill="x", expand=True)

        # Expected Yes/No
        var = tk.StringVar(value="No")
        self.section_vars[sec_name] = var
        frm = tk.Frame(hdr, bg=color)
        frm.pack(side="right", padx=12)
        tk.Label(frm, text="Expected:", font=("Arial", 9), bg=color,
                 fg="white").pack(side="left")
        for v in ("Yes", "No"):
            rb = tk.Radiobutton(frm, text=v, variable=var, value=v,
                                font=("Arial", 9, "bold"), bg=color, fg="white",
                                selectcolor=color, activebackground=color,
                                command=lambda s=sec_name: self._toggle_section(s))
            rb.pack(side="left", padx=4)

        # Tasks container (hidden by default)
        tasks_frame = tk.Frame(parent, bg=BG_CARD, relief="flat",
                               highlightbackground=BORDER_CLR, highlightthickness=1)
        tasks_frame.pack(fill="x", pady=(0, 4))
        tasks_frame.pack_forget()

        self._section_frames = getattr(self, "_section_frames", {})
        self._section_frames[sec_name] = tasks_frame

        for task in sec_data["tasks"]:
            self._build_task(tasks_frame, sec_name, task, color)

    def _toggle_section(self, sec_name):
        val = self.section_vars[sec_name].get()
        frame = self._section_frames[sec_name]
        if val == "Yes":
            frame.pack(fill="x", pady=(0, 4))
        else:
            frame.pack_forget()

    def _build_task(self, parent, sec_name, task, color):
        task_name = task["name"]
        key = (sec_name, task_name)

        # Task row
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=12, pady=4)

        var = tk.BooleanVar()
        self.task_vars[key] = var

        cb = tk.Checkbutton(row, variable=var, bg=BG_CARD,
                            activebackground=BG_CARD,
                            command=lambda k=key: self._toggle_task(k))
        cb.pack(side="left")

        tk.Label(row, text=task_name, font=("Arial", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_DARK).pack(side="left", padx=(4, 12))

        # File import button
        if task.get("has_import"):
            path_var = tk.StringVar(value="")
            self.import_paths[key] = path_var
            btn = tk.Button(row, text="📎 Import document", font=("Arial", 9),
                            bg=ACCENT_LIGHT, fg=ACCENT, relief="flat", cursor="hand2",
                            padx=8, pady=2,
                            command=lambda pv=path_var: self._pick_file(pv))
            btn.pack(side="left")
            lbl = tk.Label(row, textvariable=path_var, font=("Arial", 8),
                           bg=BG_CARD, fg=TEXT_MID, wraplength=300)
            lbl.pack(side="left", padx=6)

        # Subtasks container
        if task.get("subtasks"):
            sub_frame = tk.Frame(parent, bg="#F7F9FC")
            sub_frame.pack(fill="x", padx=24, pady=(0, 4))
            sub_frame.pack_forget()

            self._task_sub_frames = getattr(self, "_task_sub_frames", {})
            self._task_sub_frames[key] = sub_frame

            for sub in task["subtasks"]:
                self._build_subtask(sub_frame, sec_name, task_name, sub, color)

        # Separator
        tk.Frame(parent, bg=BORDER_CLR, height=1).pack(fill="x", padx=12)

    def _toggle_task(self, key):
        frames = getattr(self, "_task_sub_frames", {})
        if key in frames:
            if self.task_vars[key].get():
                frames[key].pack(fill="x", padx=24, pady=(0, 4))
            else:
                frames[key].pack_forget()

    def _build_subtask(self, parent, sec_name, task_name, sub, color):
        sub_name = sub["name"]
        key = (sec_name, task_name, sub_name)

        row = tk.Frame(parent, bg="#F7F9FC")
        row.pack(fill="x", padx=8, pady=3)

        var = tk.BooleanVar()
        self.subtask_vars[key] = var

        cb = tk.Checkbutton(row, variable=var, bg="#F7F9FC",
                            activebackground="#F7F9FC",
                            command=lambda k=key: self._toggle_subtask(k))
        cb.pack(side="left")

        dot = tk.Label(row, text="●", font=("Arial", 8), fg=color, bg="#F7F9FC")
        dot.pack(side="left")
        tk.Label(row, text=f"  {sub_name}", font=("Arial", 9, "bold"),
                 bg="#F7F9FC", fg=TEXT_DARK).pack(side="left", padx=4)

        # Fields container
        if sub.get("fields"):
            fields_frame = tk.Frame(parent, bg="#EEF2FA")
            fields_frame.pack(fill="x", padx=24, pady=(0, 4))
            fields_frame.pack_forget()

            self._subtask_field_frames = getattr(self, "_subtask_field_frames", {})
            self._subtask_field_frames[key] = fields_frame

            for fld in sub["fields"]:
                self._build_field(fields_frame, sec_name, task_name, sub_name, fld)

    def _toggle_subtask(self, key):
        frames = getattr(self, "_subtask_field_frames", {})
        if key in frames:
            if self.subtask_vars[key].get():
                frames[key].pack(fill="x", padx=24, pady=(0, 4))
            else:
                frames[key].pack_forget()

    def _build_field(self, parent, sec_name, task_name, sub_name, fld):
        label = fld["label"]
        ftype = fld.get("type", "entry")
        key   = (sec_name, task_name, sub_name, label)

        row = tk.Frame(parent, bg="#EEF2FA", pady=4)
        row.pack(fill="x", padx=8)

        tk.Label(row, text=label, font=("Arial", 9), bg="#EEF2FA",
                 fg=TEXT_MID, width=34, anchor="w").pack(side="left")

        if ftype == "entry":
            var = tk.StringVar()
            self.field_vars[key] = var
            e = tk.Entry(row, textvariable=var, width=28, font=("Arial", 9),
                         relief="solid", bd=1,
                         fg=TEXT_DARK)
            e.insert(0, fld.get("placeholder", ""))
            e.config(fg="#AAAAAA")
            e.bind("<FocusIn>",  lambda ev, w=e, ph=fld.get("placeholder",""):
                                     (w.delete(0,"end"), w.config(fg=TEXT_DARK))
                                     if w.get()==ph else None)
            e.bind("<FocusOut>", lambda ev, w=e, ph=fld.get("placeholder",""), v=var:
                                     (w.insert(0,ph), w.config(fg="#AAAAAA"))
                                     if not v.get() else None)
            e.pack(side="left")

        elif ftype == "combo":
            var = tk.StringVar()
            self.field_vars[key] = var
            cb = ttk.Combobox(row, textvariable=var, values=fld.get("values", []),
                              width=16, font=("Arial", 9), state="readonly")
            cb.pack(side="left")

        elif ftype == "text":
            var = tk.StringVar()
            self.field_vars[key] = var
            txt = tk.Text(row, width=34, height=3, font=("Arial", 9),
                          relief="solid", bd=1, fg=TEXT_DARK)
            txt.insert("1.0", fld.get("placeholder", ""))
            txt.config(fg="#AAAAAA")
            txt.bind("<FocusIn>",  lambda ev, w=txt, ph=fld.get("placeholder",""):
                                       (w.delete("1.0","end"), w.config(fg=TEXT_DARK))
                                       if w.get("1.0","end-1c")==ph else None)
            txt.bind("<FocusOut>", lambda ev, w=txt, ph=fld.get("placeholder",""), v=var:
                                       self._sync_text(w, v, ph))
            txt.pack(side="left")
            # store widget ref for export
            self.field_vars[key] = txt

    def _sync_text(self, widget, var, placeholder):
        val = widget.get("1.0", "end-1c")
        if not val or val == placeholder:
            widget.delete("1.0", "end")
            widget.insert("1.0", placeholder)
            widget.config(fg="#AAAAAA")
        # no StringVar sync needed; we'll read Text widget directly on export

    def _pick_file(self, path_var):
        path = filedialog.askopenfilename(
            title="Select document",
            filetypes=[("All files","*.*"),("PDF","*.pdf"),
                       ("Excel","*.xlsx *.xls"),("Word","*.docx")]
        )
        if path:
            path_var.set(os.path.basename(path))

    def _reset_form(self):
        if not messagebox.askyesno("Reset", "Reset all fields?"):
            return
        self.client_name.set("")
        self.client_ref.set("")
        self.client_date.set(datetime.today().strftime("%d/%m/%Y"))
        for v in self.section_vars.values():
            v.set("No")
        for k, v in self.task_vars.items():
            v.set(False)
        for k, v in self.subtask_vars.items():
            v.set(False)
        for sec in SECTIONS:
            self._section_frames[sec].pack_forget()

    # ── Excel Export ─────────────────────────────────────────────
    def _get_field_value(self, key):
        widget = self.field_vars.get(key)
        if widget is None:
            return ""
        if isinstance(widget, tk.Text):
            val = widget.get("1.0", "end-1c")
            placeholder = ""
            # find placeholder
            for sec, task_data in SECTIONS.items():
                for task in task_data["tasks"]:
                    for sub in task.get("subtasks", []):
                        for fld in sub.get("fields", []):
                            if (sec, task["name"], sub["name"], fld["label"]) == key:
                                placeholder = fld.get("placeholder", "")
            return "" if val == placeholder else val
        else:
            return widget.get()

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

        # Styles
        def hdr_font(sz=11, bold=True, color="FFFFFF"):
            return Font(name="Arial", size=sz, bold=bold, color=color)
        def cell_font(sz=10, bold=False, color="1A1A2E"):
            return Font(name="Arial", size=sz, bold=bold, color=color)
        def fill(hex_color):
            return PatternFill("solid", fgColor=hex_color.lstrip("#"))
        thin = Side(style="thin", color="DDE2EC")
        def border():
            return Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

        # Column widths
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 30
        ws.column_dimensions["F"].width = 30
        ws.column_dimensions["G"].width = 30

        row = 1

        # Title
        ws.merge_cells(f"A{row}:G{row}")
        ws[f"A{row}"] = "SGSS – Client Onboarding Checklist"
        ws[f"A{row}"].font   = hdr_font(14, True)
        ws[f"A{row}"].fill   = fill("#1C2340")
        ws[f"A{row}"].alignment = center
        ws.row_dimensions[row].height = 36
        row += 1

        # Client info
        info = [
            ("Client Name",  self.client_name.get()),
            ("Reference",    self.client_ref.get()),
            ("Date",         self.client_date.get()),
            ("Generated at", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ]
        for label, value in info:
            ws.merge_cells(f"A{row}:B{row}")
            ws[f"A{row}"] = label
            ws[f"A{row}"].font      = cell_font(10, True, "5A6478")
            ws[f"A{row}"].fill      = fill("#F0F2F5")
            ws[f"A{row}"].alignment = left
            ws.merge_cells(f"C{row}:G{row}")
            ws[f"C{row}"] = value
            ws[f"C{row}"].font      = cell_font(10, True)
            ws[f"C{row}"].fill      = fill("#F0F2F5")
            ws[f"C{row}"].alignment = left
            ws.row_dimensions[row].height = 18
            row += 1

        row += 1

        # Column headers
        headers = ["Section", "Task / Sub-task", "Expected", "Status", "Details / BIC / Commentary", "Document", "Notes"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row, col, h)
            c.font      = hdr_font(10)
            c.fill      = fill("#2E6EE1")
            c.alignment = center
            c.border    = border()
        ws.row_dimensions[row].height = 20
        row += 1

        # Data rows
        for sec_name, sec_data in SECTIONS.items():
            expected = self.section_vars[sec_name].get()
            sec_color = SECTION_COLORS.get(sec_name, ACCENT).lstrip("#")
            sec_light = "EBF1FC"  # light accent

            for ti, task in enumerate(sec_data["tasks"]):
                task_name = task["name"]
                task_key  = (sec_name, task_name)
                task_done = self.task_vars.get(task_key, tk.BooleanVar()).get()
                imp_path  = self.import_paths.get(task_key, tk.StringVar()).get()

                # Section label only on first task row
                sec_label = sec_name if ti == 0 else ""

                if task.get("subtasks"):
                    # Task header row
                    vals = [sec_label, task_name, expected,
                            "✔" if task_done else "", "", imp_path, ""]
                    for col, v in enumerate(vals, 1):
                        c = ws.cell(row, col, v)
                        c.font      = cell_font(10, True)
                        c.fill      = fill(sec_light)
                        c.alignment = left
                        c.border    = border()
                    ws.row_dimensions[row].height = 18
                    row += 1

                    for sub in task["subtasks"]:
                        sub_name = sub["name"]
                        sub_key  = (sec_name, task_name, sub_name)
                        sub_done = self.subtask_vars.get(sub_key, tk.BooleanVar()).get()

                        # Build field details string
                        field_parts = []
                        for fld in sub.get("fields", []):
                            fk  = (sec_name, task_name, sub_name, fld["label"])
                            val = self._get_field_value(fk)
                            if val:
                                field_parts.append(f"{fld['label']}: {val}")
                        details = "\n".join(field_parts)

                        vals = ["", f"  ▸ {sub_name}", expected,
                                "✔" if sub_done else "", details, "", ""]
                        for col, v in enumerate(vals, 1):
                            c = ws.cell(row, col, v)
                            c.font      = cell_font(9)
                            c.fill      = fill("FAFBFD")
                            c.alignment = left
                            c.border    = border()
                        ws.row_dimensions[row].height = 16 if not details else max(16, details.count("\n")*14+16)
                        row += 1
                else:
                    vals = [sec_label, task_name, expected,
                            "✔" if task_done else "", "", imp_path, ""]
                    for col, v in enumerate(vals, 1):
                        c = ws.cell(row, col, v)
                        c.font      = cell_font(10)
                        c.fill      = fill("FFFFFF")
                        c.alignment = left
                        c.border    = border()
                    ws.row_dimensions[row].height = 18
                    row += 1

            # Section separator
            ws.merge_cells(f"A{row}:G{row}")
            ws[f"A{row}"].fill = fill(sec_color)
            ws.row_dimensions[row].height = 4
            row += 1

        # Freeze pane
        ws.freeze_panes = "A8"

        # Auto-filter
        ws.auto_filter.ref = f"A7:G7"

        wb.save(save_path)
        messagebox.showinfo("Export successful",
                            f"File saved:\n{save_path}")

if __name__ == "__main__":
    app = ChecklistApp()
    app.mainloop()
