"""Bilingual Tkinter form connected to the independently tested engine."""
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .engine import ROOT, generate_docx
from .output import OUTPUT_DIR, client_filename
from .preview import DEVLYS_FONT, HUMAN_FIELDS, DevLysPreview
from .i18n import error_text, translate
from .model import ClientData, LandRecord, ValidationError, load_client, save_client
from .runtime import check_tk_version

FIELDS = [
    ("client_name", "Client name"), ("father_name", "Father's name"),
    ("age", "Age (years)"), ("caste", "Caste"),
    ("address", "Address"), ("village", "Village"),
    ("tehsil", "Tehsil"), ("district", "District"),
    ("bahi", "Bahi number (one per document)"),
    ("document_date", "Document date (YYYY-MM-DD)"),
]
RECORD_FIELDS = (("khasra", "Khasra number"), ("area", "Area (hectares)"), ("lagan", "Lagan"))


class RecordDialog(tk.Toplevel):
    def __init__(self, parent, record=None, language="English", preview=None):
        super().__init__(parent)
        self.language = language
        tr = lambda message: translate(message, language)
        self.title(tr("Edit land record" if record else "Add land record"))
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        panel = ttk.Frame(self, padding=20)
        panel.grid()
        self.values = {}
        for row, (key, label) in enumerate(RECORD_FIELDS):
            ttk.Label(panel, text=tr(label)).grid(row=row, column=0, sticky="w", padx=(0, 18), pady=6)
            var = tk.StringVar(value=getattr(record, key) if record else "")
            self.values[key] = var
            entry = ttk.Entry(panel, textvariable=var, width=36)
            entry.grid(row=row, column=1, pady=6)
            if preview is not None:
                preview.watch(entry, var, label, "Times New Roman")
            if row == 0:
                entry.focus_set()
        ttk.Label(panel, text=tr("Identifiers keep leading zeros and slashes.\nUse a dot for decimals; no commas.")).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 16))
        buttons = ttk.Frame(panel)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text=tr("Cancel"), command=self.destroy).pack(side="left", padx=6)
        ttk.Button(buttons, text=tr("Apply"), command=self.apply).pack(side="left")
        self.bind("<Return>", lambda _: self.apply())
        self.bind("<Escape>", lambda _: self.destroy())
        self.grab_set()

    def apply(self):
        try:
            self.result = LandRecord(**{key: value.get() for key, value in self.values.items()})
        except ValidationError as error:
            messagebox.showerror(translate("Check land record", self.language), error_text(error, self.language), parent=self)
            return
        self.destroy()


class Application(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=18)
        self.pack(fill="both", expand=True)
        self.records = []
        self.data_path = None
        self.dirty = False
        self.loading = True
        self.language = tk.StringVar(value="English")
        self.preview = DevLysPreview(master, self.language.get)
        self.values = {key: tk.StringVar() for key, _ in FIELDS}
        self.values["document_date"].set(date.today().isoformat())
        self.preview = DevLysPreview(master, self.language.get)
        self.status = tk.StringVar()
        self.status_message = "Enter client details or load sample data."
        self.status_values = {}
        self.total_label = tk.StringVar()
        self.localized = []
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)
        self.build()
        for var in self.values.values():
            var.trace_add("write", self.changed)
        self.language.trace_add("write", self.apply_language)
        self.apply_language()
        self.loading = False
        master.protocol("WM_DELETE_WINDOW", self.close)
        for shortcut in ("<Command-s>", "<Control-s>"):
            master.bind(shortcut, lambda _: self.save())
        for shortcut in ("<Command-o>", "<Control-o>"):
            master.bind(shortcut, lambda _: self.open())

    def tr(self, message, **values):
        return translate(message, self.language.get(), **values)

    def widget(self, kind, parent, message, **options):
        widget = kind(parent, text=self.tr(message), **options)
        self.localized.append((widget, message))
        return widget

    def build(self):
        label = lambda parent, message, **options: self.widget(ttk.Label, parent, message, **options)
        button = lambda parent, message, command: self.widget(ttk.Button, parent, message, command=command)
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        title = ttk.Frame(bar)
        title.pack(fill="x", pady=(0, 10))
        label(title, "Land record documents", font=("TkDefaultFont", 19, "bold")).pack(side="left")
        self.language_combo = ttk.Combobox(title, textvariable=self.language, values=("English", "हिन्दी"), state="readonly", width=10)
        self.language_combo.pack(side="right")
        label(title, "Interface language").pack(side="right", padx=10)
        for message, command in (("New", self.new), ("Open JSON…", self.open), ("Save JSON…", self.save), ("Load sample", self.sample)):
            button(bar, message, command).pack(side="left", padx=(0, 12))

        details = self.widget(ttk.LabelFrame, self, "Client details", padding=12)
        details.grid(row=1, column=0, sticky="ew")
        details.columnconfigure(1, weight=1)
        details.columnconfigure(3, weight=1)
        self.input_entries = {}
        for i, (key, message) in enumerate(FIELDS):
            row, col = divmod(i, 2)
            label(details, message).grid(row=row, column=col*2, sticky="w", padx=(0 if col == 0 else 22, 12), pady=4)
            entry = ttk.Entry(details, textvariable=self.values[key], width=24)
            entry.grid(row=row, column=col*2+1, sticky="ew", pady=4)
            self.input_entries[key] = entry
            self.preview.watch(entry, self.values[key], message, DEVLYS_FONT if key in HUMAN_FIELDS else "Times New Roman")

        encoding = self.widget(ttk.LabelFrame, self, "Text and fonts", padding=12)
        encoding.grid(row=2, column=0, sticky="ew", pady=10)
        label(encoding, "Document input font").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Label(encoding, text=DEVLYS_FONT).grid(row=0, column=1, sticky="w")
        label(encoding, "Type DevLys keyboard codes. The shared preview follows the selected field.", wraplength=970).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        label(encoding, "Documents are saved in output/ as clientname_data.docx.", wraplength=970).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.preview.mount(encoding)

        heading = ttk.Frame(self)
        heading.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        label(heading, "Land records", font=("TkDefaultFont", 13, "bold")).pack(side="left")
        for message, command in (("Add…", self.add), ("Edit…", self.edit), ("Remove", self.remove)):
            button(heading, message, command).pack(side="left", padx=(12, 0))
        label(heading, "Double-click a row to edit").pack(side="right")
        table_panel = ttk.Frame(self)
        table_panel.grid(row=4, column=0, sticky="nsew")
        table_panel.rowconfigure(0, weight=1)
        table_panel.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_panel, columns=("khasra", "area", "lagan"), show="headings", selectmode="extended", height=6)
        for key in ("khasra", "area", "lagan"):
            self.tree.column(key, width=250, minwidth=100, anchor="e" if key in ("area", "lagan") else "w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_panel, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", lambda _: self.edit())
        for event in ("<BackSpace>", "<Delete>"):
            self.tree.bind(event, lambda _: self.remove())
        footer = ttk.Frame(self)
        footer.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(footer, textvariable=self.total_label).pack(side="left")
        button(footer, "Generate new DOCX…", self.generate).pack(side="right")
        label(self, "Uses the original template in templates/. Review fixed legal declarations and valuations before sharing.", wraplength=970).grid(row=6, column=0, sticky="w", pady=(12, 0))
        ttk.Label(self, textvariable=self.status, wraplength=970).grid(row=7, column=0, sticky="w", pady=(6, 0))

    def apply_language(self, *_):
        # Relabel existing widgets; client values, selection and dirty state survive.
        for widget, message in self.localized:
            widget.configure(text=self.tr(message))
        for key, message in (("khasra", "Khasra number"), ("area", "Area (ha)"), ("lagan", "Lagan")):
            self.tree.heading(key, text=self.tr(message))
        self.refresh_records()
        self.set_status(self.status_message, **self.status_values)
        self.preview.refresh()
        self.update_title()

    def set_status(self, message, **values):
        self.status_message, self.status_values = message, values
        self.status.set(self.tr(message, **values))

    def show_error(self, title, error):
        messagebox.showerror(self.tr(title), error_text(error, self.language.get()), parent=self.master)

    def changed(self, *_):
        if not self.loading:
            self.dirty = True
            self.update_title()

    def update_title(self):
        filename = self.data_path.name if self.data_path else self.tr("Untitled")
        self.master.title(f"{'* ' if self.dirty else ''}{filename} — {self.tr('Land record documents')}")

    def show_preview(self):
        if self.preview.active_var is None:
            self.preview.activate(self.values["client_name"], "Client name")
        self.preview.show()

    def get_client(self):
        client = ClientData(**{key: var.get() for key, var in self.values.items()}, records=tuple(self.records), text_mode="legacy_devlys")
        client.require_complete()
        return client

    def refresh_records(self):
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i, record in enumerate(self.records):
            self.tree.insert("", "end", iid=str(i), values=(record.khasra, record.area, record.lagan))
        self.tree.selection_set([key for key in selected if self.tree.exists(key)])
        from decimal import Decimal, localcontext
        with localcontext() as ctx:
            ctx.prec = 40
            area = sum((Decimal(r.area) for r in self.records), Decimal(0))
            lagan = sum((Decimal(r.lagan) for r in self.records), Decimal(0))
        self.total_label.set(self.tr("{count} records  •  Total area: {area} ha  •  Total Lagan: {lagan}", count=len(self.records), area=f"{area:f}", lagan=f"{lagan:f}"))

    def add(self):
        dialog = RecordDialog(self.master, language=self.language.get(), preview=self.preview)
        self.wait_window(dialog)
        if dialog.result:
            self.records.append(dialog.result)
            self.refresh_records()
            self.changed()

    def edit(self):
        selected = self.tree.selection()
        if len(selected) != 1:
            self.set_status("Select one land record to edit.")
            return
        index = int(selected[0])
        dialog = RecordDialog(self.master, self.records[index], language=self.language.get(), preview=self.preview)
        self.wait_window(dialog)
        if dialog.result:
            self.records[index] = dialog.result
            self.refresh_records()
            self.tree.selection_set(str(index))
            self.changed()

    def remove(self):
        selected = self.tree.selection()
        if not selected:
            self.set_status("Select one or more land records to remove.")
            return
        for index in sorted(map(int, selected), reverse=True):
            del self.records[index]
        self.refresh_records()
        self.changed()

    def may_discard(self):
        if not self.dirty:
            return True
        choice = messagebox.askyesnocancel(self.tr("Unsaved client data"), self.tr("Save your changes before continuing?"), parent=self.master)
        if choice is None:
            return False
        return self.save() if choice else True

    def new(self):
        if not self.may_discard():
            return
        self.loading = True
        for key, var in self.values.items():
            var.set(date.today().isoformat() if key == "document_date" else "")
        self.records = []
        self.data_path = None
        self.loading = False
        self.dirty = False
        self.refresh_records()
        self.update_title()
        self.set_status("New client.")

    def populate(self, client, path=None):
        if client.text_mode != "legacy_devlys":
            raise ValidationError("This desktop form accepts DevLys text only. Unicode client files cannot be converted automatically.")
        self.loading = True
        for key, var in self.values.items():
            var.set(getattr(client, key))
        self.records = list(client.records)
        self.data_path = Path(path) if path else None
        self.loading = False
        self.dirty = False
        self.refresh_records()
        self.update_title()

    def open(self):
        if not self.may_discard():
            return
        path = filedialog.askopenfilename(parent=self.master, title=self.tr("Open client JSON"), filetypes=[(self.tr("Client JSON"), "*.json")])
        if path:
            self.load(path)

    def load(self, path, sample=False):
        try:
            client = load_client(path)
            self.populate(client, None if sample else path)
        except (OSError, ValueError) as error:
            self.show_error("Cannot open client data", error)
            return
        if not client.age or not client.caste:
            self.set_status("Older client loaded. Enter age and caste before saving or generating.")
        elif sample:
            self.set_status("Sample loaded. Save a copy for your client.")
        else:
            self.set_status("Loaded {path}", path=path)

    def sample(self):
        if self.may_discard():
            self.load(ROOT / "sample_data/client.json", sample=True)

    def save(self):
        try:
            client = self.get_client()
        except ValidationError as error:
            self.show_error("Check client data", error)
            return False
        path = filedialog.asksaveasfilename(parent=self.master, title=self.tr("Save client JSON"), defaultextension=".json", initialfile=self.data_path.name if self.data_path else "client.json", initialdir=str(self.data_path.parent) if self.data_path else str(ROOT), filetypes=[(self.tr("Client JSON"), "*.json")])
        if not path:
            return False
        try:
            save_client(client, path)
        except (ValueError, OSError) as error:
            self.show_error("Cannot save client data", error)
            return False
        self.data_path = Path(path)
        self.dirty = False
        self.update_title()
        self.set_status("Saved {path}", path=path)
        return True

    def generate(self):
        try:
            client = self.get_client()
        except ValidationError as error:
            self.show_error("Check client data", error)
            return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(
            parent=self.master,
            title=self.tr("Generate a new Word document"),
            defaultextension=".docx",
            initialfile=client_filename(client.client_name),
            initialdir=str(OUTPUT_DIR),
            filetypes=[(self.tr("Word document"), "*.docx")],
        )
        if not path:
            return
        try:
            self.master.configure(cursor="watch")
            self.master.update_idletasks()
            result = generate_docx(client, path)
        except (ValueError, OSError) as error:
            self.show_error("Cannot generate document", error)
            return
        finally:
            self.master.configure(cursor="")
        self.set_status("Generated {name} — updated all {count} land tables.", name=result.path.name, count=result.land_tables_updated)
        messagebox.showinfo(self.tr("Document generated"), self.tr("Saved {path}\n\nAll three land tables were updated, including the nested table.", path=result.path), parent=self.master)

    def close(self):
        if self.may_discard():
            self.master.destroy()


def run():
    check_tk_version(tk.TkVersion)
    root = tk.Tk()
    root.geometry("1100x820")
    root.minsize(1040, 760)
    Application(root)
    root.mainloop()
