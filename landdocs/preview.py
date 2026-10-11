"""Inline live DevLys glyph preview; no guessed character conversion."""
import tkinter as tk
from tkinter import font as tkfont, ttk

from .i18n import translate

DEVLYS_FONT = "DevLys 010"
HUMAN_FIELDS = {"client_name", "father_name", "caste", "address", "village", "tehsil", "district"}


class DevLysPreview:
    def __init__(self, parent, language):
        self.parent = parent
        self.language = language
        self.active_var = None
        self.active_label = "Client name"
        self.watched = {}
        self.frame = None
        self.display = None
        self.warning = None

    def tr(self, message, **values):
        return translate(message, self.language(), **values)

    def mount(self, parent):
        self.frame = ttk.Frame(parent)
        self.frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.frame.columnconfigure(1, weight=1)
        ttk.Label(self.frame, text=self.tr("DevLys preview")).grid(row=0, column=0, sticky="nw", padx=(0, 12))
        self.display = tk.Text(self.frame, height=2, width=52, wrap="word", state="disabled", takefocus=False, font=(DEVLYS_FONT, 20))
        self.display.grid(row=0, column=1, sticky="ew")
        self.warning = ttk.Label(self.frame, foreground="#a52323", wraplength=700)
        self.warning.grid(row=1, column=1, sticky="w", pady=(3, 0))
        self.refresh()

    def watch(self, entry, variable, label, display_font=DEVLYS_FONT):
        trace = variable.trace_add("write", self.refresh)
        self.watched[entry] = (variable, trace)
        entry.bind("<FocusIn>", lambda _: self.activate(variable, label, display_font), add="+")
        entry.bind("<Destroy>", lambda _: self.unwatch(entry), add="+")

    def unwatch(self, entry):
        watched = self.watched.pop(entry, None)
        if watched:
            variable, trace = watched
            variable.trace_remove("write", trace)
            if self.active_var is variable:
                self.active_var = None
                self.refresh()

    def activate(self, variable, label, display_font=DEVLYS_FONT):
        self.active_var, self.active_label = variable, label
        self.refresh()

    def show(self):
        # Compatibility hook: the preview is inline, never a separate window.
        self.refresh()

    @staticmethod
    def set_text(widget, value):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)
        widget.configure(state="disabled")

    def refresh(self, *_):
        if self.display is None:
            return
        value = self.active_var.get() if self.active_var is not None else ""
        self.set_text(self.display, value)
        families = {family.casefold() for family in tkfont.families(self.parent)}
        warning = "" if DEVLYS_FONT.casefold() in families else self.tr("DevLys 010 is not installed. Install it with Font Book for an accurate Hindi preview.")
        self.warning.configure(text=warning)
