"""Desktop integration smoke check; requires a graphical Tk session.

Run from the repository root: python -m scripts.verify_gui
No dialogs are shown, and all generated data stays in a temporary directory.
"""
from pathlib import Path
import tempfile
from types import SimpleNamespace
from unittest.mock import patch
import tkinter as tk

from landdocs.gui import Application
from landdocs.model import LandRecord, load_client
from landdocs.engine import ROOT
from landdocs.runtime import check_tk_version


def main():
    check_tk_version(tk.TkVersion)
    root = tk.Tk()
    root.withdraw()
    try:
        app = Application(root)
        app.populate(load_client(ROOT / "sample_data/client.json"))
        root.update_idletasks()
        assert len(app.tree.get_children()) == 4
        assert "1.229" in app.total_label.get()
        extra = LandRecord("0002/03", "0.1", "0.2")
        with patch("landdocs.gui.RecordDialog", return_value=SimpleNamespace(result=extra)), patch.object(app, "wait_window"):
            app.add()
        assert len(app.records) == 5 and app.dirty
        app.tree.selection_set("4")
        revised = LandRecord("0002/04", "0.2", "0.3")
        with patch("landdocs.gui.RecordDialog", return_value=SimpleNamespace(result=revised)), patch.object(app, "wait_window"):
            app.edit()
        assert app.records[4] == revised
        app.tree.selection_set("0", "4")
        app.remove()
        assert len(app.records) == 3
        assert tuple(app.tree.cget("columns")) == ("khasra", "area", "lagan")
        assert app.preview.watched and app.preview.display is not None
        assert app.get_client().text_mode == "legacy_devlys"
        before = app.get_client()
        app.tree.selection_set("1")
        app.set_status("Loaded {path}", path="test.json")
        dirty = app.dirty
        app.language.set("हिन्दी")
        assert app.get_client() == before and app.dirty == dirty
        assert app.tree.selection() == ("1",)
        assert app.tree.heading("khasra", "text") == "खसरा क्रमांक"
        assert "भूमि अभिलेख" in root.title() and "कुल क्षेत्रफल" in app.total_label.get()
        assert app.status.get() == "लोड हुआ: test.json"
        app.show_preview()
        root.update_idletasks()
        assert app.preview.display.get("1.0", "end-1c") == app.values["client_name"].get()
        from landdocs.gui import RecordDialog
        dialog = RecordDialog(root, language="हिन्दी")
        assert tuple(dialog.values) == ("khasra", "area", "lagan")
        assert dialog.title() == "भूमि अभिलेख जोड़ें"
        dialog.values["khasra"].set("0003/1")
        dialog.values["area"].set("bad")
        dialog.values["lagan"].set("0")
        with patch("landdocs.gui.messagebox.showerror") as error:
            dialog.apply()
            assert "क्षेत्रफल" in error.call_args.args[1]
        dialog.destroy()
        app.language.set("English")
        assert app.get_client() == before and app.tree.selection() == ("1",)
        with tempfile.TemporaryDirectory() as directory:
            json_path = Path(directory) / "client.json"
            docx_path = Path(directory) / "client.docx"
            with patch("landdocs.gui.filedialog.asksaveasfilename", return_value=str(json_path)):
                assert app.save()
            assert not app.dirty and load_client(json_path) == app.get_client()
            with patch("landdocs.gui.filedialog.asksaveasfilename", return_value=str(docx_path)), patch("landdocs.gui.generate_docx", return_value=SimpleNamespace(path=docx_path, land_tables_updated=3)), patch("landdocs.gui.messagebox.showinfo") as info:
                app.generate()
                info.assert_called_once()
            assert "all 3 land tables" in app.status.get()
            with patch("landdocs.gui.messagebox.showerror") as error:
                app.load(Path(directory) / "absent.json")
                error.assert_called_once()
            app.values["client_name"].set("")
            with patch("landdocs.gui.messagebox.showerror") as error:
                assert not app.save()
                error.assert_called_once()
            assert app.dirty
            with patch("landdocs.gui.messagebox.askyesnocancel", return_value=None):
                assert not app.may_discard()
            with patch("landdocs.gui.messagebox.askyesnocancel", return_value=False):
                app.new()
            assert not app.records and not app.dirty
        print(f"GUI integration passed with Tk {tk.TkVersion}: add, edit, remove, totals, DevLys preview, Hindi/English switching, translated dialogs, save/load, generate, validation, unsaved changes. Hidden-window checks do not verify visible rendering.")
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
