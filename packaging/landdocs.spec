# PyInstaller spec for the Tkinter desktop application.
from pathlib import Path
import sys

project_root = Path(SPECPATH).resolve().parent

a = Analysis(
    [str(project_root / "app.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / "templates"), "templates"),
        (str(project_root / "sample_data"), "sample_data"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="LandDocs",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
if sys.platform == "darwin":
    app = BUNDLE(exe, name="LandDocs.app", icon=None)
