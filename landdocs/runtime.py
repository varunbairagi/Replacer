"""Check GUI runtime compatibility before creating a native window."""
import sys


def check_tk_version(tk_version, platform=None):
    platform = sys.platform if platform is None else platform
    if platform == "darwin" and tuple(map(int, str(tk_version).split(".")[:2])) < (8, 6):
        raise RuntimeError(
            f"This Mac Python uses unsupported Tcl/Tk {tk_version}, which can open a blank white window. "
            "Use Python with Tk 8.6 or newer (for example, a current python.org macOS installer), "
            "then recreate .venv with that interpreter and install requirements.txt. See README.md."
        )
