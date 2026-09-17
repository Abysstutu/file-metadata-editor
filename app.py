"""Application entry point with Windows high-DPI support enabled before Tk starts."""

import ctypes
import os


def enable_high_dpi() -> None:
    if os.name != "nt":
        return
    try:
        # Per-monitor V2 avoids Windows bitmap-stretching the complete Tk window
        # when it is moved between monitors using different display scales.
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except (AttributeError, OSError):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()


enable_high_dpi()

from metadata_editor.ui import launch


if __name__ == "__main__":
    launch()
