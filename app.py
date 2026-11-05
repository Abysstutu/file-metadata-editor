"""Application entry point with Windows high-DPI support enabled before Tk starts."""

import ctypes
import os
import subprocess
import sys
import tempfile
from pathlib import Path


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
    if sys.argv[1:] == ["--self-test-media"]:
        from metadata_editor.engine import _ffmpeg, _ffprobe

        with tempfile.TemporaryDirectory() as folder:
            sample = Path(folder) / "sample.mp4"
            hidden = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.run([_ffmpeg(), "-v", "error", "-f", "lavfi", "-i", "color=c=blue:s=64x64:d=0.2",
                            "-c:v", "mpeg4", "-y", str(sample)], check=True, capture_output=True, creationflags=hidden)
            subprocess.run([_ffprobe(), "-v", "error", "-show_format", str(sample)],
                           check=True, capture_output=True, creationflags=hidden)
    else:
        launch()
