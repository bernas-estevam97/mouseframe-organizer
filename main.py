import sys
import os
import ctypes
import logging

# Ensure project root directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Windowed (no-console) builds have sys.stdout/stderr = None; libraries that print warnings
# (customtkinter does on import) would crash the app. Must run before importing customtkinter.
for _name in ("stdout", "stderr"):
    if getattr(sys, _name) is None:
        setattr(sys, _name, open(os.devnull, "w"))

from gui.app import MouseframeOrganizerApp

# Bundled resources live in sys._MEIPASS when frozen with PyInstaller
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = os.path.join(os.path.expanduser("~"), "mouseframe_organizer.log")


def main():
    logging.basicConfig(filename=LOG_FILE, encoding="utf-8", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s")

    # 1. Set explicit Windows AppUserModelID so Taskbar displays the custom icon
    if sys.platform == 'win32':
        try:
            my_appid = 'mouseframe.organizer.gui.v1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(my_appid)
        except Exception:
            pass

    app = MouseframeOrganizerApp()
    # Unhandled errors in Tk callbacks: log the traceback instead of printing to a hidden console
    app.report_callback_exception = lambda *exc: logging.getLogger("tk").error("Unhandled UI error", exc_info=exc)

    # 2. Window/taskbar icon from assets/ folder
    try:
        app.iconbitmap(os.path.join(BASE_DIR, "assets", "app_icon.ico"))
    except Exception:
        pass  # missing icon or non-Windows Tk: keep default icon

    app.mainloop()


if __name__ == "__main__":
    main()
