import sys
import os
import ctypes

# Ensure project root directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import MouseframeOrganizerApp


def main():
    # 1. Set explicit Windows AppUserModelID so Taskbar displays the custom icon
    if sys.platform == 'win32':
        try:
            my_appid = 'mouseframe.organizer.gui.v1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(my_appid)
        except Exception:
            pass

    app = MouseframeOrganizerApp()

    # 2. Automatically load custom app icon from assets/ folder
    assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
    ico_path = os.path.join(assets_dir, "app_icon.ico")
    png_path = os.path.join(assets_dir, "app_icon.png")

    if os.path.isfile(ico_path):
        try:
            app.iconbitmap(ico_path)
        except Exception:
            pass
    elif os.path.isfile(png_path):
        try:
            from PIL import Image, ImageTk
            icon_img = ImageTk.PhotoImage(Image.open(png_path))
            app.wm_iconphoto(True, icon_img)
        except Exception:
            pass

    app.mainloop()


if __name__ == "__main__":
    main()
