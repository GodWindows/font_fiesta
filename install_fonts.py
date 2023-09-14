import os
import ctypes
import ctypes.wintypes

def install_fonts_in_directory(directory):
    font_extensions = (".ttf", ".otf", ".fon", ".fnt")

    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith(font_extensions):
                font_path = os.path.join(root, filename)
                result = ctypes.windll.gdi32.AddFontResourceW(font_path)
                if result == 0:
                    print(f"Failed to install font: {font_path}")
                else:
                    print(f"Installed font: {font_path}")

    # Notify the system that the font cache has changed
    # ctypes.windll.user32.SendMessageW(ctypes.wintypes.HWND_BROADCAST, ctypes.wintypes.WM_FONTCHANGE, 0, 0)
    hwnd_broadcast = 0xFFFF
    wm_fontchange = 0x001D
    ctypes.windll.user32.SendMessageW(hwnd_broadcast, wm_fontchange, 0, 0)

if __name__ == "__main__":
    current_directory = os.getcwd()
    install_fonts_in_directory(current_directory)
