import os
import ctypes
import ctypes.wintypes
import zipfile
from py7zr import SevenZipFile

def unzip_files_in_directory(directory):
    zip_extensions = (".zip", ".7z")

    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith(zip_extensions):
                zip_file_path = os.path.join(root, filename)
                extract_path = os.path.splitext(zip_file_path)[0]

                try:
                    # Check if it's a ZIP file
                    if zip_file_path.endswith(".zip"):
                        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                            zip_ref.extractall(extract_path)

                    # Check if it's a 7z file
                    elif zip_file_path.endswith(".7z"):
                        with SevenZipFile(zip_file_path, mode='r') as zip_ref:
                            zip_ref.extractall(extract_path)

                    print(f"Unzipped: {zip_file_path}")
                except Exception as e:
                    print(f"Error extracting {zip_file_path}: {str(e)}")
                    continue

                print(f"Unzipped: {zip_file_path}")

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
    hwnd_broadcast = 0xFFFF
    wm_fontchange = 0x001D
    ctypes.windll.user32.SendMessageW(hwnd_broadcast, wm_fontchange, 0, 0)

if __name__ == "__main__":
    current_directory = os.getcwd()

    # Unzip files before font installation
    unzip_files_in_directory(current_directory)

    install_fonts_in_directory(current_directory)
