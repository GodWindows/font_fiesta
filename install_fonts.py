import os
import ctypes
import ctypes.wintypes
import sys
import tkinter as tk
import zipfile
from py7zr import SevenZipFile
from tkinter import filedialog

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
    nb_fonts_installed = 0

    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith(font_extensions):
                font_path = os.path.join(root, filename)
                result = ctypes.windll.gdi32.AddFontResourceW(font_path)
                if result == 0:
                    print(f"Failed to install font: {font_path}")
                else:
                    nb_fonts_installed+= 1
                    print(f"Installed font: {font_path}")
    if nb_fonts_installed == 0:
        print("No font have been installed :(")
    else:
        print("We just installed", nb_fonts_installed, "font(s). Awesome!")
    # Notify the system that the font cache has changed
    hwnd_broadcast = 0xFFFF
    wm_fontchange = 0x001D
    ctypes.windll.user32.SendMessageW(hwnd_broadcast, wm_fontchange, 0, 0)

def choose_folder(label):
    # Open a dialog for selecting a folder
    folder_path = filedialog.askdirectory()
    
    # Check if a folder was selected
    if folder_path:
        print(f"Selected Folder: {folder_path}")
        label.config(text=f"Selected Folder: {folder_path}")
    else:
        print("No folder selected.")

def gui():
    root = tk.Tk()
    root.title("Font Fiesta Unzipper And Installer Extravaganza 🎉🔤")
    # Set the initial size of the window
    root.geometry("800x400")

    # Create a label widget
    label = tk.Label(root, text="Click on the button below and choose the folder where your fonts files are (even if they are compressed)")
    label.pack(pady=10)

    # Create a label to show the selected directory's path
    dir = tk.Label(root, text=" ")
    # Create a button widget
    button = tk.Button(root, text="Choose Folder", command=lambda: choose_folder(dir))
    button.pack()

    dir.pack(pady=10)

    # Start the Tkinter event loop
    root.mainloop()
    

def main():
    # Check the number of command-line arguments
    if len(sys.argv) == 1:
        # Launch the GUI if no arguments were passed
        gui()
        return
    
    if len(sys.argv) != 2:
        # Launch the GUI if arguments are invalid
        print("Invalid arguments. Launching GUI.")
        gui()
        return
    
    folder_path = sys.argv[1]

    # Check if the specified folder exists
    if not os.path.exists(folder_path):
        # Launch the GUI if the folder doesn't exist
        print(f"Error: The specified folder path '{folder_path}' does not exist. Launching GUI.")
        gui()
    else:
        # Valid arguments. Proceeding to unzipping and font installng without launching GUI.
        unzip_files_in_directory(folder_path)
        install_fonts_in_directory(folder_path)




if __name__ == "__main__":
    main()