import os
import ctypes
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
from py7zr import SevenZipFile
from zipfile import ZipFile
import sys
def on_drag_enter():
    print("Drag entered!")

def on_drag_leave():
    print("Drag left!")

def on_drag_motion():
    print("Drag in motion!")

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
                        with ZipFile(zip_file_path, 'r') as zip_ref:
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
                    nb_fonts_installed += 1
                    print(f"Installed font: {font_path}")

    if nb_fonts_installed == 0:
        print("No fonts have been installed :(")
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
        unzip_files_in_directory(folder_path)
        install_fonts_in_directory(folder_path)
    else:
        print("No folder selected.")

def on_drop(event, label):
    data = event.data
    if data:
        folder_path = data
        print(f"Dropped Folder: {folder_path}")
        label.config(text=f"Dropped Folder: {folder_path}")
        unzip_files_in_directory(folder_path)
        install_fonts_in_directory(folder_path)
    else:
        print("No folder dropped.")

def gui():
    root = TkinterDnD.Tk()
    root.title("Font Fiesta Unzipper And Installer Extravaganza 🎉🔤")
    # Set the initial size of the window
    root.geometry("800x400")

    # Create a label widget
    label = tk.Label(root, text="Click on the button below and choose the folder where your fonts files are (even if they are compressed)")
    label.pack(pady=10)

    # Create a label to show the selected directory's path
    dir_label = tk.Label(root, text=" ")
    dir_label.pack(pady=10)

    # Create a button widget
    button = tk.Button(root, text="Choose Folder", command=lambda: choose_folder(dir_label))
    button.pack()

    # Create a label for the drag-and-drop area
    drag_label = tk.Label(root, text="Or drag and drop a folder in the circle below:")
    drag_label.pack(pady=10)

    # Create a frame as a circular drag-and-drop area
    drag_area = tk.Frame(root, width=200, height=200, background="green", highlightthickness=0, highlightbackground="gray", bd=0, borderwidth=0)
    drag_area.pack(pady=10)
    drag_area.place(relx=0.5, rely=0.5, anchor="center", x=0, y=50)
    drag_area.pack_propagate(False)

    # Create a circular canvas inside the frame
    canvas = tk.Canvas(drag_area, width=210, height=210, bg="#f0f0f0", highlightthickness=0)
    canvas.pack(fill="both", expand=True )
    canvas.place(relx=0.5, rely=1.0, anchor='s', y=0)

    # Create a circular shape on the canvas
    canvas.create_oval(10, 10, 200, 200, outline="lightblue", fill="lightblue", width=0)

    # Bind events to functions
    drag_area.drop_target_register(DND_FILES)
    drag_area.dnd_bind('<<Drop>>', lambda event: on_drop(event, dir_label))
    drag_area.dnd_bind('<<DragEnter>>', lambda event: on_drag_enter())
    drag_area.dnd_bind('<<DragLeave>>', lambda event: on_drag_leave())
    drag_area.dnd_bind('<<DragMotion>>', lambda event: on_drag_motion())

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
        # Valid arguments. Proceeding to unzipping and font installing without launching GUI.
        unzip_files_in_directory(folder_path)
        install_fonts_in_directory(folder_path)

if __name__ == "__main__":
    main()
