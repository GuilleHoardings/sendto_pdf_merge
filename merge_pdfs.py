import sys
import fitz  # pymupdf
import os
from pathlib import Path
import ctypes
from win32com.client import Dispatch
import tkinter as tk
from tkinter import messagebox

SCRIPT_PATH = os.path.abspath(sys.argv[0])


def show_message_box(message, title="Error"):
    """Display an error message box."""
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)


def is_valid_pdf(file_path):
    """Check if a PDF file is valid and readable."""
    try:
        with fitz.open(file_path) as doc:
            return doc.page_count > 0  # Ensure it has pages
    except Exception as e:
        show_message_box(f"Invalid PDF {file_path}: {e}", "Invalid PDF")
        return False  # Invalid or corrupted PDF


def merge_pdfs(files):
    """Merge multiple PDFs and save in the same directory as the first file."""
    if len(files) < 2:
        show_message_box("Error: Select at least two PDFs.", "Error")
        return

    # Get the directory of the first PDF
    output_dir = os.path.dirname(files[0])

    # Get the output file name from the first PDF
    output_file_name = os.path.basename(files[0]).replace(".pdf", "_merged.pdf")
    output_file = os.path.join(output_dir, output_file_name)

    doc = fitz.open()
    for pdf in files:
        # if not is_valid_pdf(pdf):
        #     show_message_box(f"Skipping invalid PDF: {pdf}", "Invalid PDF")
        #     return
        try:
            doc.insert_pdf(fitz.open(pdf))
        except Exception as e:
            show_message_box(f"Error merging {pdf}: {e}", "Merge Error")
            return

    if doc.page_count == 0:
        show_message_box("No valid PDFs to merge.", "Error")
        return

    doc.save(output_file)
    doc.close()

    # show_message_box(f"Merged successfully: {output_file}", "Success")
    os.startfile(output_file)


def get_sendto_folder():
    """Retrieve the path to the SendTo folder."""
    return Path(os.path.expandvars(r"%APPDATA%\Microsoft\Windows\SendTo"))


def create_shortcut(target, shortcut_name):
    """Create a shortcut in the SendTo folder."""
    sendto_folder = get_sendto_folder()
    shortcut_path = sendto_folder / f"{shortcut_name}.lnk"

    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = str(Path(target).parent)
    shortcut.Save()

    return shortcut_path


def install_sendto_shortcut():
    """Install the SendTo shortcut for merge_pdfs.exe."""
    exe_path = Path(sys.argv[0]).resolve()

    shortcut = create_shortcut(str(exe_path), "Merge PDFs")

    show_message_box(f"Shortcut added: {shortcut}", "Installation Complete")


def uninstall_sendto_shortcut():
    """Remove the SendTo shortcut for merge_pdfs.exe."""
    sendto_folder = get_sendto_folder()
    shortcut_path = sendto_folder / "Merge PDFs.lnk"

    if shortcut_path.exists():
        shortcut_path.unlink()
        show_message_box("Shortcut removed successfully.", "Uninstall Complete")
    else:
        show_message_box("Shortcut not found.", "Uninstall Error")


def show_install_options():
    """Show a custom dialog box with Install, Uninstall, and Cancel options."""

    def on_install():
        install_sendto_shortcut()
        root.destroy()

    def on_uninstall():
        uninstall_sendto_shortcut()
        root.destroy()

    def on_cancel():
        root.destroy()

    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Create a custom dialog box
    dialog = tk.Toplevel(root)
    dialog.title("Merge PDFs Setup")
    dialog.geometry("300x150")
    dialog.resizable(False, False)

    label = tk.Label(dialog, text="What would you like to do?", font=("Arial", 12))
    label.pack(pady=20)

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)

    install_button = tk.Button(
        button_frame, text="Install", command=on_install, width=10
    )
    install_button.grid(row=0, column=0, padx=5)

    uninstall_button = tk.Button(
        button_frame, text="Uninstall", command=on_uninstall, width=10
    )
    uninstall_button.grid(row=0, column=1, padx=5)

    cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel, width=10)
    cancel_button.grid(row=0, column=2, padx=5)

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)  # Handle window close button
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) <= 1:
        show_install_options()
        sys.exit(0)

    arg = sys.argv[1]
    if arg == "--install":
        install_sendto_shortcut()
        sys.exit(0)

    if arg == "--uninstall":
        uninstall_sendto_shortcut()
        sys.exit(0)

    # show_message_box(f"File(s) selected: {', '.join(sys.argv[1:])}", "Selected Files")
    merge_pdfs(sys.argv[1:])
