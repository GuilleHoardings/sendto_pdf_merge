import sys
import fitz  # pymupdf
import os
from pathlib import Path
import ctypes
from win32com.client import Dispatch
import tkinter as tk
from tkinter import ttk

SCRIPT_PATH = os.path.abspath(sys.argv[0])


def show_message_box(message, title="Message", icon="info"):
    """Display a message box with a specified icon."""
    icons = {
        "info": 0x40,  # Information icon
        "warning": 0x30,  # Warning icon
        "error": 0x10,  # Error icon
    }
    icon_flag = icons.get(icon, 0x40)  # Default to information icon
    ctypes.windll.user32.MessageBoxW(0, message, title, icon_flag)


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

    create_shortcut(str(exe_path), "Merge PDFs")

    show_message_box("Shortcut added successfully!", "Success", icon="info")


def uninstall_sendto_shortcut():
    """Remove the SendTo shortcut for merge_pdfs.exe."""
    sendto_folder = get_sendto_folder()
    shortcut_path = sendto_folder / "Merge PDFs.lnk"

    if shortcut_path.exists():
        shortcut_path.unlink()
        show_message_box(
            "Shortcut removed successfully.", "Uninstall Complete", icon="info"
        )
    else:
        show_message_box("Shortcut not found.", "Uninstall Error", icon="error")


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
    dialog.geometry("270x110")
    dialog.resizable(False, False)

    # Set a consistent background color
    bg_color = "#f0f0f0"  # Light gray (default for ttk)
    dialog.configure(bg=bg_color)

    # Add a title label
    label = ttk.Label(
        dialog,
        text="Merge PDFs Setup",
        font=("Segoe UI", 12, "bold"),  # Slightly smaller font
        anchor="center",
        background=bg_color,
    )
    label.pack(pady=(10, 5))  # Reduced padding

    # Add a description label
    description = ttk.Label(
        dialog,
        text="What would you like to do?",
        font=("Segoe UI", 10),
        anchor="center",
        background=bg_color,
    )
    description.pack(pady=(0, 10))  # Reduced padding

    # Create a frame for buttons with consistent background
    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=5)  # Reduced padding

    # Style and add buttons
    install_button = ttk.Button(
        button_frame,
        text="Install",
        command=on_install,
        width=10,  # Slightly smaller width
    )
    install_button.grid(row=0, column=0, padx=5)  # Reduced horizontal padding

    uninstall_button = ttk.Button(
        button_frame,
        text="Uninstall",
        command=on_uninstall,
        width=10,
    )
    uninstall_button.grid(row=0, column=1, padx=5)

    cancel_button = ttk.Button(
        button_frame,
        text="Cancel",
        command=on_cancel,
        width=10,
    )
    cancel_button.grid(row=0, column=2, padx=5)

    # Handle window close button
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    # Center the dialog on the screen
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")

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
