import sys
import fitz  # pymupdf
import os
from pathlib import Path
import ctypes
from win32com.client import Dispatch


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


if __name__ == "__main__":
    if len(sys.argv) <= 1:
        show_message_box(
            "No files selected. Please select PDF files to merge.", "No Files"
        )
        print("Usage:")
        print("  merge_pdfs.py file1.pdf file2.pdf [output.pdf]  # Merge PDFs")
        print("  merge_pdfs.py --install  # Add context menu entry")
        print("  merge_pdfs.py --uninstall  # Remove context menu entry")
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
