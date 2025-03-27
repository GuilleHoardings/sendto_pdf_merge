import sys
import fitz  # pymupdf
import os
import ctypes
import winreg

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
    output_file = os.path.join(output_dir, "merged.pdf")

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

    show_message_box(f"Merged successfully: {output_file}", "Success")
    os.startfile(output_file)


def install_context_menu():
    """Add 'Merge PDFs' option to Windows Explorer right-click menu."""
    key_path = r"Software\Classes\SystemFileAssociations\.pdf\Shell\MergePDF"
    command = f'"{SCRIPT_PATH}" "%*"'

    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Merge PDFs")

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\Command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'cmd /c "{command}"')

        print("Context menu entry added. Right-click PDFs to merge.")
    except PermissionError:
        print("Run as Administrator to install the context menu.")


def uninstall_context_menu():
    """Remove the context menu entry."""
    key_path = r"Software\Classes\SystemFileAssociations\.pdf\Shell\MergePDF"

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r"\Command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
        print("Context menu entry removed.")
    except FileNotFoundError:
        print("Context menu entry not found.")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg == "--install":
            install_context_menu()
        elif arg == "--uninstall":
            uninstall_context_menu()
        else:
            show_message_box(
                f"File(s) selected: {', '.join(sys.argv[1:])}", "Selected Files"
            )
            merge_pdfs(sys.argv[1:])
    else:
        show_message_box(
            "No files selected. Please select PDF files to merge.", "No Files"
        )
        print("Usage:")
        print("  merge_pdfs.py file1.pdf file2.pdf [output.pdf]  # Merge PDFs")
        print("  merge_pdfs.py --install  # Add context menu entry")
        print("  merge_pdfs.py --uninstall  # Remove context menu entry")
