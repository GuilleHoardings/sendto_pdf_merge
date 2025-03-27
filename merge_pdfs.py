import sys
import fitz  # pymupdf
import os
import ctypes
import winreg

SCRIPT_PATH = os.path.abspath(__file__)


def merge_pdfs(files, output="merged.pdf"):
    """Merge multiple PDF files into one."""
    if len(files) < 2:
        print("Select at least two PDFs.")
        return

    doc = fitz.open()
    for pdf in files:
        doc.insert_pdf(fitz.open(pdf))

    doc.save(output)
    doc.close()
    print(f"Merged: {output}")
    os.startfile(output)  # Open merged file


def install_context_menu():
    """Add 'Merge PDFs' option to Windows Explorer right-click menu."""
    key_path = r"Software\Classes\SystemFileAssociations\.pdf\Shell\MergePDF"
    command = f'"{SCRIPT_PATH}" "%1" "%2"'

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
            merge_pdfs(sys.argv[1:])
    else:
        print("Usage:")
        print("  merge_pdfs.py file1.pdf file2.pdf [output.pdf]  # Merge PDFs")
        print("  merge_pdfs.py --install  # Add context menu entry")
        print("  merge_pdfs.py --uninstall  # Remove context menu entry")
