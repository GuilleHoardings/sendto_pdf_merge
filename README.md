# PDF Merger

A simple utility to merge multiple PDF files into a single document.

## Overview

PDF Merger is a Windows application that allows you to easily combine multiple PDF files into a single document. It integrates with the Windows "Send To" menu for quick access and provides a simple interface for installation and usage.

## Features

- Merge multiple PDF files with a simple right-click
- Automatically names the output file based on the first selected PDF
- Validates PDFs before merging to ensure they are readable
- Automatically opens the merged PDF when complete
- Easy installation via the setup interface

## Installation

There are two ways to install PDF Merger:

### Method 1: Using the Setup Interface

1. Run the application without any arguments:
   ```bash
   pdf_merger.exe
   ```
   or
   ```bash
   python merge_pdfs.py
   ```

2. In the setup dialog that appears, click "Install" to add the application to your Windows "Send To" menu.

### Method 2: Using Command Line

Run the application with the `--install` flag:

```bash
pdf_merger.exe --install
```
or
```bash
python merge_pdfs.py --install
```

## Usage

1. Select multiple PDF files in Windows Explorer
2. Right-click and select "Send To" > "Merge PDFs"
3. The merged PDF will be created in the same directory as the first selected file
4. The merged file will automatically open after creation

## Uninstallation

### Method 1: Using the Setup Interface

1. Run the application without any arguments:
   ```bash
   pdf_merger.exe
   ```
   or
   ```bash
   python merge_pdfs.py
   ```

2. In the setup dialog that appears, click "Uninstall" to remove the application from your Windows "Send To" menu.

### Method 2: Using Command Line

Run the application with the `--uninstall` flag:

```bash
pdf_merger.exe --uninstall
```
or
```bash
python merge_pdfs.py --uninstall
```

## Requirements

- Windows operating system
- Python 3.6 or higher (if running from source)
- Required Python packages:
  - PyMuPDF (fitz)
  - pywin32

## Technical Details

The application uses PyMuPDF (fitz) to handle PDF operations and integrates with the Windows shell through the pywin32 library. The merged PDF is named after the first selected file with "_merged" appended to the filename.

## License

MIT License. See the [LICENSE](LICENSE) file for details.

## Troubleshooting

If you encounter any issues:
- Ensure all PDFs are valid and not password-protected
- Check that you have proper permissions to write to the output directory

## Development

If you wish to contribute to the project, please fork the repository and submit
a pull request. Ensure that your code adheres to the existing style and includes
appropriate tests.

## Generation of Executable

To generate an executable for the application, you can use PyInstaller. Here are
the steps:
1. Install PyInstaller:
2. ```bash
   pip install pyinstaller
    ```
3. Navigate to the directory containing `merge_pdfs.py` and run:
4. ```bash
   pyinstaller --onefile merge_pdfs.py
   ```
5. This will create a standalone executable in the `dist` folder.
6. You can then distribute this executable without requiring Python or any dependencies to be installed on the target machine.
7. Make sure to test the executable on a clean machine to ensure it works as expected.
8. You can also customize the PyInstaller spec file to include additional files or resources if needed.
9. For more advanced usage, refer to the [PyInstaller documentation](https://pyinstaller.readthedocs.io/en/stable/).
10. You can also use the `--icon` option to specify a custom icon for the executable:
11. ```bash
    pyinstaller --onefile --icon=your_icon.ico merge_pdfs.py
    ```
12. This will embed the specified icon into the executable file.
