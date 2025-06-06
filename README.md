# SendTo PDF Merge

**SendTo PDF Merge** is a small utility that lets you merge PDF files directly from the Windows "Send to" menu.

## Overview

SendTo PDF Merge is a Windows application that allows you to easily combine multiple PDF files into a single document. It integrates with the Windows "Send To" menu for quick access and provides a simple interface for installation and usage.

![SendTo PDF Merge icon](sendto_pdf_merge_icon.png)

## Features

- Merge multiple PDF files with a simple right-click
- Automatically names the output file based on the first selected PDF
- Validates PDFs before merging to ensure they are readable
- Automatically opens the merged PDF when complete
- Easy installation via the setup interface

## Installation

There are two ways to install SendTo PDF Merge:

### Method 1: Using the Setup Interface

1. Run the application without any arguments:
   ```bash
   sendto_pdf_merge.exe
   ```
   or
   ```bash
   python sendto_pdf_merge.py
   ```

2. In the setup dialog that appears, click "Install" to add the application to your Windows "Send To" menu.

### Method 2: Using Command Line

Run the application with the `--install` flag:

```bash
sendto_pdf_merge.exe --install
```
or
```bash
python sendto_pdf_merge.py --install
```

**Important:** It does not matter which method you use, but you cannot remove
the .exe file from the directory where you installed it. The application will
not work if you do so.

## Usage

1. Select multiple PDF files in Windows Explorer
2. Right-click and select "Send To" > "Merge PDFs"
3. The merged PDF will be created in the same directory as the first selected file
4. The merged file will automatically open after creation

## Uninstallation

### Method 1: Using the Setup Interface

1. Run the application without any arguments:
   ```bash
   sendto_pdf_merge.exe
   ```
   or
   ```bash
   python sendto_pdf_merge.py
   ```

2. In the setup dialog that appears, click "Uninstall" to remove the application from your Windows "Send To" menu.

### Method 2: Using Command Line

Run the application with the `--uninstall` flag:

```bash
sendto_pdf_merge.exe --uninstall
```
or
```bash
python sendto_pdf_merge.py --uninstall
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
3. Navigate to the directory containing `sendto_pdf_merge.py` and run:
4. ```bash
   pyinstaller --onefile --icon=icon.ico --add-data "icon.ico;." --windowed sendto_pdf_merge.py
   ```
5. This will create a standalone executable in the `dist` folder.
6. You can then distribute this executable without requiring Python or any dependencies to be installed on the target machine.
7. Make sure to test the executable on a clean machine to ensure it works as expected.
8. You can also customize the PyInstaller spec file to include additional files or resources if needed.
9. For more advanced usage, refer to the [PyInstaller documentation](https://pyinstaller.readthedocs.io/en/stable/).
