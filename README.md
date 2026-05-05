# VBA Macro Copier - Standalone Macro Copier for Microsoft Excel Files

A lightweight desktop tool for copying VBA macros from a source Excel file into one or more target `.xlsx` files — without requiring Excel to be open.

Built with Python and ttkbootstrap. Available as a standalone `.exe` executable or as a Python script.

## What is VBA Macro Copier?

VBA Macro Copier solves a common problem: you have Excel VBA macros in one file and need to apply them to multiple other Excel workbooks. Instead of manually copying macros through Excel's VBA editor (which requires Excel to be running), this tool automates the process by directly manipulating the underlying file structures.

**Key advantages:**
- **No Excel required** — Works independently without Microsoft Office
- **Batch processing** — Apply macros to multiple files at once
- **Safe** — Original files are never modified; outputs are named with `_new` suffix
- **Fast** — Processes files in seconds
- **Simple UI** — User-friendly interface with real-time progress logging

---

## Use Cases

- **Distribute macros** to team members across multiple spreadsheets
- **Update legacy workbooks** with new functionality
- **Automation workflows** where you need to inject macros programmatically
- **Server/headless environments** where Excel is not available
- **Bulk macro deployment** across large numbers of Excel files

---

## Features

- ✓ Auto-detects `source.xlsm` in the same folder as the script
- ✓ Select multiple target `.xlsx` files at once (hold **Ctrl** in the file picker)
- ✓ Outputs `filename_new.xlsm` alongside each original file — originals are never modified
- ✓ No Excel installation required
- ✓ Color-coded output log per file (success/error status)
- ✓ Preserves original spreadsheet content and formatting
- ✓ Automatically converts `.xlsx` to `.xlsm` format (macro-enabled)

## Installation & Getting Started

### Option 1: Use the Compiled Executable (Easiest)

Download `macro_copier.exe` from the [releases](../../releases) page. No installation required — just run it!

1. Place `source.xlsm` in the same folder as `vba_macro_copier.exe`
2. Run `vba_macro_copier.exe`
3. Follow the on-screen instructions

### Option 2: Run from Python Source

**Requirements:**
- Python 3.10 or later
- [ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap)

**Install dependency:**
```bash
pip install ttkbootstrap
```

**Run the script:**
```bash
python macro_copier.py
```

---

## Requirements

---

## Usage

1. Place `macro_copier.py` (or `macro_copier.exe`) and `source.xlsm` in the same folder
2. Run the application
3. The **Source file** field is auto-populated with `source.xlsm`. Use **Browse…** to select a different file if needed.
4. Click **Add files…** to select one or more `.xlsx` target files. Hold **Ctrl** to select multiple files at once.
5. Click **Copy Macros**.
6. Output files are saved as `<original_name>_new.xlsm` in the same folder as each target file.

---

## Technical Details

### How It Works

`.xlsm` and `.xlsx` files are actually ZIP archives containing XML and binary files. VBA macros in Excel are stored as a binary blob (`xl/vbaProject.bin`) along with metadata in XML files.

VBA Macro Copier performs these steps:

1. **Extracts** the VBA project (`vbaProject.bin`) from the source `.xlsm` file
2. **Injects** it into each target `.xlsx` file
3. **Updates** the internal XML metadata to register the macro project:
   - `[Content_Types].xml` — Registers the VBA project content type
   - `xl/_rels/workbook.xml.rels` — Creates relationship link to the macro project
4. **Saves** the result as a valid macro-enabled `.xlsm` file

### Limitations & Important Notes

- **Macro dependencies:** Ensure your VBA code doesn't reference external libraries or ActiveX controls not available on the target system
- **Excel compatibility:** The resulting files are fully compatible with Excel 2010 and later
- **File size:** The output files will be slightly larger due to the added macro project
- **Source file requirement:** A valid source `.xlsm` file with the `vbaProject.bin` binary is required

---

## Project Structure

```
macro-copier/
├── macro_copier.py          # Main application source code
├── build.bat                # Build script to compile to .exe
├── icon.ico                 # Application icon
├── source.xlsm              # Your macro source file (not included)
├── dist/                    # Compiled executable output
├── LICENSE                  # MIT License with Commons Clause
└── README.md                # This file
```

---

## Building from Source

To compile the Python script into a standalone Windows executable:

```bash
build.bat
```

This will use Nuitka to create `dist/vba_macro_copier.exe` (~12 MB) with full icon support.

---

## Troubleshooting

### "vbaProject.bin not found" error
- Ensure your source file is a valid `.xlsm` file with macros
- The file must be created in Excel (not a converted `.xlsx`)
- Try opening and re-saving the source file in Excel

### Output files won't open
- Ensure target `.xlsx` files are not corrupted
- Try with a simple test file first
- Verify that your VBA code is compatible with the target system

### Executable won't run
- Ensure Windows Defender or antivirus software hasn't quarantined it
- Try running from a different folder
- Check that you have write permissions in the working directory

---

## License

MIT License with Commons Clause — see [LICENSE](LICENSE) for details.

This allows use, copying, modification, and redistribution for non-commercial purposes. Commercial use or reselling is not permitted.

---

## Author

Bo Sundgaard, 2026  
[www.uniteapps.dk](https://www.uniteapps.dk)
