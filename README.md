# Macro Cloner - Standalone Macro Copier for Microsoft Excel Files

A lightweight desktop tool for copying VBA macros from a source Excel file into one or more target `.xlsx` files — without requiring Excel to be open.

Built with Python and ttkbootstrap.

---

## Features

- Auto-detects `source.xlsm` in the same folder as the script
- Select multiple target `.xlsx` files at once (hold **Ctrl** in the file picker)
- Outputs `filename_new.xlsm` alongside each original file — originals are never modified
- No Excel installation required
- Color-coded output log per file

---

## Requirements

- Python 3.10 or later
- [ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap)

Install the dependency:

```bash
pip install ttkbootstrap
```

---

## Usage

1. Place `macro_copier.py` and `source.xlsm` in the same folder
2. Run the script:

```bash
python macro_copier.py
```

3. The **Source file** field is auto-populated with `source.xlsm`. Use **Browse…** to select a different file if needed.
4. Click **Add files…** to select one or more `.xlsx` target files. Hold **Ctrl** to select multiple files at once.
5. Click **Copy Macros**.
6. Output files are saved as `<original_name>_new.xlsm` in the same folder as each target file.

---

## How it works

`.xlsm` and `.xlsx` files are ZIP archives. VBA macros in Excel are stored as a binary blob (`xl/vbaProject.bin`). The tool:

1. Extracts `vbaProject.bin` from `source.xlsm`
2. Injects it into each target `.xlsx` file
3. Updates the internal XML metadata (`[Content_Types].xml` and `xl/_rels/workbook.xml.rels`) to register the macro project
4. Saves the result as a valid `.xlsm` file

---

## File structure

```
macro-copier/
├── macro_copier.py   # Main script
├── source.xlsm       # Your macro source file (not included)
├── LICENSE
└── README.md
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Author

Bo Sundgaard, 2026  
[www.uniteapps.dk](https://www.uniteapps.dk)
