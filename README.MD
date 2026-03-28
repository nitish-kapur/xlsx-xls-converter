# XLSX to XLS Converter

A Python script that converts an Excel `.xlsx` file to the legacy `.xls` format using xlwings, with a timestamped output filename and bold column headers preserved. Only the first sheet of the input file is converted.

---

## Author

**Nitish Kapur**<br>
GitHub: [github.com/nitish-kapur](https://github.com/nitish-kapur)

---

## Requirements
```bash
pip install pandas openpyxl xlwings
```

> `xlwings` requires Microsoft Excel to be installed on the system. This script is therefore **Windows or macOS only**.

---

## Usage

Run the script directly:
```bash
python xlsx-to-xls-converter.py
```

A file dialog will open in the script's directory. Select your `.xlsx` file and the converted `.xls` file will be saved automatically in the working directory.

---

## Output

| File | Description |
|---|---|
| `<original_filename>_<timestamp>.xls` | Converted file saved in the working directory |

---

## Notes

- The script opens Excel invisibly in the background via xlwings — do not be alarmed if Excel briefly appears in the taskbar.
- Column headers are written in **bold** in the output file.
- Only the first sheet of the input file is converted.
- Requires Microsoft Excel to be installed; this script will not work on Linux.
