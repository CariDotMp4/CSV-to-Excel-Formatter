# CSV to Excel Formatter

A Python utility that converts CSV files to professionally formatted Excel spreadsheets with a elegant purple theme.

## Features

- **GUI Interface**: User-friendly graphical interface for easy file selection and formatting
- **Command-Line Support**: Batch processing via CLI for automation
- **Professional Styling**: 
  - Purple color scheme (headers and footers)
  - Alternating row colors (zebra striping)
  - Proper borders and fonts
  - Auto-fitted column widths
- **Flexible Input**: Supports multiple CSV separators (`,`, `;`, `\t`)
- **Character Encoding**: Handles various encodings (UTF-8, Latin-1, etc.)
- **Standalone Executable**: Generate `.exe` on Windows or binaries on macOS/Linux using PyInstaller

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)

### Setup

1. Clone the repository:
```bash
git clone <repository-url>
cd CSV-to-Excel-Formatter
```

2. Install dependencies:
```bash
pip install pandas openpyxl pyinstaller
```

## Usage

### GUI Mode (Default)
Simply run the script:
```bash
python csv_to_excel.py
```

A file browser window will open where you can:
1. Select your CSV file
2. Choose the output location
3. Enter a title for the spreadsheet
4. Click "Generate Excel" to create the formatted file

### Command Line Mode
```bash
python csv_to_excel.py --cli \
  --csv "path/to/input.csv" \
  --saida "path/to/output.xlsx" \
  --titulo "My Report" \
  --sep ";" \
  --enc "utf-8"
```

#### Parameters:
| Parameter | Description | Default |
|-----------|-------------|---------|
| `--csv` | Path to input CSV file | Required |
| `--saida` | Path for output Excel file | Required |
| `--titulo` | Spreadsheet title | "Data Report" |
| `--sep` | CSV delimiter: `,` `;` or `\t` | `,` |
| `--enc` | Character encoding (utf-8, latin-1, etc.) | utf-8 |

## Building Executable

### Windows (.exe)
```bash
pyinstaller --onefile --windowed --name "CSV_Excel_Formatter" csv_to_excel.py
```

### macOS / Linux
```bash
pyinstaller --onefile --name "csv_to_excel" csv_to_excel.py
```

The executable will be created in the `dist/` folder.

## Output Format

The generated Excel file includes:
- **Title Row**: Centered, bold, purple background
- **Header Row**: White text on purple background, bold
- **Data Rows**: Alternating white and light purple backgrounds
- **Footer**: Summary information (if provided)
- **Borders**: Professional purple borders
- **Formatting**: Auto-fitted columns, proper alignment

## Dependencies

- **pandas**: CSV reading and data manipulation
- **openpyxl**: Excel file generation and styling
- **pyinstaller**: (Optional) For creating standalone executables

## License

[Add your license here]

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## Author

Created as a practical tool for efficient CSV to Excel conversion with professional formatting.
