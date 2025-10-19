# Persian PDF to Excel Converter

A Python tool to convert Persian/Arabic PDF tables to Excel format with proper Right-to-Left (RTL) text support.

## Features

- ✅ Handles Persian/Arabic RTL text correctly
- ✅ Preserves text direction while keeping numbers readable
- ✅ Auto-detects and extracts tables from PDF
- ✅ Creates Excel files with RTL layout
- ✅ Properly formats headers and data
- ✅ Auto-adjusts column widths
- ✅ Freezes header row for easy scrolling

## Requirements

```bash
pip install pdfplumber openpyxl
```

## Installation

1. Clone this repository or download the script
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Convert a PDF file to Excel (auto-generates output filename):
```bash
python pdf_to_excel_improved.py input.pdf
```

Specify output filename:
```bash
python pdf_to_excel_improved.py input.pdf output.xlsx
```

### Advanced Options

```bash
python pdf_to_excel_improved.py input.pdf -o output.xlsx -s "My Data" -f "B Nazanin"
```

**Arguments:**
- `input.pdf` - Input PDF file (required)
- `output.xlsx` - Output Excel file (optional)
- `-o, --output` - Output file path
- `-s, --sheet` - Sheet name (default: "Data")
- `-f, --font` - Font name (default: "Arial", can use "B Nazanin" for Persian)

### Examples

```bash
# Simple conversion
python pdf_to_excel_improved.py bank_statement.pdf

# Custom output name
python pdf_to_excel_improved.py report.pdf converted_report.xlsx

# With custom sheet name and font
python pdf_to_excel_improved.py data.pdf -o result.xlsx -s "تراکنش‌ها" -f "B Nazanin"
```

## How It Works

1. **Extracts tables** from PDF using pdfplumber
2. **Fixes text direction** - Reverses Persian/Arabic text while keeping numbers intact
3. **Creates RTL Excel** - Sets up proper right-to-left layout
4. **Formats cells** - Applies borders, alignment, and styling
5. **Auto-adjusts** - Sets appropriate column widths and row heights

## Text Direction Handling

The script intelligently handles mixed content:
- Persian/Arabic text: Reversed for proper RTL display
- Numbers and dates: Kept in original order
- Mixed text: Each part handled correctly

## Output Format

The generated Excel file includes:
- RTL (Right-to-Left) sheet layout
- First row from PDF as headers (bold, centered)
- Data rows with right-aligned text
- Numbers center-aligned for readability
- Border around all cells
- Frozen header row
- Auto-adjusted column widths

## Troubleshooting

**Issue:** Persian text appears broken or reversed
- **Solution:** The script automatically fixes this. Make sure you're using the latest version.

**Issue:** Numbers appear reversed
- **Solution:** The script preserves number order. If you see reversed numbers, they might be in the original PDF.

**Issue:** Missing data
- **Solution:** Adjust `min_table_rows` and `min_columns` parameters in `extract_tables_from_pdf()` function.

## License

MIT License - Feel free to use and modify!

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

Created for handling Persian/Arabic PDF documents with proper text direction support.
