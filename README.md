# Bank Statement to Excel Converter

## Overview

This Python tool extracts transaction data from Punjab & Sind Bank PDF statements and converts them into well-formatted Excel files. It handles the semi-structured format of bank statements, preserving all transaction details with proper formatting.

## Features

- Extracts transaction data from PDF bank statements
- Preserves all transaction details (date, description, amount, balance)
- Handles multi-line transaction descriptions
- Formats output Excel file with proper:
  - Date formatting
  - Number formatting (with thousand separators)
  - Column alignment
  - Borders and styling
- Auto-sizes columns based on content

## Installation

1. Clone this repository:
   ```bash
   https://github.com/ashishsai960/Scoreme-Placement-Drive_-Hackathon.git
   ```

2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

```bash
python assignment_iitropar_scoreme_2021CHB1046.py input.pdf output.xlsx
```

### Options

- `input.pdf`: Path to your Punjab & Sind Bank statement PDF
- `output.xlsx`: Path for the output Excel file (optional, defaults to "statement.xlsx")

### Example

```bash
python assignment_iitropar_scoreme_2021CHB1046.py statement.pdf transactions.xlsx
```

## Requirements

- Python 3.6+
- Required packages (automatically installed with requirements.txt):
  - pdfplumber
  - pandas
  - openpyxl

## How It Works

1. **PDF Text Extraction**: Uses pdfplumber to extract text while preserving layout
2. **Transaction Parsing**:
   - Identifies transaction lines by date patterns
   - Handles multi-line descriptions
   - Extracts amounts and balances with proper Dr/Cr notation
3. **Excel Export**:
   - Creates formatted Excel file with pandas/openpyxl
   - Applies consistent styling and formatting
   - Auto-sizes columns for optimal display

## Output Format

The generated Excel file contains these columns:

| Column | Format | Description |
|--------|--------|-------------|
| Date | DD-MMM-YYYY | Transaction date |
| Description | Text | Transaction description |
| Amount | Numeric (0.00) | Transaction amount |
| Type | Text | Credit/Debit indicator |
| Formatted Balance | Text | Account balance with Dr/Cr notation |

## Limitations

- Currently optimized for Punjab & Sind Bank statements
- Requires properly formatted PDFs (not scanned documents)
- May need adjustments for statements with different layouts

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
