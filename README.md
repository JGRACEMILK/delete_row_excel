```md
# Excel Row Deleter

A simple Python script to delete a specific row from an `.xlsx` file.

## Requirements

Install the dependency:

```bash
pip install openpyxl
```

## Usage

```bash
python delete_row.py <input.xlsx> <row_number> [--sheet SHEETNAME] [--output OUTPUT.xlsx]
```

- `input.xlsx` : Excel file to edit  
- `row_number` : Row number to delete (1-based)  
- `--sheet` : Sheet name (required if multiple sheets exist)  
- `--output` : Output file (optional, overwrites input if omitted)

## Examples

### Delete row 5 (single sheet file, overwrite)

```bash
python delete_row.py data.xlsx 5
```

### Delete row 10 from a specific sheet

```bash
python delete_row.py data.xlsx 10 --sheet Sheet2
```

### Delete row 7 and save to a new file

```bash
python delete_row.py data.xlsx 7 --output result.xlsx
```

### Delete row 3 when multiple sheets exist

```bash
python delete_row.py data.xlsx 3 --sheet "Sales 2025"
```

## Notes

- Row numbers are 1-based (same as Excel).
- If multiple sheets exist, `--sheet` is required.
- If `--output` is not provided, the input file is overwritten.
```