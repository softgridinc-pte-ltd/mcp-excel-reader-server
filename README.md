[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/softgridinc-pte-ltd-mcp-excel-reader-server-badge.png)](https://mseep.ai/app/softgridinc-pte-ltd-mcp-excel-reader-server)

# Excel Reader Server

A Model Context Protocol (MCP) server that provides tools for reading Excel (xlsx) files.

<a href="https://glama.ai/mcp/servers/kniyyx0gej">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/kniyyx0gej/badge" alt="Excel Reader Server MCP server" />
</a>

## Features

- Read content from all sheets in an Excel file
- Read content from a specific sheet by name
- Read content from a specific sheet by index
- Returns data in JSON format
- Handles empty cells and data type conversions

## Installation

Requires Python 3.10 or higher.

```bash
# Using pip
pip install excel-reader-server

# Using uv (recommended)
uv pip install excel-reader-server
```

## Dependencies

- mcp >= 1.2.1
- openpyxl >= 3.1.5

## Usage

The server provides three main tools:

### 1. read_excel

Reads content from all sheets in an Excel file.

```python
{
  "file_path": "path/to/your/excel/file.xlsx"
}
```

### 2. read_excel_by_sheet_name

Reads content from a specific sheet by name. If no sheet name is provided, reads the first sheet.

```python
{
  "file_path": "path/to/your/excel/file.xlsx",
  "sheet_name": "Sheet1"  # optional
}
```

### 3. read_excel_by_sheet_index

Reads content from a specific sheet by index. If no index is provided, reads the first sheet (index 0).

```python
{
  "file_path": "path/to/your/excel/file.xlsx",
  "sheet_index": 0  # optional
}
```

## Response Format

The server returns data in the following JSON format:

```json
{
  "Sheet1": [
    ["Header1", "Header2", "Header3"],
    ["Value1", "Value2", "Value3"],
    ["Value4", "Value5", "Value6"]
  ]
}
```

- Each sheet is represented as a key in the top-level object
- Sheet data is an array of arrays, where each inner array represents a row
- All values are converted to strings
- Empty cells are represented as empty strings

## Error Handling

The server provides clear error messages for common issues:
- File not found
- Invalid sheet name
- Index out of range
- General Excel file reading errors

## License

This project is released under the Apache 2 License. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
