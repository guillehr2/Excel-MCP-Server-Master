# API Reference - Excel MCP Master Server 📚

Complete reference for all available tools in the Excel MCP Master Server.

## 🏗️ Architecture Overview

The Excel MCP Master Server is a unified server that combines all Excel manipulation functionality into a single file: `master_excel_mcp.py`. This approach simplifies deployment and reduces complexity while providing comprehensive Excel capabilities.

## 📂 Tool Categories

### 🗂️ Workbook Management Tools

#### `create_workbook_tool`
Creates a new Excel workbook.

**Parameters:**
- `filename` (str): Full path and name of the Excel file to create
- `overwrite` (bool, optional): Whether to overwrite existing file. Default: False

**Returns:**
```python
{
    "success": bool,
    "file_path": str,
    "message": str
}
```

**Example:**
```python
create_workbook_tool("reports/monthly_report.xlsx", overwrite=True)
```

#### `open_workbook_tool`
Opens an existing Excel workbook.

**Parameters:**
- `filename` (str): Full path to the Excel file

**Returns:**
```python
{
    "success": bool,
    "file_path": str,
    "sheets": list,
    "sheet_count": int,
    "message": str
}
```

#### `save_workbook_tool`
Saves the workbook to disk.

**Parameters:**
- `filename` (str): Path to the current workbook
- `new_filename` (str, optional): Save with a different name

**Returns:**
```python
{
    "success": bool,
    "original_file": str,
    "saved_file": str,
    "message": str
}
```

#### `list_sheets_tool`
Lists all worksheets in a workbook.

**Parameters:**
- `filename` (str): Path to the Excel file

**Returns:**
```python
{
    "success": bool,
    "file_path": str,
    "sheets": list,
    "count": int,
    "message": str
}
```

#### `add_sheet_tool`
Adds a new worksheet.

**Parameters:**
- `filename` (str): Path to the Excel file
- `sheet_name` (str): Name for the new sheet
- `index` (int, optional): Position to insert the sheet

**Returns:**
```python
{
    "success": bool,
    "file_path": str,
    "sheet_name": str,
    "sheet_index": int,
    "all_sheets": list,
    "message": str
}
```

#### `delete_sheet_tool`
Removes a worksheet.

**Parameters:**
- `filename` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet to delete

#### `rename_sheet_tool`
Renames a worksheet.

**Parameters:**
- `filename` (str): Path to the Excel file
- `old_name` (str): Current name of the sheet
- `new_name` (str): New name for the sheet

### ✏️ Data Operations Tools

#### `write_sheet_data_tool`
Writes a 2D array of data to a worksheet.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the target sheet
- `start_cell` (str): Starting cell (e.g., "A1")
- `data` (list): 2D array of data to write

**Returns:**
```python
{
    "success": bool,
    "file_path": str,
    "sheet_name": str,
    "start_cell": str,
    "rows_written": int,
    "columns_written": int,
    "message": str
}
```

**Example:**
```python
write_sheet_data_tool(
    "report.xlsx",
    "Sales",
    "B2",
    [
        ["Product", "Q1", "Q2", "Q3", "Q4"],
        ["Widget A", 1000, 1200, 1100, 1300],
        ["Widget B", 800, 900, 850, 950]
    ]
)
```

#### `update_cell_tool`
Updates a single cell value or formula.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the target sheet
- `cell` (str): Cell reference (e.g., "C5")
- `value_or_formula` (any): Value or formula (formulas start with "=")

**Example:**
```python
update_cell_tool("report.xlsx", "Sales", "D10", "=SUM(D2:D9)")
```

#### `create_sheet_with_data_tool`
Creates a new sheet with data in one operation.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name for the new sheet
- `data` (list): 2D array of data
- `overwrite` (bool, optional): Whether to overwrite if sheet exists

### 📊 Table and Formatting Tools

#### `add_table_tool`
Creates an Excel table from a data range.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet
- `table_name` (str): Unique name for the table
- `cell_range` (str): Range in Excel format (e.g., "A1:D10")
- `style` (str, optional): Table style name

**Available Table Styles:**
- `TableStyleLight1` through `TableStyleLight21`
- `TableStyleMedium1` through `TableStyleMedium28`
- `TableStyleDark1` through `TableStyleDark11`

**Example:**
```python
add_table_tool(
    "report.xlsx",
    "Sales",
    "SalesTable",
    "A1:D10",
    "TableStyleMedium9"
)
```

#### `create_formatted_table_tool`
Creates and formats a table in one step.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet
- `start_cell` (str): Starting cell
- `data` (list): 2D array of data
- `table_name` (str): Name for the table
- `table_style` (str, optional): Table style
- `formats` (dict, optional): Additional formatting options

**Format Options:**
```python
formats = {
    "B2:B10": "#,##0.00",  # Number format
    "C2:C10": "0.0%",      # Percentage format
    "A1:Z1": {             # Style format
        "bold": True,
        "fill_color": "4472C4",
        "font_color": "FFFFFF"
    }
}
```

### 📈 Chart and Visualization Tools

#### `add_chart_tool`
Creates a chart in the worksheet.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet
- `chart_type` (str): Type of chart
- `data_range` (str): Data range for the chart
- `title` (str, optional): Chart title
- `position` (str, optional): Position to place chart
- `style` (any, optional): Chart style
- `theme` (str, optional): Color theme
- `custom_palette` (list, optional): Custom color palette

**Chart Types:**
- `column` - Column chart
- `bar` - Bar chart  
- `line` - Line chart
- `pie` - Pie chart
- `scatter` - Scatter plot
- `area` - Area chart

**Chart Styles:**
- Numeric: 1-48
- Named: `light-1`, `dark-blue`, `colorful-1`, etc.

**Example:**
```python
add_chart_tool(
    "report.xlsx",
    "Sales", 
    "column",
    "A1:B10",
    title="Monthly Sales",
    position="D2",
    style="colorful-1"
)
```

#### `create_chart_from_data_tool`
Creates a chart from new data in one step.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet
- `data` (list): 2D array of data
- `chart_type` (str): Type of chart
- `position` (str, optional): Chart position
- `title` (str, optional): Chart title
- `style` (any, optional): Chart style

### 🏗️ Advanced Features

#### `create_dashboard_tool`
Creates a complete dashboard with multiple visualizations.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `data` (dict): Data organized by sheet
- `dashboard_config` (dict): Configuration for dashboard elements
- `overwrite` (bool, optional): Whether to overwrite existing file

**Dashboard Configuration:**
```python
dashboard_config = {
    "tables": [
        {
            "sheet": "Dashboard",
            "name": "DataTable",
            "range": "Data!A1:D10",
            "style": "TableStyleMedium9"
        }
    ],
    "charts": [
        {
            "sheet": "Dashboard",
            "type": "column",
            "data_range": "Data!A1:B10", 
            "title": "Sales Trend",
            "position": "F2",
            "style": "dark-blue"
        }
    ]
}
```

#### `create_report_from_template_tool`
Creates a report based on an Excel template.

**Parameters:**
- `template_file` (str): Path to template file
- `output_file` (str): Path for output file
- `data_mappings` (dict): Data to substitute in template
- `chart_mappings` (dict, optional): Chart updates
- `format_mappings` (dict, optional): Format updates

#### `update_report_tool`
Updates an existing report with new data.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `data_updates` (dict): New data by sheet and range
- `config_updates` (dict, optional): Configuration updates
- `recalculate` (bool, optional): Whether to recalculate formulas

### 🔄 Import/Export Tools

#### `import_data_tool`
Imports data from multiple sources.

**Parameters:**
- `excel_file` (str): Target Excel file
- `import_config` (dict): Import configuration
- `sheet_name` (str, optional): Default sheet name
- `start_cell` (str, optional): Default start cell
- `create_tables` (bool, optional): Whether to create Excel tables

**Import Configuration:**
```python
import_config = {
    "csv": [
        {
            "file_path": "data.csv",
            "sheet_name": "CSVData",
            "delimiter": ",",
            "encoding": "utf-8"
        }
    ],
    "json": [
        {
            "file_path": "data.json", 
            "sheet_name": "JSONData",
            "format": "records"
        }
    ]
}
```

#### `export_data_tool`
Exports Excel data to various formats.

**Parameters:**
- `excel_file` (str): Source Excel file
- `export_config` (dict): Export configuration

**Export Configuration:**
```python
export_config = {
    "csv": [
        {
            "sheet_name": "Sales",
            "range": "A1:D10",
            "output_file": "sales.csv",
            "delimiter": ","
        }
    ],
    "json": [
        {
            "sheet_name": "Products",
            "output_file": "products.json",
            "format": "records"
        }
    ]
}
```

#### `filter_data_tool`
Filters and extracts data from tables or ranges.

**Parameters:**
- `file_path` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet
- `range_str` (str, optional): Data range
- `table_name` (str, optional): Table name
- `filters` (dict, optional): Filter criteria

**Filter Examples:**
```python
filters = {
    "Region": ["North", "South"],     # Value in list
    "Sales": {"gt": 10000},          # Greater than
    "Date": {"lt": "2024-01-01"},    # Less than
    "Product": {"contains": "Widget"} # Contains text
}
```

#### `export_single_sheet_pdf_tool`
Exports a single-sheet workbook to PDF.

**Parameters:**
- `excel_file` (str): Path to Excel file
- `output_pdf` (str, optional): Output PDF path

#### `export_sheets_pdf_tool`
Exports specified sheets to PDF.

**Parameters:**
- `excel_file` (str): Path to Excel file
- `sheets` (list, optional): Sheet names to export
- `output_dir` (str, optional): Output directory
- `single_file` (bool, optional): Combine into single PDF

## 🎨 Formatting Reference

### Number Formats
- `"#,##0"` - Thousands separator
- `"#,##0.00"` - Two decimal places
- `"0.0%"` - Percentage
- `"mm/dd/yyyy"` - Date format
- `"$#,##0.00"` - Currency

### Style Properties
```python
style = {
    "bold": True,
    "italic": True,
    "font_name": "Arial",
    "font_size": 12,
    "font_color": "000000",
    "fill_color": "FFFF00",
    "alignment": "center",
    "border_style": "thin"
}
```

### Chart Style Names
- **Light themes**: `light-1` to `light-6`
- **Dark themes**: `dark-blue`, `dark-gray`, `dark-green`
- **Colorful themes**: `colorful-1` to `colorful-8`
- **Office themes**: `office-1` to `office-6`

## 🔧 Error Handling

All tools return a standardized response format:

**Success Response:**
```python
{
    "success": True,
    "message": "Operation completed successfully",
    # Additional data specific to the operation
}
```

**Error Response:**
```python
{
    "success": False,
    "error": "Error description",
    "message": "User-friendly error message"
}
```

## 📝 Best Practices

### File Paths
- Use absolute paths when possible
- Ensure directories exist before creating files
- Handle file permissions appropriately

### Data Formats
- Ensure data is properly structured as 2D arrays
- Handle None/null values appropriately
- Validate data types for numeric operations

### Performance
- Use bulk operations when possible
- Close workbooks when done (handled automatically)
- Consider memory usage for large datasets

### Error Prevention
- Check file existence before operations
- Validate sheet names and ranges
- Handle edge cases (empty data, invalid formats)

## 🚀 Advanced Usage Patterns

### Batch Processing
```python
# Process multiple files
files = ["report1.xlsx", "report2.xlsx", "report3.xlsx"]
for file in files:
    # Apply same operations to each file
    add_chart_tool(file, "Summary", "column", "A1:B10")
```

### Template-Based Reporting
```python
# Create standardized reports
template = "monthly_template.xlsx"
for month in ["Jan", "Feb", "Mar"]:
    create_report_from_template_tool(
        template,
        f"report_{month}.xlsx", 
        data_mappings={month: get_month_data(month)}
    )
```

### Complex Dashboards
```python
# Multi-step dashboard creation
create_dashboard_tool(file, raw_data, config)
update_report_tool(file, additional_data)
export_sheets_pdf_tool(file, ["Dashboard"])
```

---

For more examples and tutorials, see the [Quick Start Guide](quick-start.md) and [Examples](examples.md).