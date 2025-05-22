# Excel MCP Server - Quick Example

This example demonstrates how to use the Excel MCP Server to create a simple report.

## Prerequisites

Make sure you have configured the Excel MCP Server in your MCP client.

## Example: Creating a Sales Report

```python
# 1. Create a new Excel file
create_workbook_tool("sales_report_2024.xlsx", overwrite=True)

# 2. Add sales data
sales_data = [
    ["Month", "Product A", "Product B", "Product C", "Total"],
    ["January", 12500, 8900, 15600, "=SUM(B2:D2)"],
    ["February", 13200, 9200, 14800, "=SUM(B3:D3)"],
    ["March", 14100, 9500, 16200, "=SUM(B4:D4)"],
    ["Q1 Total", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)", "=SUM(E2:E4)"]
]

write_sheet_data_tool(
    file_path="sales_report_2024.xlsx",
    sheet_name="Sheet",
    start_cell="A1",
    data=sales_data
)

# 3. Create a formatted table
create_formatted_table_tool(
    file_path="sales_report_2024.xlsx",
    sheet_name="Sheet",
    start_cell="A1",
    data=sales_data,
    table_name="SalesTable",
    table_style="TableStyleMedium9",
    formats={
        "B2:E5": "#,##0",
        "A1:E1": {"bold": True, "fill_color": "366092", "font_color": "FFFFFF"}
    }
)

# 4. Add a chart
add_chart_tool(
    file_path="sales_report_2024.xlsx",
    sheet_name="Sheet",
    chart_type="column",
    data_range="A1:D4",
    title="Q1 Sales by Product",
    position="G2",
    style="colorful-1"
)

# 5. Save the file
save_workbook_tool("sales_report_2024.xlsx")
```

## Result

This will create a professional Excel report with:
- Formatted sales data table
- Automatic calculations for totals
- A colorful column chart
- Professional styling

The file will be saved as `sales_report_2024.xlsx` in your current directory.
