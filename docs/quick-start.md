# Quick Start Guide - Excel MCP Master Server 🚀

This guide will help you get started with the Excel MCP Master Server quickly and efficiently.

## 📋 Prerequisites

Before you begin, ensure you have:
- Python 3.8 or higher
- Git (for cloning the repository)
- A compatible MCP client (Claude Desktop, or other MCP-compatible applications)

## 🔧 Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/MCP_Server_Excel_Suite.git
cd MCP_Server_Excel_Suite
```

### Step 2: Install Dependencies

```bash
pip install fastmcp openpyxl pandas numpy
```

### Optional Dependencies (for advanced features)

```bash
# For SQL data imports
pip install pyodbc sqlalchemy

# For additional data processing
pip install xlsxwriter

# For PDF export (Windows)
pip install pywin32
```

## ⚙️ Configuration

### MCP Client Setup

The Excel MCP Master Server uses a single unified server file: `master_excel_mcp.py`

#### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "python",
      "args": ["C:/path/to/your/master_excel_mcp.py"],
      "env": {
        "PYTHONPATH": "C:/path/to/your/project"
      }
    }
  }
}
```

#### Alternative MCP Client Configuration

For other MCP clients, use similar configuration adjusting the command format as needed:

```json
{
  "servers": [
    {
      "name": "excel-master",
      "command": ["python", "/path/to/master_excel_mcp.py"]
    }
  ]
}
```

## 🏃‍♂️ First Steps

### 1. Test the Connection

After configuring your MCP client, restart it and verify that the Excel MCP Master Server is loaded:

1. Open your MCP client
2. Look for available tools starting with `excel_` or `create_`, `add_`, etc.
3. You should see tools like:
   - `create_workbook_tool`
   - `write_sheet_data_tool`
   - `add_chart_tool`
   - `create_dashboard_tool`

### 2. Create Your First Excel File

Try this simple example to create your first Excel file:

```python
# Create a new Excel workbook
create_workbook_tool(
    filename="my_first_excel.xlsx",
    overwrite=True
)

# Add some data
write_sheet_data_tool(
    file_path="my_first_excel.xlsx",
    sheet_name="Sheet",
    start_cell="A1",
    data=[
        ["Product", "Sales", "Profit"],
        ["Product A", 10000, 2000],
        ["Product B", 15000, 3500],
        ["Product C", 8000, 1200]
    ]
)
```

### 3. Create a Professional Table

Enhance your data with professional formatting:

```python
# Create a formatted table
create_formatted_table_tool(
    file_path="my_first_excel.xlsx",
    sheet_name="Sheet",
    start_cell="A1",
    data=[
        ["Product", "Sales", "Profit"],
        ["Product A", 10000, 2000],
        ["Product B", 15000, 3500],
        ["Product C", 8000, 1200]
    ],
    table_name="SalesData",
    table_style="TableStyleMedium9",
    formats={
        "B2:B4": "#,##0",  # Number format for sales
        "C2:C4": "#,##0",  # Number format for profit
        "A1:C1": {"bold": True, "fill_color": "4472C4"}  # Header styling
    }
)
```

### 4. Add a Chart

Visualize your data with a chart:

```python
# Add a column chart
add_chart_tool(
    file_path="my_first_excel.xlsx",
    sheet_name="Sheet",
    chart_type="column",
    data_range="A1:B4",
    title="Sales by Product",
    position="E2",
    style="colorful-1"
)
```

## 🎯 Common Use Cases

### Creating Reports

```python
# Complete report creation in one step
create_sheet_with_data_tool(
    file_path="monthly_report.xlsx",
    sheet_name="March Report",
    data=[
        ["Metric", "Value", "Target", "Variance"],
        ["Revenue", 125000, 120000, 5000],
        ["Costs", 89000, 95000, -6000],
        ["Profit", 36000, 25000, 11000]
    ],
    overwrite=True
)
```

### Building Dashboards

```python
# Create a comprehensive dashboard
create_dashboard_tool(
    file_path="executive_dashboard.xlsx",
    data={
        "KPIs": [
            ["Metric", "Q1", "Q2", "Q3", "Q4"],
            ["Revenue", 100000, 120000, 115000, 140000],
            ["Customers", 1200, 1350, 1400, 1600],
            ["Satisfaction", 4.2, 4.3, 4.1, 4.5]
        ]
    },
    dashboard_config={
        "tables": [
            {
                "sheet": "Dashboard",
                "name": "KPITable",
                "range": "KPIs!A1:E4",
                "style": "TableStyleDark1"
            }
        ],
        "charts": [
            {
                "sheet": "Dashboard", 
                "type": "line",
                "data_range": "KPIs!A1:E2",
                "title": "Revenue Trend",
                "position": "A6",
                "style": "dark-blue"
            }
        ]
    }
)
```

### Importing External Data

```python
# Import data from CSV
import_data_tool(
    excel_file="analysis.xlsx",
    import_config={
        "csv": [
            {
                "file_path": "sales_data.csv",
                "sheet_name": "Sales",
                "start_cell": "A1",
                "encoding": "utf-8"
            }
        ]
    },
    create_tables=True
)
```

## 🔍 Troubleshooting

### Common Issues

#### 1. Server Not Loading
- Verify Python path is correct in configuration
- Check that all dependencies are installed
- Ensure the `master_excel_mcp.py` file path is correct

#### 2. Permission Errors
- Make sure you have write permissions to the target directory
- Close any Excel files that might be open
- Run your MCP client with appropriate permissions

#### 3. Import Errors
- Verify all required Python packages are installed:
  ```bash
  pip list | grep -E "(fastmcp|openpyxl|pandas|numpy)"
  ```

#### 4. File Not Found Errors
- Use absolute paths for file operations
- Ensure target directories exist
- Check file extensions (.xlsx, .csv, etc.)

### Getting Help

If you encounter issues:

1. Check the [main documentation](../README.md)
2. Look at the error messages in your MCP client's logs
3. Verify your configuration matches the examples above
4. Try with a simple example first before complex operations

## 🎓 Next Steps

Once you have the basic setup working:

1. **Explore Advanced Features**: Try creating dashboards, importing data from multiple sources
2. **Learn Template Usage**: Use `create_report_from_template_tool` for consistent reporting
3. **Automate Workflows**: Combine multiple tools for complex data processing
4. **Customize Styling**: Experiment with different chart styles and table formats

## 📚 Additional Resources

- [Full API Reference](api-reference.md)
- [Examples Collection](examples.md)
- [Troubleshooting Guide](troubleshooting.md)
- [Advanced Configuration](advanced-config.md)

---

**Ready to create amazing Excel reports! 🎉**