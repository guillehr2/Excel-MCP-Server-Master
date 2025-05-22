# Quick Start Guide - Excel MCP Master Server 🚀

This guide will help you get started with the Excel MCP Master Server quickly and efficiently.

## 📋 Prerequisites

Before you begin, ensure you have:
- Node.js 14.0 or higher
- Python 3.8 or higher
- A compatible MCP client (Claude Desktop, or other MCP-compatible applications)

## 🔧 Installation

### Recommended: Using NPX (No Installation Required)

The easiest way to use the Excel MCP Server is directly with `npx`:

```bash
# Test it works
npx @guillehr2/excel-mcp-server --version
```

### Alternative: Global Installation

Install globally for faster startup:

```bash
npm install -g @guillehr2/excel-mcp-server

# Verify installation
excel-mcp-server --version
```

### Development Setup

For development or customization:

```bash
git clone https://github.com/guillehr2/Excel-MCP-Server-Master.git
cd Excel-MCP-Server-Master
npm install
pip install -r requirements.txt
```

## ⚙️ Configuration

### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Linux:** `~/.config/Claude/claude_desktop_config.json`

#### Using NPX (Recommended)

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "npx",
      "args": [
        "-y",
        "@guillehr2/excel-mcp-server@latest"
      ]
    }
  }
}
```

#### Using Specific Version

For stability, you can pin to a specific version:

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "npx",
      "args": [
        "-y",
        "@guillehr2/excel-mcp-server@1.0.3"
      ]
    }
  }
}
```

#### Using Global Installation

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "excel-mcp-server"
    }
  }
}
```

#### Using Local Development

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "node",
      "args": ["C:/path/to/Excel-MCP-Server-Master/index.js"]
    }
  }
}
```

### Other MCP Clients

For other MCP clients, adapt the configuration format as needed. The key components are:
- **Command**: `npx` or `excel-mcp-server` or `node`
- **Arguments**: Package name or script path

## 🏃‍♂️ First Steps

### 1. Verify Installation

After configuring your MCP client:

1. Restart your MCP client completely
2. Look for Excel-related tools in the available tools list
3. You should see tools like:
   - `create_workbook_tool`
   - `write_sheet_data_tool`
   - `add_chart_tool`
   - `create_dashboard_tool`

### 2. Create Your First Excel File

Try this simple example:

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

# Save the file
save_workbook_tool("my_first_excel.xlsx")
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
        ["Product", "Sales", "Profit", "Margin %"],
        ["Product A", 10000, 2000, "=C2/B2"],
        ["Product B", 15000, 3500, "=C3/B3"],
        ["Product C", 8000, 1200, "=C4/B4"]
    ],
    table_name="SalesData",
    table_style="TableStyleMedium9",
    formats={
        "B2:C4": "#,##0",          # Number format
        "D2:D4": "0.0%",           # Percentage format
        "A1:D1": {                 # Header styling
            "bold": True,
            "fill_color": "4472C4",
            "font_color": "FFFFFF"
        }
    }
)
```

### 4. Add a Chart

Visualize your data:

```python
# Add a column chart
add_chart_tool(
    file_path="my_first_excel.xlsx",
    sheet_name="Sheet",
    chart_type="column",
    data_range="A1:B4",
    title="Sales by Product",
    position="F2",
    style="colorful-1"
)
```

## 🎯 Common Tasks

### Creating a Dashboard

```python
# Create a dashboard with multiple visualizations
create_dashboard_tool(
    file_path="dashboard.xlsx",
    data={
        "KPIs": [
            ["Metric", "Value", "Target", "Status"],
            ["Revenue", 125000, 120000, "Exceeded"],
            ["Customers", 1543, 1500, "Exceeded"],
            ["Satisfaction", 4.5, 4.0, "Exceeded"]
        ]
    },
    dashboard_config={
        "tables": [
            {
                "sheet": "Dashboard",
                "name": "KPITable",
                "range": "KPIs!A1:D4",
                "style": "TableStyleDark1"
            }
        ],
        "charts": [
            {
                "sheet": "Dashboard",
                "type": "column",
                "data_range": "KPIs!A1:B4",
                "title": "Performance Metrics",
                "position": "F2",
                "style": "colorful-3"
            }
        ]
    }
)
```

### Importing Data

```python
# Import CSV data
import_data_tool(
    excel_file="imported_data.xlsx",
    import_config={
        "csv": [
            {
                "file_path": "sales_data.csv",
                "sheet_name": "Sales",
                "encoding": "utf-8"
            }
        ]
    },
    create_tables=True
)
```

## 🔍 Troubleshooting

### Python Dependencies

On first run, the server automatically installs required Python packages. If this fails:

1. **Manual installation**:
   ```bash
   pip install fastmcp openpyxl pandas numpy matplotlib xlsxwriter xlrd xlwt
   ```

2. **Using virtual environment**:
   ```bash
   python -m venv venv
   # Windows: venv\Scripts\activate
   # Unix/macOS: source venv/bin/activate
   pip install -r requirements.txt
   ```

### Common Issues

#### Server Not Found
- Ensure Node.js and npm are installed: `node --version`
- Try clearing npm cache: `npm cache clean --force`
- Install globally: `npm install -g @guillehr2/excel-mcp-server`

#### MCP Client Doesn't Recognize Server
- Restart your MCP client completely
- Check the configuration file syntax (valid JSON)
- Verify the path in configuration matches your setup

#### Permission Errors
- Windows: Run terminal as Administrator
- Unix/macOS: Check file permissions
- Ensure write access to the directory

#### NPX Cache Issues
If npx is using an old version:
```bash
# Clear npx cache
rmdir /s /q %LOCALAPPDATA%\npm-cache\_npx

# Use specific version
npx @guillehr2/excel-mcp-server@1.0.3
```

### Debug Mode

Enable debug logging for more information:

```bash
# Windows
set EXCEL_MCP_DEBUG=true
npx @guillehr2/excel-mcp-server

# Unix/macOS
export EXCEL_MCP_DEBUG=true
npx @guillehr2/excel-mcp-server
```

## 🎓 Next Steps

1. **Explore Examples**: Check out the [examples collection](examples.md)
2. **API Reference**: Learn about all available tools in the [API reference](api-reference.md)
3. **Advanced Features**: Try creating dashboards and complex reports
4. **Customize**: Modify the server for your specific needs

## 📚 Resources

- [Full Documentation](../README.md)
- [API Reference](api-reference.md)
- [Examples](examples.md)
- [Troubleshooting Guide](troubleshooting.md)
- [GitHub Repository](https://github.com/guillehr2/Excel-MCP-Server-Master)
- [NPM Package](https://www.npmjs.com/package/@guillehr2/excel-mcp-server)

---

**You're now ready to create amazing Excel reports with AI! 🎉**

If you encounter any issues, please check the [troubleshooting guide](troubleshooting.md) or open an issue on GitHub.
