# Excel MCP Master Server 📊

[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen.svg)](https://modelcontextprotocol.io/)
[![npm version](https://img.shields.io/npm/v/@guillehr2/excel-mcp-server.svg)](https://www.npmjs.com/package/@guillehr2/excel-mcp-server)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/@guillehr2/excel-mcp-server.svg)](https://www.npmjs.com/package/@guillehr2/excel-mcp-server)
[![GitHub Stars](https://img.shields.io/github/stars/guillehr2/Excel-MCP-Server-Master.svg)](https://github.com/guillehr2/Excel-MCP-Server-Master/stargazers)
[![AI Powered](https://img.shields.io/badge/AI-Powered-purple.svg)](https://claude.ai)

A unified and comprehensive Model Context Protocol (MCP) server for complete Excel file manipulation. This single server provides all the functionality needed for reading, writing, formatting, and analyzing Excel files through LLM interactions.

![Excel MCP Master](./assets/banner.svg)

## 🌟 Features

### Unified Architecture
- **🎯 Single Server**: All functionality in one place - `master_excel_mcp.py`
- **📖 Complete Reading**: Data extraction, exploration, and analysis
- **✍️ Advanced Writing**: Professional formatting and styling
- **📋 Workbook Management**: Full lifecycle operations
- **📈 Rich Visualizations**: Charts, tables, pivot tables, and dashboards
- **🔄 Automation**: Templates, imports, exports, and batch operations
- **🎨 Professional Output**: Auto-formatting and styling for publication-ready documents

### Key Capabilities

#### 📊 Data Operations
- Read and write Excel files with full formatting support
- Create professional tables with automatic styling
- Generate charts and visualizations
- Import from CSV, JSON, and SQL sources
- Export to multiple formats (CSV, JSON, PDF)

#### 🎨 Professional Formatting
- Automatic column width adjustment
- Rich text formatting and styling
- Professional color schemes and themes
- Publication-ready document generation

#### 🏗️ Advanced Features
- Dynamic dashboards with multiple visualizations
- Template-based report generation
- Data filtering and analysis
- Pivot tables and advanced calculations
- Batch processing and automation

## 🚀 Quick Start

### Installation

The easiest way to use Excel MCP Server is with `npx` (no installation required):

```bash
npx @guillehr2/excel-mcp-server@latest
```

Or install globally:

```bash
npm install -g @guillehr2/excel-mcp-server
```

### Configuration

Add to your MCP client configuration (e.g., Claude Desktop):

#### Using npx (Recommended)

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

#### Using specific version

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "npx",
      "args": [
        "-y",
        "@guillehr2/excel-mcp-server@1.0.7"
      ]
    }
  }
}
```

#### Using global installation

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "excel-mcp-server"
    }
  }
}
```

#### Development mode

If you're developing or want to run from source:

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "node",
      "args": ["path/to/Excel-MCP-Server-Master/index.js"]
    }
  }
}
```

## 🛠️ Available Tools

### 📁 Workbook Management
- `create_workbook_tool` - Create new Excel files
- `open_workbook_tool` - Open existing files
- `save_workbook_tool` - Save workbooks
- `list_sheets_tool` - List all worksheets
- `add_sheet_tool` - Add new worksheets
- `delete_sheet_tool` - Remove worksheets
- `rename_sheet_tool` - Rename worksheets

### ✍️ Data Operations
- `write_sheet_data_tool` - Write data arrays
- `update_cell_tool` - Update individual cells
- `create_sheet_with_data_tool` - Create sheet with data in one step

### 📊 Tables and Formatting
- `add_table_tool` - Create professional Excel tables
- `create_formatted_table_tool` - Create and format tables in one step

### 📈 Charts and Visualizations
- `add_chart_tool` - Create various chart types
- `create_chart_from_data_tool` - Generate charts from new data

### 🏗️ Advanced Features
- `create_dashboard_tool` - Build dynamic dashboards
- `create_report_from_template_tool` - Template-based reports
- `update_report_tool` - Update existing reports
- `import_data_tool` - Import from multiple sources
- `export_data_tool` - Export to various formats
- `filter_data_tool` - Filter and analyze data
- `export_single_sheet_pdf_tool` - Export single sheet to PDF
- `export_sheets_pdf_tool` - Export multiple sheets to PDF

## 💡 Usage Examples

### Creating a Professional Report

```python
# Create a new workbook with formatted data
result = create_formatted_table_tool(
    file_path="sales_report.xlsx",
    sheet_name="Q4 Sales",
    start_cell="A1",
    data=[
        ["Region", "Q4 Sales", "Growth %"],
        ["North", 125000, 15.2],
        ["South", 98000, 8.7],
        ["East", 156000, 22.1],
        ["West", 89000, -3.2]
    ],
    table_name="Q4SalesData",
    table_style="TableStyleMedium9",
    formats={
        "B2:B5": "#,##0",  # Number format for sales
        "C2:C5": "0.0%",   # Percentage format
        "A1:C1": {"bold": True, "fill_color": "366092"}  # Header styling
    }
)

# Add a chart based on the table data
chart_result = add_chart_tool(
    file_path="sales_report.xlsx",
    sheet_name="Q4 Sales",
    chart_type="column",
    data_range="A1:B5",
    title="Q4 Sales by Region",
    position="E2",
    style="colorful-1"
)
```

### Building a Dynamic Dashboard

```python
# Create a comprehensive dashboard
dashboard_result = create_dashboard_tool(
    file_path="executive_dashboard.xlsx",
    data={
        "Data": [
            ["Month", "Revenue", "Expenses", "Profit"],
            ["Jan", 50000, 30000, 20000],
            ["Feb", 55000, 32000, 23000],
            ["Mar", 48000, 29000, 19000]
        ]
    },
    dashboard_config={
        "tables": [
            {
                "sheet": "Dashboard",
                "name": "MonthlyData",
                "range": "Data!A1:D4",
                "style": "TableStyleMedium9"
            }
        ],
        "charts": [
            {
                "sheet": "Dashboard",
                "type": "line",
                "data_range": "Data!A1:B4",
                "title": "Revenue Trend",
                "position": "E1",
                "style": "dark-blue"
            },
            {
                "sheet": "Dashboard",
                "type": "column",
                "data_range": "Data!A1:D4",
                "title": "Monthly Comparison",
                "position": "E15",
                "style": "colorful-2"
            }
        ]
    }
)
```

### Data Import and Analysis

```python
# Import data from multiple sources
import_result = import_data_tool(
    excel_file="analysis.xlsx",
    import_config={
        "csv": [
            {
                "file_path": "sales_data.csv",
                "sheet_name": "Sales",
                "delimiter": ",",
                "encoding": "utf-8"
            }
        ],
        "json": [
            {
                "file_path": "customer_data.json",
                "sheet_name": "Customers",
                "format": "records"
            }
        ]
    },
    create_tables=True
)

# Filter and analyze the imported data
filtered_data = filter_data_tool(
    file_path="analysis.xlsx",
    sheet_name="Sales",
    table_name="Table_Sales_1",
    filters={
        "Region": ["North", "South"],
        "Sales": {"gt": 10000}
    }
)
```

## 🎨 Professional Features

### Automatic Formatting
The server automatically applies professional formatting:
- **Column width adjustment** based on content length
- **Row height optimization** for wrapped text
- **Professional color schemes** for charts and tables
- **Consistent styling** throughout documents

### Chart Styling
Extensive chart customization options:
- **50+ predefined styles** (light, dark, colorful themes)
- **Custom color palettes** for brand consistency
- **Professional layouts** with proper spacing
- **Multiple chart types**: column, bar, line, pie, scatter, area

### Template System
Create reports from templates:
- **Reusable templates** for consistent reporting
- **Dynamic data substitution**
- **Automatic chart updates**
- **Format preservation**

## 📋 Requirements

- Node.js 14.0 or higher
- Python 3.8 or higher
- Operating System: Windows, macOS, or Linux

Python dependencies are automatically installed on first run:
- fastmcp
- openpyxl
- pandas
- numpy
- matplotlib
- xlsxwriter
- xlrd
- xlwt

## 📚 Documentation

For detailed documentation, see:
- [📖 Quick Start Guide](docs/quick-start.md)
- [🔧 API Reference](docs/api-reference.md)
- [💡 Examples](docs/examples.md)
- [📝 Changelog](CHANGELOG.md)

## 🤝 Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### Development Setup

```bash
# Clone the repository
git clone https://github.com/guillehr2/Excel-MCP-Server-Master.git
cd Excel-MCP-Server-Master

# Install dependencies
npm install
pip install -r requirements.txt

# Run in development mode
node index.js
```

## 🐛 Troubleshooting

### Common Issues

1. **Python not found**: Ensure Python 3.8+ is installed and in your PATH
2. **Dependencies fail to install**: Try running with administrator privileges
3. **MCP client doesn't recognize the server**: Restart your MCP client after configuration

For more help, see our [troubleshooting guide](docs/troubleshooting.md) or open an issue.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp)
- Excel manipulation powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Data processing with [pandas](https://pandas.pydata.org/)
- Published on [npm](https://www.npmjs.com/) for easy distribution

## 📊 Stats

![GitHub last commit](https://img.shields.io/github/last-commit/guillehr2/Excel-MCP-Server-Master)
![GitHub issues](https://img.shields.io/github/issues/guillehr2/Excel-MCP-Server-Master)
![GitHub pull requests](https://img.shields.io/github/issues-pr/guillehr2/Excel-MCP-Server-Master)
![npm bundle size](https://img.shields.io/bundlephobia/min/@guillehr2/excel-mcp-server)

---

**Made with ❤️ for the MCP ecosystem**

*If you find this project useful, please consider giving it a ⭐ on GitHub!

---

**Created by Guillem Hermida** | [GitHub](https://github.com/guillehr2) | Contact: qtmsuite@gmail.com*
