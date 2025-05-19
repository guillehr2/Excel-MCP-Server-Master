# 🚀 MCP Excel Suite - The Evolution of Excel Control with MCP

[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen.svg)](https://modelcontextprotocol.io/)
[![Python 3.11](https://img.shields.io/badge/Python-3.11-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![AI Powered](https://img.shields.io/badge/AI-Powered-purple.svg)](https://claude.ai)

![MCP Excel Suite Banner](./assets/banner.png)

## 📊 Empower Your Excel Interaction Through LLMs

**MCP Excel Suite** is a comprehensive collection of Model Context Protocol (MCP) servers specifically designed to enable language models to interact with Excel spreadsheets in a natural and intuitive way. This suite provides unprecedented control over Excel through natural language, allowing tools like Claude, GPT, and other LLMs to manipulate data, create visualizations, and automate complex tasks in Excel without writing code.

> "MCP Excel Suite transforms how LLMs interact with Excel, taking productivity to the next level"

## 🌟 Key Features

- **Natural Interaction**: Communicate with Excel using everyday language
- **5 Specialized MCP Servers**: Tools for reading, editing, inserting, advanced operations, and extra functionality
- **Seamless Integration**: Works with any MCP-compatible LLM
- **Simple Configuration**: Easy to install and run
- **Powerful Automation**: Perform complex Excel tasks with simple natural language instructions
- **Advanced Analytics**: Enable LLMs to perform detailed analysis of your Excel data
- **Database Integration**: Create Excel reports by pulling data from various databases via DBHub

## 🧩 Suite Components

MCP Excel Suite consists of 5 specialized MCP servers:

| MCP Server | Main File | Description |
|------------|-----------|-------------|
| **Excel Read Tools** | excel_mcp_complete.py | Excel reading and data query functionalities |
| **Excel Edit Tools** | workbook_manager_mcp.py | Tools for modifying Excel workbooks and sheets |
| **Excel Insert Tools** | excel_writer_mcp.py | Functions for writing and inserting data into Excel |
| **Excel Advanced Tools** | advanced_excel_mcp.py | Advanced capabilities like formulas, validation, charts, and complex functions |
| **Excel Extra Tools** | excel_axtra_operations_mcp.py | Additional tools and special functionalities |

## 🔍 What is Model Context Protocol (MCP)?

[Model Context Protocol (MCP)](https://modelcontextprotocol.io/) is an open standard that enables language models (LLMs) to access external tools through a unified interface. MCP facilitates structured and secure interactions between LLMs and external applications, services, and data sources.

The MCP architecture allows:
- Extending LLM capabilities beyond their knowledge cutoff
- Ensuring secure interactions with external systems
- Providing enriched context for more accurate responses
- Standardizing how LLMs access external tools

## 🔌 Integration with Other MCP Servers

MCP Excel Suite can be combined with other MCP servers to create powerful workflows:

### Database Integration with DBHub

[DBHub](https://github.com/bytebase/dbhub) is a universal database MCP server connecting to various database systems including PostgreSQL, MySQL, SQL Server, SQLite, MariaDB, and Oracle. By using DBHub alongside MCP Excel Suite, you can:

- Query databases and create Excel reports from the results
- Analyze database data within Excel using charts and pivot tables
- Set up automated workflows that extract data from databases and format it in Excel
- Compare and visualize data from multiple database sources in a single Excel workbook

Example workflow:
```
"Query customer data from our PostgreSQL database using DBHub and create a sales report in Excel with a monthly trend chart"
```

## 📋 Prerequisites

- Python 3.11 or higher
- UV package manager (modern alternative to pip)
- Excel installed (for local operations)
- Python dependencies: matplotlib, mcp[cli], numpy, openpyxl, pandas, xlsxwriter, xlrd, xlwt, pywin32 (for some functionalities)

## 💻 Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/username/mcp-excel-suite.git
cd mcp-excel-suite
```

### 2. Install the UV package manager

```bash
pip install uv
```

### 3. Install the necessary dependencies

```bash
uv pip install matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt pywin32
```

### 4. Configure the MCP servers

Add the following servers to your MCP configuration:

```json
{
  "Excel read Tools": {
    "command": "[PYTHON_PATH]\\uv.exe",
    "args": [
      "run",
      "--with", "matplotlib",
      "--with", "mcp[cli]",
      "--with", "numpy",
      "--with", "openpyxl",
      "--with", "pandas",
      "--with", "xlsxwriter",
      "--with", "xlrd",
      "--with", "xlwt",
      "mcp", "run",
      "[PROJECT_PATH]\\excel_mcp_complete.py"
    ]
  },
  "Excel edit Tools": {
    "command": "[PYTHON_PATH]\\uv.exe",
    "args": [
      "run",
      "--with", "matplotlib",
      "--with", "mcp[cli]",
      "--with", "numpy",
      "--with", "openpyxl",
      "--with", "pywin32",
      "--with", "pandas",
      "--with", "xlsxwriter",
      "--with", "xlrd",
      "--with", "xlwt",
      "mcp", "run",
      "[PROJECT_PATH]\\workbook_manager_mcp.py"
    ]
  },
  "Excel insert Tools": {
    "command": "[PYTHON_PATH]\\uv.exe",
    "args": [
      "run",
      "--with", "matplotlib",
      "--with", "mcp[cli]",
      "--with", "numpy",
      "--with", "openpyxl",
      "--with", "pandas",
      "--with", "xlsxwriter",
      "--with", "xlrd",
      "--with", "xlwt",
      "mcp", "run",
      "[PROJECT_PATH]\\excel_writer_mcp.py"
    ]
  },
  "Excel Advanced Tools": {
    "command": "[PYTHON_PATH]\\uv.exe",
    "args": [
      "run",
      "--with", "matplotlib",
      "--with", "mcp[cli]",
      "--with", "numpy",
      "--with", "openpyxl",
      "--with", "pandas",
      "--with", "xlsxwriter",
      "--with", "xlrd",
      "--with", "xlwt",
      "mcp", "run",
      "[PROJECT_PATH]\\advanced_excel_mcp.py"
    ]
  },
  "Excel Extra Tools": {
    "command": "[PYTHON_PATH]\\uv.exe",
    "args": [
      "run",
      "--with", "matplotlib",
      "--with", "mcp[cli]",
      "--with", "numpy",
      "--with", "openpyxl",
      "--with", "pandas",
      "--with", "xlsxwriter",
      "--with", "xlrd",
      "--with", "xlwt",
      "mcp", "run",
      "[PROJECT_PATH]\\excel_axtra_operations_mcp.py"
    ]
  }
}
```

Replace `[PYTHON_PATH]` and `[PROJECT_PATH]` with your specific paths.

## 🔧 Complete Installation Guide

### Windows Configuration

1. **Install Python 3.11**
   - Download from [python.org](https://www.python.org/downloads/)
   - Make sure to select "Add Python to PATH" during installation

2. **Install UV**
   - Open PowerShell as administrator
   ```powershell
   pip install uv
   ```

3. **Configure the Project**
   - Clone or download this repository
   - Navigate to the project directory
   ```powershell
   cd path\to\mcp-excel-suite
   ```

4. **Install Dependencies**
   ```powershell
   uv pip install matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt pywin32
   ```

5. **Configure MCP servers**
   - Edit your MCP configuration file (usually in `AppData\Roaming\Claude`)
   - Add each of the servers as shown in the previous section

### macOS/Linux Configuration

1. **Install Python 3.11**
   ```bash
   # macOS (using Homebrew)
   brew install python@3.11
   
   # Linux (Ubuntu/Debian)
   sudo apt update
   sudo apt install python3.11 python3.11-venv python3-pip
   ```

2. **Install UV**
   ```bash
   pip3 install uv
   ```

3. **Configure the Project**
   ```bash
   git clone https://github.com/username/mcp-excel-suite.git
   cd mcp-excel-suite
   ```

4. **Install Dependencies**
   ```bash
   uv pip install matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt
   ```

5. **Configure MCP servers**
   - Edit your MCP configuration file (usually in `~/.config/mcp/config.json`)
   - Add each of the servers as shown above, adapting the paths to your system

## 📚 Usage and Examples

Each MCP server is used independently and configured separately. Once an MCP server is running, you can communicate with it using natural language through a compatible LLM.

### Examples with Excel Read Tools

```
"Show me the first 5 records from sales.xlsx"
"Calculate the average of the 'Revenue' column in the 'Finance' sheet"
"How many rows and columns does my inventory.xlsx file have?"
```

### Examples with Excel Edit Tools

```
"Create a new Excel workbook called 'Budget2025.xlsx'"
"Add a new sheet called 'Monthly Expenses'"
"Change the format of cell A1:B10 to currency with two decimal places"
```

### Examples with Excel Insert Tools

```
"Insert this data into the 'Customers' sheet: [Data list]"
"Create a table in the range A1:F20 with list format"
"Add a row with the following values: [Values]"
```

### Examples with Excel Advanced Tools

```
"Create a bar chart using the data from range A1:B10"
"Set up data validation for column C to only accept dates"
"Add conditional formatting to the 'Sales' column to highlight values above 5000"
```

### Examples with Excel Extra Tools

```
"Save the current sheet as PDF"
"Apply auto filters to the data table"
"Create a pivot table summarizing sales by region"
```

### Examples with Database Integration

```
"Connect to my PostgreSQL database using DBHub and fetch today's sales data"
"Query the SQL Server database for customer information and create an Excel report"
"Use DBHub to execute this SQL: 'SELECT * FROM orders WHERE order_date > '2025-01-01'' and save results to Excel"
"Import CSV data from local file, then create a chart comparing it with the data from our database"
```

## 📊 Data Source Integration

MCP Excel Suite can work with various data sources:

### Text Files and CSV

Import data from text files directly into Excel:

```
"Import the data from 'quarterly_results.csv' and create a summary sheet"
"Read the text file 'customer_feedback.txt' and parse it into a structured table"
"Convert the data from 'legacy_system_export.txt' into Excel format with proper columns"
```

### Database Integration via DBHub

Using the [DBHub MCP server](https://github.com/bytebase/dbhub), you can:

- Connect to various database systems (PostgreSQL, MySQL, SQL Server, SQLite, etc.)
- Execute SQL queries and import results directly into Excel
- Create Excel reports based on database data
- Compare data from different database sources

```
"Query our PostgreSQL database for monthly sales and create a trend chart"
"Fetch customer data from SQL Server and generate a formatted report"
"Create a dashboard comparing data from our MySQL and Oracle databases"
```

### Web Data

Through web integration MCP servers:

```
"Fetch current stock prices and create a tracking spreadsheet"
"Get weather data for the next week and create a formatted table"
"Import exchange rates from the web and create a currency conversion tool"
```

## 🛠️ Detailed Capabilities by Server

### Excel Read Tools
- Reading complete workbooks or specific sheets
- Exploring structure (sheets, ranges, tables)
- Reading specific cells, ranges, or tables
- Getting file and sheet properties
- Basic data analysis (descriptive statistics)

### Excel Edit Tools
- Creating and saving workbooks
- Sheet management (add, delete, rename)
- Setting workbook and sheet properties
- Cell and range manipulation
- Format and style management

### Excel Insert Tools
- Writing data to cells and ranges
- Creating and modifying tables
- Adding rows and columns
- Updating individual cells
- Creating dropdown lists and validations

### Excel Advanced Tools
- Creating and managing charts
- Applying conditional formatting
- Setting up filters and sorting
- Creating advanced formulas and functions
- Creating and managing pivot tables

### Excel Extra Tools
- Converting files to PDF
- Setting page and print options
- Manipulating images and objects
- Advanced operations like search and replace
- Setting up groups and subtotals

## 🤖 LLM Integration

MCP Excel Suite is compatible with any language model that supports the MCP protocol, including:

- Claude by Anthropic
- GPT-4 by OpenAI (with MCP adapters)
- LLaMA and other open-source models (with MCP implementations)

Integration is done through the MCP standard, ensuring compatibility and ease of use.

## 📈 Use Cases

- **Data Analysis**: Allows analysts to interact with complex datasets through natural language requests
- **Report Automation**: Create and update regular reports with a simple command
- **Data Exploration**: Investigate and visualize data without advanced Excel knowledge
- **Data Preparation**: Clean, transform, and prepare data for further analysis
- **Interactive Visualization**: Create custom visualizations based on your data
- **Financial Auditing**: Review and analyze financial data with ease
- **Database Reporting**: Generate Excel reports directly from database queries
- **Multi-source Integration**: Combine data from different sources (databases, CSV files, web data) into unified reports

## 🔄 Updates and Maintenance

To keep MCP Excel Suite updated:

```bash
# Update the repository
git pull origin main

# Update dependencies
uv pip install --upgrade matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt pywin32
```

## 🤝 Contributing

Contributions are welcome. If you'd like to contribute:

1. Fork the repository
2. Create a new branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Commit your changes (`git commit -m 'Add some amazing feature'`)
5. Push your branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📞 Contact and Support

- **GitHub Issues**: For bug reports or feature requests
- **Email**: qtmsuite@gmai.com
- **Twitter**: [@MCPExcelSuite](https://twitter.com/MCPExcelSuite)

---

## 📊 Performance and Resources

MCP Excel Suite is optimized for efficient performance, but keep in mind the following recommended requirements:

- **RAM**: 4GB minimum, 8GB recommended for large datasets
- **CPU**: Dual-core processor or higher
- **Disk space**: 500MB for basic installation
- **Network**: Internet connection for updates and some functionalities

## 🔗 Useful Links

- [Official MCP Documentation](https://modelcontextprotocol.io/)
- [MCP GitHub Repository](https://github.com/modelcontextprotocol)
- [Quick Start Guide](https://github.com/username/mcp-excel-suite/wiki/quick-start)
- [API Reference Guide](https://github.com/username/mcp-excel-suite/wiki/api-reference)
- [Community and Forum](https://github.com/username/mcp-excel-suite/discussions)
- [DBHub MCP Server](https://github.com/bytebase/dbhub)
- [Excel MCP Server Examples](https://github.com/negokaz/excel-mcp-server)

---

<p align="center">
  <b>MCP Excel Suite</b> - Empowering Excel interaction through artificial intelligence<br>
  Created with ❤️ for the MCP community
</p>

<!-- Keywords for SEO -->
<!-- model context protocol, mcp, excel automation, ai excel, llm tools, excel api, natural language excel, claude excel, gpt excel, data analysis, excel mcp, anthropic tools, excel automation, python excel, dbhub, database integration, MCP servers, Excel MCP, Excel MCP server, claude MCP, Excel Claude MCP, MCP, Claude, anthropic, excel server, python server -->
