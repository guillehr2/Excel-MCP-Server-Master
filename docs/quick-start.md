# Quick Start Guide - MCP Excel Suite

This guide will help you quickly set up and start using MCP Excel Suite.

## Prerequisites

- Python 3.11 or higher
- UV package manager
- Microsoft Excel installed
- Basic understanding of MCP (Model Context Protocol)

## Installation

1. **Clone or download the repository**

```bash
git clone https://github.com/username/mcp-excel-suite.git
cd mcp-excel-suite
```

2. **Install dependencies**

```bash
# Using UV (recommended)
uv pip install matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt pywin32

# Alternative with pip
pip install matplotlib mcp[cli] numpy openpyxl pandas xlsxwriter xlrd xlwt pywin32
```

## Configuration

1. **Configure MCP**

Edit your MCP configuration file (typically found in `AppData\Roaming\Claude` on Windows or `~/.config/claude/config.json` on macOS/Linux) and add definitions for the Excel Suite MCP servers.

You can use the `mcp-config-example.json` file included in this repository as a reference, adjusting paths according to your environment.

2. **Verify the installation**

Ensure all paths in the configuration are correct, especially:

- Path to the UV executable
- Paths to the MCP server Python files

## Running the MCP Servers

To use MCP Excel Suite, you need to configure each MCP server in your preferred MCP manager:

1. **Configure the servers in your MCP interface**

Each MCP server is added separately. Make sure to include all the ones you need for your use case:

- **Excel Read Tools** (excel_mcp_complete.py)
- **Excel Edit Tools** (workbook_manager_mcp.py)
- **Excel Insert Tools** (excel_writer_mcp.py)
- **Excel Advanced Tools** (advanced_excel_mcp.py)
- **Excel Extra Tools** (excel_axtra_operations_mcp.py)

2. **Start interaction**

Once the MCP servers are configured, you can communicate with them using natural language through an MCP-compatible LLM.

## Quick Usage Examples

### Reading Excel Data

```
"Open sales.xlsx and show me a summary of the information"
"Read the first 10 records from customers.xlsx"
"What's the average of sales in column C?"
```

### Modifying an Excel Workbook

```
"Create a new workbook called budget.xlsx"
"Add a sheet called 'Income'"
"Format range A1:D10 as currency"
```

### Inserting Data

```
"Insert this data table into the active sheet: [data]"
"Add a row with the following values: Product A, 150, 10.99, In stock"
"Insert a header in A1 with the text 'Monthly Report'"
```

### Advanced Features

```
"Create a line chart with data from range A1:B20"
"Add data validation to make column F accept only dates"
"Set up conditional formatting in cells C5:C20 to highlight values greater than 1000"
```

### Working with Databases via DBHub

If you have [DBHub](https://github.com/bytebase/dbhub) configured as an MCP server, you can combine it with Excel Suite for powerful database integration:

```
"Query the PostgreSQL database for monthly sales data and create an Excel report"
"Execute this SQL query: 'SELECT * FROM customers WHERE region='Europe'' and save the results to a new Excel file"
"Create a bar chart comparing sales data from our MySQL database by quarter"
```

### Working with Text Files

```
"Import the CSV file quarterly_results.csv and format it as a table"
"Read data from customer_list.txt and organize it into columns"
"Parse the JSON file api_response.json and convert it to an Excel-friendly format"
```

## Data Source Integration Examples

### Database to Excel Workflow

```
"Connect to our PostgreSQL database, query the sales for Q1 2025, and create a formatted report with a summary chart"
```

This will:
1. Use DBHub to connect to your PostgreSQL database 
2. Execute the necessary SQL query
3. Import the data into Excel
4. Format it properly
5. Create a summary chart

### Text Files to Excel

```
"Import the CSV files from the /data directory, combine them into a single sheet, and create a pivot table showing monthly totals"
```

This will:
1. Read multiple CSV files
2. Combine their data
3. Create a consolidated Excel sheet
4. Generate a pivot table for analysis

## Next Steps

- Check out the [User Manual](user-manual.md) for more detailed information
- Review the [API Reference](api-reference.md) to learn about all available functionalities
- Explore the [Examples](examples.md) for more advanced use cases

## Troubleshooting

If you encounter issues during installation or execution:

1. Verify that all dependencies are correctly installed
2. Make sure the paths in your MCP configuration are correct
3. Check the MCP documentation for protocol-specific issues
4. Consult the [FAQ](faq.md) section for solutions to common problems
5. Look for error logs in the claude-mcp.log file (typically found in the same directory as your config file)

## Community Resources

- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [MCP GitHub Repository](https://github.com/modelcontextprotocol)
- [MCP Excel Server Examples](https://github.com/negokaz/excel-mcp-server)
- [DBHub MCP Server](https://github.com/bytebase/dbhub)
