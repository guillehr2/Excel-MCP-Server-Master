# Excel MCP Server - Examples Collection 💡

This document provides comprehensive examples for using the Excel MCP Server in various scenarios.

## Table of Contents

1. [Basic Operations](#basic-operations)
2. [Data Manipulation](#data-manipulation)
3. [Formatting and Styling](#formatting-and-styling)
4. [Charts and Visualizations](#charts-and-visualizations)
5. [Advanced Dashboards](#advanced-dashboards)
6. [Template Reports](#template-reports)
7. [Data Import/Export](#data-importexport)
8. [Real-World Scenarios](#real-world-scenarios)

## Basic Operations

### Creating a Simple Spreadsheet

```python
# Create a new workbook
create_workbook_tool("my_data.xlsx", overwrite=True)

# Add a new sheet
add_sheet_tool("my_data.xlsx", "Sales Data")

# Write some data
write_sheet_data_tool(
    file_path="my_data.xlsx",
    sheet_name="Sales Data",
    start_cell="A1",
    data=[
        ["Date", "Product", "Quantity", "Price", "Total"],
        ["2024-01-01", "Widget A", 10, 25.99, "=C2*D2"],
        ["2024-01-02", "Widget B", 5, 35.50, "=C3*D3"],
        ["2024-01-03", "Widget A", 8, 25.99, "=C4*D4"]
    ]
)

# Save the workbook
save_workbook_tool("my_data.xlsx")
```

### Working with Multiple Sheets

```python
# Create workbook with multiple sheets
create_workbook_tool("multi_sheet.xlsx", overwrite=True)

# Add multiple sheets
sheets = ["January", "February", "March", "Q1 Summary"]
for sheet in sheets:
    add_sheet_tool("multi_sheet.xlsx", sheet)

# Delete the default sheet
delete_sheet_tool("multi_sheet.xlsx", "Sheet")

# Write data to each month
for month in ["January", "February", "March"]:
    write_sheet_data_tool(
        file_path="multi_sheet.xlsx",
        sheet_name=month,
        start_cell="A1",
        data=[
            ["Day", "Sales", "Expenses"],
            [1, 1000, 800],
            [2, 1200, 750],
            [3, 900, 820]
        ]
    )
```

## Data Manipulation

### Creating a Financial Summary

```python
# Create financial report
financial_data = [
    ["Account", "Q1", "Q2", "Q3", "Q4", "Total"],
    ["Revenue", 100000, 120000, 115000, 140000, "=SUM(B2:E2)"],
    ["Cost of Goods", 60000, 72000, 69000, 84000, "=SUM(B3:E3)"],
    ["Gross Profit", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=SUM(B4:E4)"],
    ["Operating Expenses", 20000, 22000, 21000, 25000, "=SUM(B5:E5)"],
    ["Net Income", "=B4-B5", "=C4-C5", "=D4-D5", "=E4-E5", "=SUM(B6:E6)"]
]

create_sheet_with_data_tool(
    file_path="financial_summary.xlsx",
    sheet_name="P&L Statement",
    data=financial_data,
    overwrite=True
)

# Add percentage calculations
update_cell_tool("financial_summary.xlsx", "P&L Statement", "G1", "% of Revenue")
update_cell_tool("financial_summary.xlsx", "P&L Statement", "G3", "=F3/F2")
update_cell_tool("financial_summary.xlsx", "P&L Statement", "G4", "=F4/F2")
update_cell_tool("financial_summary.xlsx", "P&L Statement", "G5", "=F5/F2")
update_cell_tool("financial_summary.xlsx", "P&L Statement", "G6", "=F6/F2")
```

### Dynamic Inventory Tracking

```python
# Create inventory tracker
inventory_data = [
    ["Item Code", "Product Name", "Current Stock", "Reorder Level", "Status"],
    ["P001", "Widget A", 150, 100, '=IF(C2<D2,"Reorder","OK")'],
    ["P002", "Widget B", 75, 100, '=IF(C3<D3,"Reorder","OK")'],
    ["P003", "Widget C", 200, 50, '=IF(C4<D4,"Reorder","OK")'],
    ["P004", "Widget D", 30, 75, '=IF(C5<D5,"Reorder","OK")']
]

create_formatted_table_tool(
    file_path="inventory.xlsx",
    sheet_name="Stock Levels",
    start_cell="A1",
    data=inventory_data,
    table_name="InventoryTable",
    table_style="TableStyleMedium4",
    formats={
        "E2:E5": {
            "bold": True
        }
    }
)
```

## Formatting and Styling

### Professional Invoice Template

```python
# Create an invoice
invoice_data = [
    ["INVOICE", "", "", "", ""],
    ["", "", "", "", ""],
    ["Invoice #:", "INV-2024-001", "", "Date:", "2024-01-15"],
    ["", "", "", "", ""],
    ["Bill To:", "", "", "", ""],
    ["Customer Name", "", "", "", ""],
    ["123 Main Street", "", "", "", ""],
    ["City, State 12345", "", "", "", ""],
    ["", "", "", "", ""],
    ["Item", "Description", "Qty", "Price", "Total"],
    ["001", "Professional Services", 10, 150.00, "=C10*D10"],
    ["002", "Consulting Hours", 5, 200.00, "=C11*D11"],
    ["003", "Support Package", 1, 500.00, "=C12*D12"],
    ["", "", "", "", ""],
    ["", "", "", "Subtotal:", "=SUM(E10:E12)"],
    ["", "", "", "Tax (10%):", "=E14*0.1"],
    ["", "", "", "Total:", "=E14+E15"]
]

# Create invoice with formatting
create_sheet_with_data_tool(
    file_path="invoice.xlsx",
    sheet_name="Invoice",
    data=invoice_data,
    overwrite=True
)

# Apply professional formatting
ws = open_workbook_tool("invoice.xlsx")

# Title formatting
apply_style_tool(
    file_path="invoice.xlsx",
    sheet_name="Invoice",
    cell_range="A1",
    style_dict={
        "font_size": 24,
        "bold": True,
        "font_color": "000080"
    }
)

# Header row formatting
apply_style_tool(
    file_path="invoice.xlsx",
    sheet_name="Invoice",
    cell_range="A9:E9",
    style_dict={
        "bold": True,
        "fill_color": "D3D3D3",
        "border_style": "thin"
    }
)

# Number formatting
apply_number_format_tool("invoice.xlsx", "Invoice", "D10:E16", "#,##0.00")
```

### Conditional Formatting Example

```python
# Sales performance tracker with visual indicators
performance_data = [
    ["Sales Rep", "Target", "Actual", "% Achieved", "Status"],
    ["John Smith", 50000, 55000, "=C2/B2", "=IF(D2>=1,\"✓\",\"✗\")"],
    ["Jane Doe", 60000, 45000, "=C3/B3", "=IF(D3>=1,\"✓\",\"✗\")"],
    ["Bob Johnson", 45000, 48000, "=C4/B4", "=IF(D4>=1,\"✓\",\"✗\")"],
    ["Alice Brown", 55000, 58000, "=C5/B5", "=IF(D5>=1,\"✓\",\"✗\")"]
]

create_formatted_table_tool(
    file_path="performance.xlsx",
    sheet_name="Sales Performance",
    start_cell="A1",
    data=performance_data,
    table_name="PerformanceTable",
    table_style="TableStyleLight10",
    formats={
        "B2:C5": "#,##0",
        "D2:D5": "0%",
        "A1:E1": {
            "bold": True,
            "fill_color": "4472C4",
            "font_color": "FFFFFF"
        }
    }
)
```

## Charts and Visualizations

### Sales Trend Analysis

```python
# Monthly sales data
sales_trend_data = [
    ["Month", "2023", "2024", "Growth"],
    ["Jan", 45000, 50000, "=(C2-B2)/B2"],
    ["Feb", 48000, 55000, "=(C3-B3)/B3"],
    ["Mar", 52000, 58000, "=(C4-B4)/B4"],
    ["Apr", 51000, 62000, "=(C5-B5)/B5"],
    ["May", 54000, 65000, "=(C6-B6)/B6"],
    ["Jun", 58000, 68000, "=(C7-B7)/B7"]
]

# Create data and format it
create_formatted_table_tool(
    file_path="sales_trend.xlsx",
    sheet_name="Sales Analysis",
    start_cell="A1",
    data=sales_trend_data,
    table_name="SalesTrend",
    table_style="TableStyleMedium2",
    formats={
        "B2:C7": "#,##0",
        "D2:D7": "0.0%"
    }
)

# Add comparison chart
add_chart_tool(
    file_path="sales_trend.xlsx",
    sheet_name="Sales Analysis",
    chart_type="column",
    data_range="A1:C7",
    title="Sales Comparison 2023 vs 2024",
    position="F2",
    style="colorful-1"
)

# Add growth trend line chart
add_chart_tool(
    file_path="sales_trend.xlsx",
    sheet_name="Sales Analysis",
    chart_type="line",
    data_range="A1:A7,D1:D7",
    title="Growth Trend",
    position="F18",
    style="dark-blue"
)
```

### Multi-Series Chart Example

```python
# Product category performance
category_data = [
    ["Quarter", "Electronics", "Clothing", "Home & Garden", "Sports"],
    ["Q1", 125000, 89000, 76000, 54000],
    ["Q2", 132000, 92000, 81000, 58000],
    ["Q3", 128000, 105000, 79000, 62000],
    ["Q4", 145000, 118000, 88000, 71000]
]

# Create and chart the data
result = create_chart_from_data_tool(
    file_path="category_performance.xlsx",
    sheet_name="Categories",
    data=category_data,
    chart_type="column",
    title="Quarterly Performance by Category",
    position="G2",
    style="colorful-3",
    create_table=True,
    table_name="CategoryData",
    table_style="TableStyleLight15"
)

# Add a pie chart for Q4 distribution
add_chart_tool(
    file_path="category_performance.xlsx",
    sheet_name="Categories",
    chart_type="pie",
    data_range="B1:E1,B5:E5",
    title="Q4 Revenue Distribution",
    position="G20",
    style="colorful-1"
)
```

## Advanced Dashboards

### Executive Dashboard

```python
# Create comprehensive executive dashboard
dashboard_data = {
    "KPIs": [
        ["Metric", "Current", "Previous", "Change"],
        ["Revenue", 2500000, 2300000, "=(B2-C2)/C2"],
        ["Customers", 15420, 14200, "=(B3-C3)/C3"],
        ["Avg Order Value", 162.3, 155.8, "=(B4-C4)/C4"],
        ["Customer Satisfaction", 4.6, 4.4, "=(B5-C5)/C5"]
    ],
    "Monthly": [
        ["Month", "Revenue", "Orders", "New Customers"],
        ["Jan", 200000, 1230, 210],
        ["Feb", 215000, 1380, 245],
        ["Mar", 208000, 1290, 198],
        ["Apr", 225000, 1420, 267],
        ["May", 232000, 1465, 289],
        ["Jun", 240000, 1510, 301]
    ],
    "Products": [
        ["Product", "Units Sold", "Revenue", "Margin %"],
        ["Premium Widget", 2500, 625000, 0.42],
        ["Standard Widget", 4200, 420000, 0.35],
        ["Basic Widget", 6800, 340000, 0.28],
        ["Accessories", 3200, 160000, 0.55]
    ]
}

# Create the dashboard
dashboard_result = create_dashboard_tool(
    file_path="executive_dashboard.xlsx",
    data=dashboard_data,
    dashboard_config={
        "tables": [
            {
                "sheet": "Dashboard",
                "name": "KPITable",
                "range": "KPIs!A1:D5",
                "style": "TableStyleDark2"
            },
            {
                "sheet": "Dashboard",
                "name": "MonthlyTable",
                "range": "Monthly!A1:D7",
                "style": "TableStyleMedium9"
            }
        ],
        "charts": [
            {
                "sheet": "Dashboard",
                "type": "line",
                "data_range": "Monthly!A1:B7",
                "title": "Revenue Trend",
                "position": "F2",
                "style": "dark-blue"
            },
            {
                "sheet": "Dashboard",
                "type": "column",
                "data_range": "Products!A1:C5",
                "title": "Product Performance",
                "position": "F18",
                "style": "colorful-2"
            },
            {
                "sheet": "Dashboard",
                "type": "pie",
                "data_range": "Products!A1:A5,C1:C5",
                "title": "Revenue by Product",
                "position": "M18",
                "style": "colorful-1"
            }
        ]
    }
)
```

### Interactive Sales Dashboard

```python
# Regional sales dashboard with multiple views
regional_data = {
    "Summary": [
        ["Region", "Q1", "Q2", "Q3", "Q4", "Total"],
        ["North", 250000, 275000, 265000, 290000, "=SUM(B2:E2)"],
        ["South", 180000, 195000, 205000, 215000, "=SUM(B3:E3)"],
        ["East", 220000, 240000, 235000, 260000, "=SUM(B4:E4)"],
        ["West", 195000, 210000, 225000, 245000, "=SUM(B5:E5)"],
        ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)", "=SUM(F2:F5)"]
    ],
    "Details": [
        ["Region", "Sales Rep", "Q1", "Q2", "Q3", "Q4"],
        ["North", "John Smith", 125000, 140000, 135000, 145000],
        ["North", "Jane Doe", 125000, 135000, 130000, 145000],
        ["South", "Bob Johnson", 90000, 95000, 100000, 105000],
        ["South", "Alice Brown", 90000, 100000, 105000, 110000],
        ["East", "Charlie Davis", 110000, 120000, 115000, 130000],
        ["East", "Diana Evans", 110000, 120000, 120000, 130000],
        ["West", "Frank Garcia", 95000, 105000, 110000, 120000],
        ["West", "Grace Harris", 100000, 105000, 115000, 125000]
    ]
}

# Create the regional dashboard
create_dashboard_tool(
    file_path="regional_dashboard.xlsx",
    data=regional_data,
    dashboard_config={
        "tables": [
            {
                "sheet": "Dashboard",
                "name": "RegionalSummary",
                "range": "Summary!A1:F6",
                "style": "TableStyleDark1"
            }
        ],
        "charts": [
            {
                "sheet": "Dashboard",
                "type": "column",
                "data_range": "Summary!A1:E5",
                "title": "Quarterly Performance by Region",
                "position": "A10",
                "style": "colorful-3"
            },
            {
                "sheet": "Dashboard",
                "type": "line",
                "data_range": "Summary!A1:F5",
                "title": "Regional Trends",
                "position": "I10",
                "style": "dark-blue"
            }
        ]
    }
)
```

## Template Reports

### Monthly Report Template

```python
# Create a reusable monthly report template
template_data = {
    "Summary": [
        ["Monthly Report - [MONTH] [YEAR]", "", "", ""],
        ["", "", "", ""],
        ["Key Metrics", "Actual", "Target", "Variance"],
        ["Revenue", "[REVENUE]", "[REVENUE_TARGET]", "=B4-C4"],
        ["Units Sold", "[UNITS]", "[UNITS_TARGET]", "=B5-C5"],
        ["New Customers", "[CUSTOMERS]", "[CUSTOMERS_TARGET]", "=B6-C6"],
        ["", "", "", ""],
        ["Performance", "This Month", "Last Month", "Change %"],
        ["Conversion Rate", "[CONV_RATE]", "[PREV_CONV_RATE]", "=(B9-C9)/C9"],
        ["Avg Deal Size", "[AVG_DEAL]", "[PREV_AVG_DEAL]", "=(B10-C10)/C10"]
    ]
}

# Create the template
create_report_from_template_tool(
    template_file="monthly_template.xlsx",
    output_file="january_2024_report.xlsx",
    data_mappings={
        "[MONTH]": "January",
        "[YEAR]": "2024",
        "[REVENUE]": 325000,
        "[REVENUE_TARGET]": 300000,
        "[UNITS]": 1543,
        "[UNITS_TARGET]": 1500,
        "[CUSTOMERS]": 287,
        "[CUSTOMERS_TARGET]": 250,
        "[CONV_RATE]": 0.032,
        "[PREV_CONV_RATE]": 0.028,
        "[AVG_DEAL]": 210.50,
        "[PREV_AVG_DEAL]": 195.25
    },
    format_mappings={
        "B4:D6": "#,##0",
        "B9:D10": "0.0%",
        "A1": {
            "font_size": 16,
            "bold": True
        }
    }
)
```

## Data Import/Export

### CSV Import with Processing

```python
# Import and process CSV data
import_result = import_data_tool(
    excel_file="processed_data.xlsx",
    import_config={
        "csv": [
            {
                "file_path": "raw_sales_data.csv",
                "sheet_name": "Raw Data",
                "delimiter": ",",
                "encoding": "utf-8",
                "start_cell": "A1"
            }
        ]
    },
    create_tables=True
)

# Add calculated columns
update_cell_tool("processed_data.xlsx", "Raw Data", "F1", "Total Value")
update_cell_tool("processed_data.xlsx", "Raw Data", "G1", "Commission")

# Add formulas (assuming quantity in D and price in E)
for row in range(2, 100):  # Adjust based on data size
    update_cell_tool(
        "processed_data.xlsx", 
        "Raw Data", 
        f"F{row}", 
        f"=D{row}*E{row}"
    )
    update_cell_tool(
        "processed_data.xlsx", 
        "Raw Data", 
        f"G{row}", 
        f"=F{row}*0.05"
    )
```

### Multi-Format Export

```python
# Export data to multiple formats
export_result = export_data_tool(
    excel_file="master_data.xlsx",
    export_config={
        "csv": [
            {
                "sheet_name": "Sales",
                "output_file": "sales_export.csv",
                "delimiter": ",",
                "include_headers": True
            }
        ],
        "json": [
            {
                "sheet_name": "Customers",
                "output_file": "customers.json",
                "format": "records",
                "orient": "records"
            }
        ],
        "pdf": [
            {
                "sheets": ["Summary", "Charts"],
                "output_file": "report.pdf",
                "orientation": "landscape"
            }
        ]
    }
)
```

## Real-World Scenarios

### Inventory Management System

```python
# Complete inventory management example
def create_inventory_system():
    # Create workbook
    create_workbook_tool("inventory_system.xlsx", overwrite=True)
    
    # Current inventory sheet
    inventory_data = [
        ["SKU", "Product", "Category", "Stock", "Min Stock", "Max Stock", "Status", "Reorder Qty"],
        ["SKU001", "Widget A", "Electronics", 150, 50, 500, '=IF(D2<E2,"REORDER",IF(D2>F2,"OVERSTOCK","OK"))', '=F2-D2'],
        ["SKU002", "Widget B", "Electronics", 45, 100, 400, '=IF(D3<E3,"REORDER",IF(D3>F3,"OVERSTOCK","OK"))', '=F3-D3'],
        ["SKU003", "Gadget X", "Accessories", 200, 75, 300, '=IF(D4<E4,"REORDER",IF(D4>F4,"OVERSTOCK","OK"))', '=F4-D4']
    ]
    
    # Create inventory sheet with formatting
    create_formatted_table_tool(
        file_path="inventory_system.xlsx",
        sheet_name="Current Stock",
        start_cell="A1",
        data=inventory_data,
        table_name="InventoryTable",
        table_style="TableStyleMedium7",
        formats={
            "D2:F10": "#,##0",
            "H2:H10": "#,##0",
            "G2:G10": {
                "bold": True
            }
        }
    )
    
    # Add order history sheet
    add_sheet_tool("inventory_system.xlsx", "Order History")
    
    order_history = [
        ["Order Date", "SKU", "Product", "Quantity", "Unit Cost", "Total Cost", "Supplier"],
        ["2024-01-15", "SKU001", "Widget A", 200, 15.50, "=D2*E2", "Supplier A"],
        ["2024-01-20", "SKU002", "Widget B", 150, 22.00, "=D3*E3", "Supplier B"],
        ["2024-02-01", "SKU003", "Gadget X", 100, 8.75, "=D4*E4", "Supplier A"]
    ]
    
    write_sheet_data_tool(
        file_path="inventory_system.xlsx",
        sheet_name="Order History",
        start_cell="A1",
        data=order_history
    )
    
    # Add dashboard
    add_sheet_tool("inventory_system.xlsx", "Dashboard")
    
    # Create inventory value chart
    add_chart_tool(
        file_path="inventory_system.xlsx",
        sheet_name="Dashboard",
        chart_type="column",
        data_range="'Current Stock'!B1:B4,D1:D4",
        title="Current Stock Levels",
        position="A2",
        style="colorful-2"
    )
    
    return "Inventory system created successfully!"

# Run the inventory system creation
create_inventory_system()
```

### Financial Analysis Workbook

```python
# Comprehensive financial analysis
def create_financial_analysis():
    # Financial data
    financial_data = {
        "Income Statement": [
            ["Income Statement", "", "", "", ""],
            ["For Year Ended December 31, 2024", "", "", "", ""],
            ["", "", "", "", ""],
            ["", "Q1", "Q2", "Q3", "Q4", "Total"],
            ["Revenue", 250000, 275000, 265000, 310000, "=SUM(B5:E5)"],
            ["Cost of Revenue", 150000, 165000, 159000, 186000, "=SUM(B6:E6)"],
            ["Gross Profit", "=B5-B6", "=C5-C6", "=D5-D6", "=E5-E6", "=SUM(B7:E7)"],
            ["", "", "", "", ""],
            ["Operating Expenses:", "", "", "", ""],
            ["Sales & Marketing", 40000, 44000, 42000, 49000, "=SUM(B10:E10)"],
            ["R&D", 30000, 32000, 31000, 35000, "=SUM(B11:E11)"],
            ["G&A", 20000, 22000, 21000, 24000, "=SUM(B12:E12)"],
            ["Total OpEx", "=SUM(B10:B12)", "=SUM(C10:C12)", "=SUM(D10:D12)", "=SUM(E10:E12)", "=SUM(B13:E13)"],
            ["", "", "", "", ""],
            ["Operating Income", "=B7-B13", "=C7-C13", "=D7-D13", "=E7-E13", "=SUM(B15:E15)"],
            ["", "", "", "", ""],
            ["Margin Analysis:", "", "", "", ""],
            ["Gross Margin %", "=B7/B5", "=C7/C5", "=D7/D5", "=E7/E5", "=F7/F5"],
            ["Operating Margin %", "=B15/B5", "=C15/C5", "=D15/D5", "=E15/E5", "=F15/F5"]
        ],
        "Balance Sheet": [
            ["Balance Sheet", "", ""],
            ["As of December 31, 2024", "", ""],
            ["", "", ""],
            ["Assets", "", "Amount"],
            ["Current Assets:", "", ""],
            ["Cash", "", 500000],
            ["Accounts Receivable", "", 250000],
            ["Inventory", "", 150000],
            ["Total Current Assets", "", "=SUM(C6:C8)"],
            ["", "", ""],
            ["Fixed Assets:", "", ""],
            ["Property & Equipment", "", 800000],
            ["Less: Depreciation", "", -200000],
            ["Net Fixed Assets", "", "=C12+C13"],
            ["", "", ""],
            ["Total Assets", "", "=C9+C14"],
            ["", "", ""],
            ["Liabilities & Equity", "", ""],
            ["Current Liabilities:", "", ""],
            ["Accounts Payable", "", 150000],
            ["Accrued Expenses", "", 50000],
            ["Total Current Liabilities", "", "=SUM(C20:C21)"],
            ["", "", ""],
            ["Long-term Debt", "", 300000],
            ["Total Liabilities", "", "=C22+C24"],
            ["", "", ""],
            ["Shareholders' Equity", "", "=C16-C25"],
            ["", "", ""],
            ["Total Liab. & Equity", "", "=C25+C27"]
        ]
    }
    
    # Create the workbook
    create_dashboard_tool(
        file_path="financial_analysis.xlsx",
        data=financial_data,
        dashboard_config={
            "tables": [
                {
                    "sheet": "Analysis",
                    "name": "IncomeTable",
                    "range": "'Income Statement'!A4:F15",
                    "style": "TableStyleMedium2"
                },
                {
                    "sheet": "Analysis",
                    "name": "BalanceTable",
                    "range": "'Balance Sheet'!A4:C29",
                    "style": "TableStyleMedium2"
                }
            ],
            "charts": [
                {
                    "sheet": "Analysis",
                    "type": "column",
                    "data_range": "'Income Statement'!A4:E7",
                    "title": "Quarterly Revenue vs Costs",
                    "position": "A2",
                    "style": "colorful-3"
                },
                {
                    "sheet": "Analysis",
                    "type": "line",
                    "data_range": "'Income Statement'!A17:E19",
                    "title": "Margin Trends",
                    "position": "H2",
                    "style": "dark-blue"
                }
            ]
        }
    )
    
    # Apply professional formatting
    apply_number_format_tool("financial_analysis.xlsx", "Income Statement", "B5:F15", "#,##0")
    apply_number_format_tool("financial_analysis.xlsx", "Income Statement", "B18:F19", "0.0%")
    apply_number_format_tool("financial_analysis.xlsx", "Balance Sheet", "C6:C29", "#,##0")
    
    return "Financial analysis workbook created!"

# Create the financial analysis
create_financial_analysis()
```

## Tips and Best Practices

### 1. Data Validation
Always validate your data before creating charts:
```python
# Check for empty cells in data range
# Ensure numeric columns don't contain text
# Verify date formats are consistent
```

### 2. Error Handling
Wrap operations in appropriate error handling:
```python
try:
    result = create_chart_tool(...)
    if result["success"]:
        print("Chart created successfully")
    else:
        print(f"Error: {result['error']}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

### 3. Performance Optimization
For large datasets:
- Use batch operations instead of individual cell updates
- Consider creating tables for better performance
- Use appropriate data types (avoid storing numbers as text)

### 4. Naming Conventions
- Use descriptive names for sheets, tables, and charts
- Avoid special characters in names
- Keep names under 31 characters (Excel limit)

### 5. Formula Best Practices
- Use absolute references ($A$1) when copying formulas
- Validate formula syntax before applying
- Consider using named ranges for complex formulas

---

For more information, see the [API Reference](api-reference.md) or [Quick Start Guide](quick-start.md).
