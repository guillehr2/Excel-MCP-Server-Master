# Troubleshooting Guide - Excel MCP Server 🔧

This guide helps you resolve common issues when using the Excel MCP Server.

## Table of Contents

1. [Installation Issues](#installation-issues)
2. [Configuration Problems](#configuration-problems)
3. [Runtime Errors](#runtime-errors)
4. [Excel-Specific Issues](#excel-specific-issues)
5. [Performance Problems](#performance-problems)
6. [Common Error Messages](#common-error-messages)

## Installation Issues

### npm/npx fails to run the server

**Problem**: When running `npx @guillehr2/excel-mcp-server`, you get an error.

**Solutions**:

1. **Use specific version**:
   ```bash
   npx @guillehr2/excel-mcp-server@1.0.3
   ```

2. **Clear npm cache**:
   ```bash
   npm cache clean --force
   ```

3. **Check Node.js version**:
   ```bash
   node --version
   ```
   Ensure you have Node.js 14.0 or higher.

4. **Try global installation**:
   ```bash
   npm install -g @guillehr2/excel-mcp-server
   excel-mcp-server
   ```

### Python dependencies fail to install

**Problem**: The server fails to install Python dependencies on first run.

**Solutions**:

1. **Check Python version**:
   ```bash
   python --version
   # or
   python3 --version
   ```
   Ensure you have Python 3.8 or higher.

2. **Install dependencies manually**:
   ```bash
   pip install fastmcp openpyxl pandas numpy matplotlib xlsxwriter xlrd xlwt
   ```

3. **Use virtual environment**:
   ```bash
   python -m venv venv
   # On Windows:
   venv\Scripts\activate
   # On Unix/macOS:
   source venv/bin/activate
   
   pip install -r requirements.txt
   ```

## Configuration Problems

### MCP client doesn't recognize the server

**Problem**: After adding the configuration, the MCP client doesn't show Excel tools.

**Solutions**:

1. **Verify configuration path**:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. **Check JSON syntax**:
   ```json
   {
     "mcpServers": {
       "excel-master": {
         "command": "npx",
         "args": ["-y", "@guillehr2/excel-mcp-server"]
       }
     }
   }
   ```

3. **Restart MCP client**: Close and reopen your MCP client completely.

4. **Check logs**: Look for error messages in the MCP client logs.

### Permission denied errors

**Problem**: Server fails to start due to permission issues.

**Solutions**:

1. **Windows**: Run as administrator:
   ```bash
   # Right-click Command Prompt -> Run as administrator
   npx @guillehr2/excel-mcp-server
   ```

2. **macOS/Linux**: Use sudo cautiously:
   ```bash
   sudo npx @guillehr2/excel-mcp-server
   ```

3. **Fix npm permissions**:
   ```bash
   npm config set prefix ~/.npm-global
   export PATH=~/.npm-global/bin:$PATH
   ```

## Runtime Errors

### "Module not found" errors

**Problem**: Python modules are not found when running the server.

**Solutions**:

1. **Reinstall dependencies**:
   ```bash
   pip uninstall fastmcp openpyxl pandas numpy -y
   pip install fastmcp openpyxl pandas numpy
   ```

2. **Check Python path**:
   ```python
   import sys
   print(sys.path)
   ```

3. **Use absolute imports**: Ensure the server is using absolute imports.

### Server crashes on startup

**Problem**: The server starts but immediately crashes.

**Solutions**:

1. **Check for port conflicts**: Ensure no other service is using the same port.

2. **Verify file permissions**: Ensure the server can read/write to its directory.

3. **Run in debug mode**:
   ```bash
   node index.js --debug
   ```

## Excel-Specific Issues

### "File is locked" errors

**Problem**: Cannot modify Excel files because they're locked.

**Solutions**:

1. **Close Excel**: Ensure the file isn't open in Microsoft Excel.

2. **Check file permissions**:
   ```bash
   # Windows
   icacls "path\to\file.xlsx"
   
   # Unix/macOS
   ls -la path/to/file.xlsx
   ```

3. **Use a copy**: Work with a copy of the file:
   ```python
   import shutil
   shutil.copy2("original.xlsx", "working_copy.xlsx")
   ```

### Charts not displaying correctly

**Problem**: Charts are created but don't display as expected.

**Solutions**:

1. **Verify data range**: Ensure the data range contains valid numeric data.

2. **Check chart type compatibility**: Some data layouts work better with specific chart types.

3. **Use explicit ranges**:
   ```python
   # Good
   data_range = "Sheet1!A1:B10"
   
   # Avoid
   data_range = "A:B"
   ```

### Large files cause memory errors

**Problem**: Server runs out of memory with large Excel files.

**Solutions**:

1. **Increase Node.js memory**:
   ```bash
   node --max-old-space-size=4096 index.js
   ```

2. **Process in chunks**: Break large operations into smaller parts.

3. **Use read_only mode** for large files:
   ```python
   wb = openpyxl.load_workbook('large_file.xlsx', read_only=True)
   ```

## Performance Problems

### Slow file operations

**Problem**: Excel operations take too long.

**Solutions**:

1. **Disable automatic calculation**:
   ```python
   wb.calculation.calcMode = 'manual'
   ```

2. **Use write_only mode** for new files:
   ```python
   wb = openpyxl.Workbook(write_only=True)
   ```

3. **Batch operations**: Group multiple operations together.

### High CPU usage

**Problem**: Server uses excessive CPU resources.

**Solutions**:

1. **Limit concurrent operations**: Process files sequentially instead of in parallel.

2. **Add delays between operations**:
   ```python
   import time
   time.sleep(0.1)  # Small delay
   ```

3. **Use more efficient data structures**: Convert to pandas DataFrames for complex operations.

## Common Error Messages

### "ExcelMCPError: Workbook cannot be None"

**Cause**: Trying to operate on a workbook that hasn't been opened.

**Solution**: Ensure you open the workbook first:
```python
open_workbook_tool("file.xlsx")
```

### "SheetNotFoundError: Sheet 'X' does not exist"

**Cause**: Referencing a sheet that doesn't exist.

**Solution**: List sheets first:
```python
list_sheets_tool("file.xlsx")
```

### "RangeError: Invalid range"

**Cause**: Using an invalid cell range format.

**Solution**: Use proper Excel notation:
```python
# Correct
"A1:B10"
"Sheet1!A1:B10"

# Incorrect
"A-B"
"1:10"
```

### "TableError: A table named 'X' already exists"

**Cause**: Trying to create a table with a duplicate name.

**Solution**: Use unique table names or delete existing table first.

### "ChartError: The data range contains blank cells"

**Cause**: Chart data range includes empty cells.

**Solution**: Ensure data is continuous without gaps.

## Getting More Help

If you're still experiencing issues:

1. **Check the logs**: Enable verbose logging:
   ```bash
   export MCP_LOG_LEVEL=debug
   npx @guillehr2/excel-mcp-server
   ```

2. **Create a minimal example**: Isolate the problem with a simple test case.

3. **Report an issue**: Open an issue on GitHub with:
   - Error message
   - Steps to reproduce
   - System information
   - Sample files (if applicable)

4. **Community support**: Ask in the MCP community forums or Discord.

## Debug Mode

Run the server in debug mode for more information:

```bash
# Set environment variable
export EXCEL_MCP_DEBUG=true

# Run server
npx @guillehr2/excel-mcp-server
```

This will provide:
- Detailed error messages
- Stack traces
- File operation logs
- Memory usage information

---

For more information, see the [main documentation](../README.md) or [API reference](api-reference.md).
