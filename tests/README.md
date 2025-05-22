# Excel MCP Server - Tests

This directory contains the test suite for the Excel MCP Server.

## Current Status

The test files currently contain placeholder tests. Real tests should be implemented to ensure the reliability of the server.

## Running Tests

```bash
# Install test dependencies
pip install pytest pytest-cov

# Run all tests
pytest

# Run with coverage
pytest --cov=master_excel_mcp

# Run specific test file
pytest tests/test_basic_operations.py
```

## Test Structure

Tests should be organized by functionality:

- `test_basic_operations.py` - Workbook creation, opening, saving
- `test_data_operations.py` - Data reading and writing
- `test_formatting.py` - Styling and formatting
- `test_charts.py` - Chart creation
- `test_advanced_features.py` - Dashboards, templates, etc.

## Writing Tests

Example test structure:

```python
import pytest
from master_excel_mcp import create_workbook_tool

def test_create_workbook_success():
    """Test successful workbook creation."""
    result = create_workbook_tool("test.xlsx", overwrite=True)
    assert result["success"] is True
    assert "file_path" in result
    
def test_create_workbook_exists():
    """Test workbook creation when file exists."""
    # First create
    create_workbook_tool("test.xlsx", overwrite=True)
    
    # Try again without overwrite
    result = create_workbook_tool("test.xlsx", overwrite=False)
    assert result["success"] is False
    assert "exists" in result["error"]
```

## TODO

- [ ] Implement real unit tests for all tools
- [ ] Add integration tests
- [ ] Add performance tests for large files
- [ ] Add edge case tests
- [ ] Set up continuous integration
