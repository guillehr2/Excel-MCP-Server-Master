# Contributing to Excel MCP Master Server 🤝

Thank you for your interest in contributing to the Excel MCP Master Server! This document provides guidelines and information for contributors.

## 🌟 Ways to Contribute

### 🐛 Bug Reports
- Use the GitHub issue tracker
- Provide detailed reproduction steps
- Include system information (OS, Python version, dependencies)
- Attach sample files when relevant

### 💡 Feature Requests
- Describe the use case and benefits
- Provide examples of desired functionality
- Consider backward compatibility

### 📝 Code Contributions
- Fix bugs or implement new features
- Improve documentation
- Add tests for new functionality
- Optimize performance

### 📚 Documentation
- Improve existing documentation
- Add examples and tutorials
- Fix typos and clarify content

## 🔧 Development Setup

### Prerequisites
- Python 3.8 or higher
- Git
- Virtual environment (recommended)

### Setup Steps

1. **Fork and clone the repository:**
```bash
git clone https://github.com/guillehr2/Excel-MCP-Server-Master.git
cd Excel-MCP-Server-Master
```

2. **Create a virtual environment:**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install development dependencies:**
```bash
pip install -r requirements-dev.txt
```

4. **Install the package in development mode:**
```bash
pip install -e .
```

5. **Run tests to verify setup:**
```bash
pytest tests/
```

## 📁 Project Structure

The project follows a unified architecture with a single main file:

```
MCP_Server_Excel_Suite/
├── master_excel_mcp.py          # Main unified server file
├── README.md                    # Project documentation
├── LICENSE                      # License file
├── requirements.txt             # Dependencies
├── requirements-dev.txt         # Development dependencies
├── mcp-config-example.json      # Configuration example
├── docs/                        # Documentation
│   ├── quick-start.md
│   ├── api-reference.md
│   └── examples.md
├── tests/                       # Test suite
│   ├── test_basic_operations.py
│   ├── test_charts.py
│   └── test_advanced_features.py
└── assets/                      # Project assets
    └── banner.svg
```

## 🔄 Development Workflow

### 1. Create a Feature Branch
```bash
git checkout -b feature/your-feature-name
```

### 2. Make Changes
- Follow the existing code style
- Add docstrings for new functions
- Include type hints where appropriate
- Follow the naming conventions

### 3. Add Tests
- Write tests for new functionality
- Ensure existing tests still pass
- Aim for good test coverage

### 4. Update Documentation
- Update docstrings
- Add examples if relevant
- Update README.md if needed

### 5. Run Quality Checks
```bash
# Run tests
pytest tests/

# Check code style
black master_excel_mcp.py
flake8 master_excel_mcp.py

# Type checking
mypy master_excel_mcp.py
```

### 6. Commit and Push
```bash
git add .
git commit -m "feat: add new dashboard functionality"
git push origin feature/your-feature-name
```

### 7. Create Pull Request
- Use a descriptive title
- Include detailed description
- Reference related issues
- Add screenshots for UI changes

## 📝 Code Style Guidelines

### Python Code Style
- Follow PEP 8
- Use Black for formatting
- Maximum line length: 88 characters
- Use meaningful variable names
- Add type hints for function parameters and return values

### Function Documentation
```python
def create_chart(
    wb: Any,
    sheet_name: str,
    chart_type: str,
    data_range: str,
    title: Optional[str] = None
) -> Tuple[int, Any]:
    """
    Creates a chart in an Excel worksheet.
    
    Args:
        wb: Workbook object
        sheet_name: Name of the sheet
        chart_type: Type of chart ('column', 'bar', 'line', etc.)
        data_range: Range of data in A1:B5 format
        title: Optional chart title
        
    Returns:
        Tuple of (chart_id, chart_object)
        
    Raises:
        ChartError: If chart creation fails
        
    Example:
        chart_id, chart = create_chart(
            wb, "Sales", "column", "A1:B10", "Monthly Sales"
        )
    """
```

### Error Handling
- Use specific exception types
- Provide helpful error messages
- Include context information
- Log warnings appropriately

### MCP Tool Registration
```python
@mcp.tool(description="Clear, descriptive explanation of what the tool does")
def tool_name(param1: str, param2: int = 0) -> Dict[str, Any]:
    """
    Tool function with proper documentation.
    
    Args:
        param1: Description of parameter
        param2: Description with default value
        
    Returns:
        Dictionary with operation results
    """
    try:
        # Implementation
        return {
            "success": True,
            "message": "Operation completed successfully"
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "message": f"Error: {e}"
        }
```

## 🧪 Testing Guidelines

### Test Structure
- Use pytest framework
- Organize tests by functionality
- Use descriptive test names
- Include both positive and negative test cases

### Test Examples
```python
def test_create_workbook_success():
    """Test successful workbook creation."""
    result = create_workbook_tool("test.xlsx", overwrite=True)
    assert result["success"] is True
    assert os.path.exists("test.xlsx")

def test_create_workbook_file_exists():
    """Test workbook creation when file exists."""
    # Create file first
    create_workbook_tool("test.xlsx", overwrite=True)
    
    # Try to create again without overwrite
    result = create_workbook_tool("test.xlsx", overwrite=False)
    assert result["success"] is False
    assert "already exists" in result["error"]
```

### Running Tests
```bash
# Run all tests
pytest

# Run specific test file
pytest tests/test_basic_operations.py

# Run with coverage
pytest --cov=master_excel_mcp

# Run specific test
pytest tests/test_basic_operations.py::test_create_workbook_success
```

## 📋 Pull Request Guidelines

### Before Submitting
- [ ] Tests pass locally
- [ ] Code follows style guidelines
- [ ] Documentation is updated
- [ ] No merge conflicts with main branch
- [ ] Commit messages are clear

### PR Template
```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Documentation update
- [ ] Performance improvement
- [ ] Code refactoring

## Testing
- [ ] Tests added/updated
- [ ] All tests pass
- [ ] Manual testing completed

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No breaking changes (or documented)
```

## 🐛 Bug Report Template

```markdown
## Bug Description
Clear description of the bug

## To Reproduce
1. Step 1
2. Step 2
3. See error

## Expected Behavior
What should happen

## Screenshots
If applicable

## Environment
- OS: [Windows/macOS/Linux]
- Python version: [3.8/3.9/3.10/etc.]
- Dependencies versions: [openpyxl version, etc.]

## Additional Context
Any other relevant information
```

## 🎯 Feature Request Template

```markdown
## Feature Description
Clear description of the proposed feature

## Use Case
Why is this feature needed?

## Proposed Solution
How should this feature work?

## Alternatives Considered
Other approaches that were considered

## Additional Context
Mockups, examples, related issues
```

## 📞 Getting Help

If you need help with development:

1. Check existing documentation
2. Look at similar implementations in the codebase
3. Create a discussion in the GitHub repository
4. Reach out to maintainers

## 🏆 Recognition

Contributors will be:
- Listed in the project's contributors section
- Mentioned in release notes for significant contributions
- Eligible for maintainer status based on sustained contributions

Thank you for contributing to the Excel MCP Master Server! 🙏