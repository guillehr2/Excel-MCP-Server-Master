# Excel MCP Server - Structure 🏗️

## Project Structure

```
excel-mcp-server/
│
├── 📄 master_excel_mcp.py      # Main server file with all functionality
├── 📄 index.js                 # NPM wrapper for easy installation
├── 📄 package.json             # NPM package configuration
├── 📄 pyproject.toml           # Python project configuration
├── 📄 setup.py                 # Python setup script
├── 📄 requirements.txt         # Python dependencies
│
├── 📄 README.md                # Main documentation
├── 📄 LICENSE                  # MIT License
├── 📄 LICENSE.md               # MIT License (markdown)
├── 📄 CHANGELOG.md             # Version history
├── 📄 CONTRIBUTING.md          # Contribution guidelines
│
├── 📁 .github/                 # GitHub specific files
│   └── 📁 workflows/
│       └── 📄 ci-cd.yml        # GitHub Actions CI/CD
│
├── 📁 docs/                    # Documentation
│   ├── 📄 quick-start.md       # Getting started guide
│   ├── 📄 api-reference.md     # Complete API documentation
│   ├── 📄 examples.md          # Comprehensive examples
│   └── 📄 troubleshooting.md   # Problem resolution guide
│
├── 📁 tests/                   # Test suite
│   ├── 📄 __init__.py
│   └── 📄 test_basic_operations.py
│
├── 📁 examples/                # Example files
│   └── 📄 simple_report.md     # Basic usage example
│
├── 📁 assets/                  # Project assets
│   └── 📄 banner.svg           # Project banner
│
├── 📄 .gitignore               # Git ignore rules
├── 📄 .npmignore               # NPM ignore rules
├── 📄 .editorconfig            # Editor configuration
├── 📄 MANIFEST.in              # Python package manifest
├── 📄 mcp-config-example.json  # MCP configuration example
├── 📄 publish.sh               # Publishing script (Unix)
└── 📄 publish.bat              # Publishing script (Windows)
```

## Installation Methods

### 1. Using NPX (Recommended)
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

### 2. Global NPM Installation
```bash
npm install -g @guillehr2/excel-mcp-server
```

### 3. Python Installation
```bash
pip install excel-mcp-server
```

### 4. From Source
```bash
git clone https://github.com/guillehr2/Excel-MCP-Server-Master.git
cd Excel-MCP-Server-Master
npm install
pip install -r requirements.txt
```

## Publishing

### To NPM
```bash
# Unix/Linux/macOS
./publish.sh

# Windows
publish.bat
```

### To PyPI
```bash
python -m build
twine upload dist/*
```

## Features

- ✅ Single unified server file
- ✅ NPM package for easy distribution
- ✅ Automatic dependency installation
- ✅ Cross-platform support
- ✅ Comprehensive documentation
- ✅ Professional CI/CD setup
- ✅ Multiple installation methods
- ✅ Version management tools


## Important Notes

- The server automatically installs Python dependencies on first run
- Supports both `uv` and `pip` for dependency management
- Works with Node.js 14+ and Python 3.8+
- All Excel functionality is contained in `master_excel_mcp.py`

---


