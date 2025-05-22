# Changelog

All notable changes to the Excel MCP Server will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.3] - 2025-01-22

### Fixed
- Removed `--system` flag from `uv run` command (only kept for `uv pip install`)
- Fixed compatibility issues with `uv` dependency manager

## [1.0.2] - 2025-01-22

### Fixed
- Removed all console output that was interfering with MCP protocol
- Made dependency installation completely silent
- Changed output to stderr instead of stdout for better MCP compatibility
- Updated marker file to force reinstallation

## [1.0.1] - 2025-01-22

### Fixed
- Added `--system` flag to `uv` commands for proper system-wide installation
- Improved error handling for dependency installation

## [1.0.0] - 2025-01-22

### Added
- Initial release of Excel MCP Master Server
- Unified architecture with all functionality in `master_excel_mcp.py`
- Comprehensive workbook management tools (create, open, save, list sheets)
- Advanced data operations (write, update, read with formatting)
- Professional table creation with 50+ built-in styles
- Chart creation with 48 predefined styles and custom palettes
- Dashboard creation tools for complex visualizations
- Template-based report generation
- Import/export functionality (CSV, JSON, PDF)
- Automatic column width and row height adjustment
- Filter and data analysis tools
- Pivot table support
- Full MCP (Model Context Protocol) integration
- NPM package for easy installation with `npx`
- Comprehensive documentation and examples
- Cross-platform support (Windows, macOS, Linux)

### Security
- Input validation for all file operations
- Safe file path handling
- Protection against formula injection

### Technical Details
- Built with FastMCP for MCP protocol support
- Uses openpyxl for Excel manipulation
- Pandas integration for data processing
- Matplotlib support for advanced visualizations
- Automatic dependency installation on first run
- Support for both `uv` and `pip` package managers

## [0.9.0] - 2024-12-15 (Beta)

### Added
- Beta testing phase
- Core functionality implementation
- Basic documentation

### Changed
- Migrated from multiple modules to unified architecture
- Improved error handling

### Fixed
- Memory management for large files
- Chart positioning issues

## [0.1.0] - 2024-11-01 (Alpha)

### Added
- Initial proof of concept
- Basic Excel file manipulation
- MCP protocol integration

---

For detailed information about each release, see the [GitHub Releases](https://github.com/guillehr2/Excel-MCP-Server-Master/releases) page.

---

**Author**: Guillem Hermida ([GitHub](https://github.com/guillehr2))
