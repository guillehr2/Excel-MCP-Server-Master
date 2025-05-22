#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel MCP Master (Model Context Protocol for Excel)
--------------------------------------------------
Unified library to manipulate Excel files with advanced features:
- Combines all Excel MCP modules under a single interface
- Provides high level functions for common operations
- Optimizes the workflow when working with Excel

This module integrates:
- excel_mcp_complete.py: Data reading and exploration
- workbook_manager_mcp.py: Workbook and sheet management
- excel_writer_mcp.py: Cell writing and formatting
- advanced_excel_mcp.py: Tables, formulas, charts and pivot tables

Author: MCP Team
Version: 1.0

Usage guide for LLMs and agents
-------------------------------
All functions of this library are designed to be used by language models or
automatic tools that generate Excel files. To obtain the best results, follow
these context recommendations for each operation:

- **Apply styles at all times** so the resulting sheets look visually pleasant.
  Use the functions in this library to style cells, tables and charts.
- **Avoid element overlap**. Place charts in free cells and leave at least a
  couple of rows of separation from tables or blocks of text. Never place charts
  over text.
- **Automatically adjust column width**. After writing tables or datasets,
  check which cells contain long text and increase the width so everything is
  readable without breaking the layout.
- **Always seek the clearest and most organised layout**, separating sections
  and grouping related elements so that the final file is easy to understand.
- **Check the orientation of the data**. If tables are not obvious, explicitly
  indicate whether the categories are in rows or columns so the chart functions
  interpret them correctly.
"""

import os
import sys
import json
import logging
import tempfile
import time
from pathlib import Path
from typing import List, Dict, Union, Optional, Tuple, Any, Callable
import math

# Logging configuration
logger = logging.getLogger("excel_mcp_master")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stderr)
handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
logger.addHandler(handler)

# Import MCP
try:
    from mcp.server.fastmcp import FastMCP
    HAS_MCP = True
except ImportError:
    logger.warning("FastMCP could not be imported. MCP server features will be unavailable.")
    HAS_MCP = False

# Attempt to import required libraries
try:
    import pandas as pd
    import numpy as np
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.styles import (
        Font, PatternFill, Border, Side, Alignment, 
        NamedStyle, Protection, Color, colors
    )
    from openpyxl.chart import (
        BarChart, LineChart, PieChart, ScatterChart, AreaChart,
        Reference, Series
    )
    from openpyxl.worksheet.filters import AutoFilter
    from openpyxl.pivot.table import PivotTable, PivotField
    from openpyxl.pivot.cache import PivotCache
    HAS_OPENPYXL = True
except ImportError as e:
    logger.warning(f"Failed to import required libraries: {e}")
    logger.warning("Some functionality may be unavailable")
    HAS_OPENPYXL = False

# Import existing Excel MCP modules
# Note: In a real implementation we would import the functions from the existing modules
# However, for this example the key functions are reimplemented directly

# Base exception classes (unified)
class ExcelMCPError(Exception):
    """Base exception for all Excel MCP errors."""
    pass

class FileNotFoundError(ExcelMCPError):
    """Raised when an Excel file is not found."""
    pass

class FileExistsError(ExcelMCPError):
    """Raised when attempting to create a file that already exists."""
    pass

class SheetNotFoundError(ExcelMCPError):
    """Raised when a sheet is not found in the Excel file."""
    pass

class SheetExistsError(ExcelMCPError):
    """Raised when attempting to create a sheet that already exists."""
    pass

class CellReferenceError(ExcelMCPError):
    """Raised when there is an issue with a cell reference."""
    pass

class RangeError(ExcelMCPError):
    """Raised when there is an issue with a cell range."""
    pass

class TableError(ExcelMCPError):
    """Raised when there is an issue with an Excel table."""
    pass

class ChartError(ExcelMCPError):
    """Raised when there is an issue with a chart."""
    pass

class FormulaError(ExcelMCPError):
    """Raised when there is an issue with a formula."""
    pass

class PivotTableError(ExcelMCPError):
    """Raised when there is an issue with a pivot table."""
    pass

# Common utilities 
class ExcelRange:
    """Utility class for manipulating and converting Excel ranges.

    This class offers helpers to convert between Excel notation (A1:B5) and
    zero-based Python coordinates, as well as utilities for validating ranges.
    """
    
    @staticmethod
    def parse_cell_ref(cell_ref: str) -> Tuple[int, int]:
        """
        Convert an A1-style cell reference to zero-based (row, column) coordinates.
        
        Args:
            cell_ref: Cell reference in Excel format (e.g. "A1", "B5")
            
        Returns:
            Tuple (row, column) using zero-based indices
            
        Raises:
            ValueError: If the cell reference is not valid
        """
        if not cell_ref or not isinstance(cell_ref, str):
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        
        # Extract the column portion (letters)
        col_str = ''.join(c for c in cell_ref if c.isalpha())
        # Extract the row portion (numbers)
        row_str = ''.join(c for c in cell_ref if c.isdigit())
        
        if not col_str or not row_str:
            raise ValueError(f"Invalid cell format: {cell_ref}")
        
        # Convert column letters to an index (A->0, B->1, etc.)
        col_idx = 0
        for c in col_str.upper():
            col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
        col_idx -= 1  # Adjust to zero-based index

        # Convert row number to zero-based index
        row_idx = int(row_str) - 1
        
        return row_idx, col_idx
    
    @staticmethod
    def parse_range(range_str: str) -> Tuple[int, int, int, int]:
        """
        Convert a range in A1:B5 style to zero-based coordinates (row1, col1, row2, col2).
        
        Args:
            range_str: Range in Excel format (e.g. "A1:B5")
            
        Returns:
            Tuple (start_row, start_col, end_row, end_col) using zero-based indices
            
        Raises:
            ValueError: If the range is not valid
        """
        if not range_str or not isinstance(range_str, str):
            raise ValueError(f"Invalid range: {range_str}")
        
        # Handle ranges that include a sheet reference
        if '!' in range_str:
            parts = range_str.split('!')
            if len(parts) != 2:
                raise ValueError(f"Invalid range with sheet format: {range_str}")
            range_str = parts[1]  # Use only the range portion
        
        # Split the range into starting and ending cells
        if ':' in range_str:
            start_cell, end_cell = range_str.split(':')
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row, end_col = ExcelRange.parse_cell_ref(end_cell)
        else:
            # If only one cell, start and end are the same
            start_row, start_col = ExcelRange.parse_cell_ref(range_str)
            end_row, end_col = start_row, start_col
        
        return start_row, start_col, end_row, end_col
    
    @staticmethod
    def cell_to_a1(row: int, col: int) -> str:
        """
        Convert zero-based (row, column) coordinates to an A1 cell reference.
        
        Args:
            row: Row index (zero-based)
            col: Column index (zero-based)
            
        Returns:
            Cell reference in A1 format
        """
        if row < 0 or col < 0:
            raise ValueError(f"Negative indices not allowed: row={row}, column={col}")
        
        # Convert column index to letters
        col_str = ""
        col_val = col + 1  # Convert to base 1 for calculation
        
        while col_val > 0:
            remainder = (col_val - 1) % 26
            col_str = chr(65 + remainder) + col_str
            col_val = (col_val - 1) // 26
        
        # Convert row index to a 1-based Excel number
        row_val = row + 1
        
        return f"{col_str}{row_val}"
    
    @staticmethod
    def range_to_a1(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
        """
        Convert zero-based range coordinates to an A1:B5 style range.
        
        Args:
            start_row: Starting row (zero-based)
            start_col: Starting column (zero-based)
            end_row: Ending row (zero-based)
            end_col: Ending column (zero-based)
            
        Returns:
            Range in A1:B5 format
        """
        start_cell = ExcelRange.cell_to_a1(start_row, start_col)
        end_cell = ExcelRange.cell_to_a1(end_row, end_col)
        
        if start_cell == end_cell:
            return start_cell
        return f"{start_cell}:{end_cell}"

    @staticmethod
    def parse_range_with_sheet(range_str: str) -> Tuple[Optional[str], int, int, int, int]:
        """Convert a range that may include a sheet to a tuple ``(sheet, row1, col1, row2, col2)``.

        Args:
            range_str: Range string, possibly with sheet prefix ``Sheet!A1:B2``.

        Returns:
            Tuple ``(sheet, start_row, start_col, end_row, end_col)`` where ``sheet``
            is ``None`` if no sheet was specified.
        """
        if not range_str or not isinstance(range_str, str):
            raise ValueError(f"Invalid range: {range_str}")

        sheet_name = None
        pure_range = range_str
        if "!" in range_str:
            parts = range_str.split("!", 1)
            if len(parts) != 2:
                raise ValueError(f"Invalid range with sheet format: {range_str}")
            sheet_name = parts[0].strip("'")
            pure_range = parts[1]

        start_row, start_col, end_row, end_col = ExcelRange.parse_range(pure_range)
        return sheet_name, start_row, start_col, end_row, end_col

# Constants and mappings
# Map style names to Excel style numbers
CHART_STYLE_NAMES = {
    # Light styles
    'light-1': 1, 'light-2': 2, 'light-3': 3, 'light-4': 4, 'light-5': 5, 'light-6': 6,
    'office-1': 1, 'office-2': 2, 'office-3': 3, 'office-4': 4, 'office-5': 5, 'office-6': 6,
    'white': 1, 'minimal': 2, 'soft': 3, 'gradient': 4, 'muted': 5, 'outlined': 6,
    
    # Dark styles
    'dark-1': 7, 'dark-2': 8, 'dark-3': 9, 'dark-4': 10, 'dark-5': 11, 'dark-6': 12, 
    'dark-blue': 7, 'dark-gray': 8, 'dark-green': 9, 'dark-red': 10, 'dark-purple': 11, 'dark-orange': 12,
    'navy': 7, 'charcoal': 8, 'forest': 9, 'burgundy': 10, 'indigo': 11, 'rust': 12,
    
    # Colorful styles
    'colorful-1': 13, 'colorful-2': 14, 'colorful-3': 15, 'colorful-4': 16, 
    'colorful-5': 17, 'colorful-6': 18, 'colorful-7': 19, 'colorful-8': 20,
    'bright': 13, 'vivid': 14, 'rainbow': 15, 'multi': 16, 'contrast': 17, 'vibrant': 18,
    
    # Office themes
    'ion-1': 21, 'ion-2': 22, 'ion-3': 23, 'ion-4': 24,
    'wisp-1': 25, 'wisp-2': 26, 'wisp-3': 27, 'wisp-4': 28,
    'aspect-1': 29, 'aspect-2': 30, 'aspect-3': 31, 'aspect-4': 32,
    'badge-1': 33, 'badge-2': 34, 'badge-3': 35, 'badge-4': 36,
    'gallery-1': 37, 'gallery-2': 38, 'gallery-3': 39, 'gallery-4': 40,
    'median-1': 41, 'median-2': 42, 'median-3': 43, 'median-4': 44,
    
    # Styles for specific chart types
    'column-default': 1, 'column-dark': 7, 'column-colorful': 13, 
    'bar-default': 1, 'bar-dark': 7, 'bar-colorful': 13,
    'line-default': 1, 'line-dark': 7, 'line-markers': 3, 'line-dash': 5,
    'pie-default': 1, 'pie-dark': 7, 'pie-explosion': 4, 'pie-3d': 10,
    'area-default': 1, 'area-dark': 7, 'area-transparent': 5, 'area-stacked': 9,
    'scatter-default': 1, 'scatter-dark': 7, 'scatter-bubble': 4, 'scatter-smooth': 9,
}

# Mapping between styles and recommended color palettes
STYLE_TO_PALETTE = {
    # Light styles (1-6)
    1: 'office', 2: 'office', 3: 'colorful', 4: 'colorful', 5: 'pastel', 6: 'pastel',
    # Dark styles (7-12)
    7: 'dark-blue', 8: 'dark-gray', 9: 'dark-green', 10: 'dark-red', 11: 'dark-purple', 12: 'dark-orange',
    # Colorful styles (13-20)
    13: 'colorful', 14: 'colorful', 15: 'colorful', 16: 'colorful', 
    17: 'colorful', 18: 'colorful', 19: 'colorful', 20: 'colorful',
}

# CHART_COLOR_SCHEMES - normally defined in the original module
# Included here in simplified form
CHART_COLOR_SCHEMES = {
    'default': ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '8549BA', 'C55A11'],
    'colorful': ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47', '264478', '9E480E'],
    'pastel': ['9DC3E6', 'FFD966', 'C5E0B3', 'F4B183', 'B4A7D6', '8FBCDB', 'D89595', 'B7B7B7'],
    'dark-blue': ['2F5597', '1F3864', '4472C4', '5B9BD5', '8FAADC', '2E75B5', '255E91', '1C4587'],
    'dark-red': ['952213', 'C0504D', 'FF8B6B', 'EA6B66', 'DA3903', 'FF4500', 'B22222', '8B0000'],
    'dark-green': ['1E6C41', '375623', '548235', '70AD47', '9BC169', '006400', '228B22', '3CB371'],
    'dark-purple': ['5C3292', '7030A0', '8064A2', '9A7FBA', 'B3A2C7', '800080', '9400D3', '8B008B'],
    'dark-orange': ['C55A11', 'ED7D31', 'F4B183', 'FFC000', 'FFD966', 'FF8C00', 'FF7F50', 'FF4500']
}

# Helper function to obtain a worksheet (unified)
def get_sheet(wb, sheet_name_or_index) -> Any:
    """Retrieve a worksheet by name or index.

    Args:
        wb: Openpyxl workbook object.
        sheet_name_or_index: Sheet name or numeric index.

    Returns:
        Worksheet object.

    Raises:
        SheetNotFoundError: If the sheet does not exist.
    """
    if wb is None:
        raise ExcelMCPError("Workbook cannot be None")
    
    if isinstance(sheet_name_or_index, int):
        # If an index is provided, try to access by position
        if 0 <= sheet_name_or_index < len(wb.worksheets):
            return wb.worksheets[sheet_name_or_index]
        else:
            raise SheetNotFoundError(f"No sheet exists with index {sheet_name_or_index}")
    else:
        # If a name is provided, try to access by name
        try:
            return wb[sheet_name_or_index]
        except KeyError:
            sheets_info = ", ".join(wb.sheetnames)
            raise SheetNotFoundError(f"Sheet '{sheet_name_or_index}' does not exist in the file. Available sheets: {sheets_info}")
        except Exception as e:
            raise ExcelMCPError(f"Error accessing sheet: {e}")

# Helper function to convert a style specifier into a numeric style
def parse_chart_style(style):
    """
    Convert different style formats into a numeric Excel style (1-48).

    Args:
        style: Style as an int, numeric str, ``styleN`` or descriptive name.

    Returns:
        Integer style between 1 and 48, or ``None`` if not valid.
    """
    if isinstance(style, int) and 1 <= style <= 48:
        return style
        
    if isinstance(style, str):
        # Case 1: numeric string like '5'
        if style.isdigit():
            style_num = int(style)
            if 1 <= style_num <= 48:
                return style_num
                
        # Case 2: format 'styleN' or 'Style N'
        style_lower = style.lower()
        if style_lower.startswith('style'):
            try:
                # Extract the number following 'style'
                num_part = ''.join(c for c in style_lower[5:] if c.isdigit())
                if num_part:
                    style_num = int(num_part)
                    if 1 <= style_num <= 48:
                        return style_num
            except (ValueError, IndexError):
                pass
                
        # Case 3: descriptive name ('dark-blue', etc.)
        if style_lower in CHART_STYLE_NAMES:
            return CHART_STYLE_NAMES[style_lower]
            
    return None

# Helper to apply chart styles and colors simultaneously
def apply_chart_style(chart, style):
    """
    Apply a predefined style to a chart including the appropriate color palette.

    Args:
        chart: Openpyxl chart object.
        style: Style in any supported format (number, name, etc.).

    Returns:
        ``True`` if applied successfully, ``False`` otherwise.
    """
    # Convert to a style number if needed
    style_number = parse_chart_style(style)
    
    if style_number is None:
        style_str = str(style) if style else "None"
        logger.warning(f"Invalid chart style: '{style_str}'. Must be a number between 1-48 or a valid style name.")
        logger.info("Valid style names include: 'dark-blue', 'light-1', 'colorful-3', etc.")
        return False
        
    if not (1 <= style_number <= 48):
        logger.warning(f"Invalid chart style: {style_number}. It must be between 1 and 48.")
        return False
    
    # Step 1: apply the numeric style to native Excel attributes
    try:
        # The style property in openpyxl corresponds to the Excel style number
        chart.style = style_number
        logger.info(f"Applied native style {style_number} to chart")
    except Exception as e:
        logger.warning(f"Error applying style {style_number}: {e}")
    
    # Step 2: apply the color palette associated with the style's theme
    palette_name = STYLE_TO_PALETTE.get(style_number, 'default')
    colors = CHART_COLOR_SCHEMES.get(palette_name, CHART_COLOR_SCHEMES['default'])
    
    # Apply colors to the series
    try:
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.fill import ColorChoice
        
        for i, series in enumerate(chart.series):
            if i < len(colors):
                # Ensure graphical properties exist
                if not hasattr(series, 'graphicalProperties') or series.graphicalProperties is None:
                    series.graphicalProperties = GraphicalProperties()
                    
                # Assign color using ColorChoice for better compatibility
                color = colors[i % len(colors)]
                if isinstance(color, str) and color.startswith('#'):
                    color = color[1:]
                    
                series.graphicalProperties.solidFill = ColorChoice(srgbClr=color)
                
        logger.info(f"Applied style {style_number} with palette '{palette_name}' to chart")
        return True
        
    except Exception as e:
        logger.warning(f"Error applying colors for style {style_number}: {e}")
        return False

def determine_orientation(ws: Any, min_row: int, min_col: int, max_row: int, max_col: int) -> bool:
    """Attempt to guess the orientation of the data.

    Returns ``True`` if categories appear to be in the first column (column
    oriented) and ``False`` if they are more likely in the first row. The
    algorithm compares the ratio of numeric values for both interpretations and
    uses the shape of the range as a tiebreaker. This helps a language model
    avoid choosing wrong headers in ambiguous tables.
    """

    def _is_number(value: Any) -> bool:
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    # Calculate ratio of numbers assuming categories are in the first column
    col_numeric = col_total = 0
    for c in range(min_col + 1, max_col + 1):
        for r in range(min_row, max_row + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                col_total += 1
                if _is_number(val):
                    col_numeric += 1

    col_ratio = (col_numeric / col_total) if col_total else 0

    # Calculate ratio of numbers assuming categories are in the first row
    row_numeric = row_total = 0
    for r in range(min_row + 1, max_row + 1):
        for c in range(min_col, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                row_total += 1
                if _is_number(val):
                    row_numeric += 1

    row_ratio = (row_numeric / row_total) if row_total else 0

    if row_ratio > col_ratio:
        return False  # headers in the first row
    if col_ratio > row_ratio:
        return True   # headers in the first column

    # Tiebreaker based on range shape
    return (max_row - min_row) >= (max_col - min_col)

def _trim_range_to_data(ws: Any, min_row: int, min_col: int, max_row: int, max_col: int) -> Tuple[int, int, int, int]:
    """Remove trailing empty rows and columns from a range."""
    while max_row >= min_row:
        if all(ws.cell(row=max_row, column=c).value in (None, "") for c in range(min_col, max_col + 1)):
            max_row -= 1
        else:
            break
    while max_col >= min_col:
        if all(ws.cell(row=r, column=max_col).value in (None, "") for r in range(min_row, max_row + 1)):
            max_col -= 1
        else:
            break
    return min_row, min_col, max_row, max_col

def _range_has_blank(ws: Any, min_row: int, min_col: int, max_row: int, max_col: int) -> bool:
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            if ws.cell(row=r, column=c).value in (None, ""):
                return True
    return False

# ----------------------------------------
# BASE FUNCTIONS
# ----------------------------------------

# 1. Workbook management
def create_workbook(filename: str, overwrite: bool = False) -> Any:
    """
    Create a new empty Excel file.

    Args:
        filename (str): Full path and name of the file to create.
        overwrite (bool, optional): Overwrite existing file if ``True``.

    Returns:
        Workbook object.

    Raises:
        FileExistsError: If the file exists and ``overwrite`` is ``False``.
    """
    if os.path.exists(filename) and not overwrite:
        raise FileExistsError(f"El archivo '{filename}' ya existe. Use overwrite=True para sobreescribir.")
    
    wb = openpyxl.Workbook()
    if overwrite:
        save_workbook(wb, filename)
    return wb

def open_workbook(filename: str) -> Any:
    """
    Open an existing Excel file.

    Args:
        filename (str): Path to the file.

    Returns:
        Workbook object.

    Raises:
        FileNotFoundError: If the file does not exist.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"El archivo '{filename}' no existe.")
    
    try:
        wb = openpyxl.load_workbook(filename)
        return wb
    except Exception as e:
        logger.error(f"Error opening file '{filename}': {e}")
        raise ExcelMCPError(f"Error opening file: {e}")

def save_workbook(wb: Any, filename: Optional[str] = None) -> str:
    """
    Save the workbook to disk.

    Args:
        wb: Workbook object.
        filename (str, optional): Alternative file name if provided.

    Returns:
        Path to the saved file.

    Raises:
        ExcelMCPError: If an error occurs while saving.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    try:
        # Si no se proporciona filename, usar el filename original si existe
        if not filename and hasattr(wb, 'path'):
            filename = wb.path
        elif not filename:
            raise ExcelMCPError("Debe proporcionar un nombre de archivo")
        
        wb.save(filename)
        return filename
    except Exception as e:
        logger.error(f"Error saving workbook to '{filename}': {e}")
        raise ExcelMCPError(f"Error saving workbook: {e}")

def close_workbook(wb: Any) -> None:
    """
    Close the workbook in memory.

    Args:
        wb: Workbook object.

    Returns:
        None.
    """
    if not wb:
        return
    
    try:
        # Openpyxl does not really have a close() method,
        # but we can remove references to help the GC
        if hasattr(wb, "_archive"):
            wb._archive.close()
    except Exception as e:
        logger.warning(f"Warning while closing workbook: {e}")

def list_sheets(wb: Any) -> List[str]:
    """
    Return a list of sheet names.

    Args:
        wb: Workbook object.

    Returns:
        List[str]: Names of the sheets.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    if hasattr(wb, 'sheetnames'):
        return wb.sheetnames
    
    # Alternative if sheetnames cannot be accessed
    sheet_names = []
    for sheet in wb.worksheets:
        if hasattr(sheet, 'title'):
            sheet_names.append(sheet.title)
    
    return sheet_names

def add_sheet(wb: Any, sheet_name: str, index: Optional[int] = None) -> Any:
    """
    Add a new empty worksheet.

    Args:
        wb: Workbook object.
        sheet_name (str): Name of the sheet.
        index (int, optional): Position in the tab list.

    Returns:
        The created worksheet.

    Raises:
        SheetExistsError: If a sheet with that name already exists.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    # Check if a sheet with that name already exists
    if sheet_name in list_sheets(wb):
        raise SheetExistsError(f"A sheet named '{sheet_name}' already exists")
    
    # Create new sheet
    if index is not None:
        ws = wb.create_sheet(sheet_name, index)
    else:
        ws = wb.create_sheet(sheet_name)
    
    return ws

def delete_sheet(wb: Any, sheet_name: str) -> None:
    """
    Delete the specified worksheet.

    Args:
        wb: Workbook object.
        sheet_name (str): Name of the sheet to remove.

    Raises:
        SheetNotFoundError: If the sheet does not exist.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    # Check that the sheet exists
    if sheet_name not in list_sheets(wb):
        raise SheetNotFoundError(f"Sheet '{sheet_name}' does not exist in the workbook")
    
    # Delete the sheet
    try:
        del wb[sheet_name]
    except Exception as e:
        logger.error(f"Error deleting sheet '{sheet_name}': {e}")
        raise ExcelMCPError(f"Error deleting sheet: {e}")

def rename_sheet(wb: Any, old_name: str, new_name: str) -> None:
    """
    Rename a worksheet.

    Args:
        wb: Workbook object.
        old_name (str): Current sheet name.
        new_name (str): New sheet name.

    Raises:
        SheetNotFoundError: If the original sheet does not exist.
        SheetExistsError: If a sheet with the new name already exists.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    # Check that the original sheet exists
    if old_name not in list_sheets(wb):
        raise SheetNotFoundError(f"Sheet '{old_name}' does not exist in the workbook")
    
    # Check that no sheet with the new name exists
    if new_name in list_sheets(wb) and old_name != new_name:
        raise SheetExistsError(f"A sheet named '{new_name}' already exists")
    
    # Rename the sheet
    try:
        wb[old_name].title = new_name
    except Exception as e:
        logger.error(f"Error renaming sheet '{old_name}' to '{new_name}': {e}")
        raise ExcelMCPError(f"Error renaming sheet: {e}")

# 2. Data reading and exploration
def read_sheet_data(wb: Any, sheet_name: str, range_str: Optional[str] = None,
                   formulas: bool = False) -> List[List[Any]]:
    """
    Read values and optionally formulas from an Excel sheet.
    
    Args:
        wb: Openpyxl workbook object
        sheet_name: Sheet name
        range_str: Range in ``A1:B5`` format or ``None`` for the whole sheet
        formulas: If ``True`` return formulas instead of calculated values
    
    Returns:
        List of lists with cell values or formulas
        
    Raises:
        SheetNotFoundError: If the sheet does not exist
        RangeError: If the range is invalid
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # If no range is specified, use the whole sheet
    if not range_str:
        # Determine the used range (min_row, min_col, max_row, max_col)
        min_row, min_col = 1, 1
        max_row = ws.max_row
        max_col = ws.max_column
    else:
        # Parse the specified range
        try:
            min_row, min_col, max_row, max_col = ExcelRange.parse_range(range_str)
            # Convert to 1-based for openpyxl
            min_row += 1
            min_col += 1
            max_row += 1
            max_col += 1
        except ValueError as e:
            raise RangeError(f"Invalid range '{range_str}': {e}")
    
    # Extract data from the range
    data = []
    for row in range(min_row, max_row + 1):
        row_data = []
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            
            # Get the appropriate value (formula or calculated value)
            if formulas and cell.data_type == 'f':
                # If formulas are requested and the cell has one
                value = cell.value  # This is the formula with '='
            else:
                # Normal or calculated value
                value = cell.value
            
            row_data.append(value)
        data.append(row_data)
    
    return data

def list_tables(wb: Any, sheet_name: str) -> List[Dict[str, Any]]:
    """
    List all tables defined on an Excel sheet.
    
    Args:
        wb: Openpyxl workbook object
        sheet_name: Sheet name
        
    Returns:
        List of dictionaries with table information
        
    Raises:
        SheetNotFoundError: If the sheet does not exist
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # List to store table information
    tables_info = []
    
    # Check if the sheet has tables
    if hasattr(ws, 'tables') and ws.tables:
        for table_name, table in ws.tables.items():
            table_info = {
                'name': table_name,
                'ref': table.ref,
                'display_name': table.displayName,
                'header_row': table.headerRowCount > 0,
                'totals_row': table.totalsRowCount > 0,
                'style': table.tableStyleInfo.name if table.tableStyleInfo else None
            }
            
            tables_info.append(table_info)
    
    return tables_info

def get_table_data(wb: Any, sheet_name: str, table_name: str) -> List[Dict[str, Any]]:
    """
    Get the data from a specific table as records.
    
    Args:
        wb: Openpyxl workbook object
        sheet_name: Sheet name
        table_name: Table name
        
    Returns:
        List of dictionaries, where each dictionary represents a row
        
    Raises:
        SheetNotFoundError: If the sheet does not exist
        TableError: If the table does not exist
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # Check if the table exists
    if not hasattr(ws, 'tables') or table_name not in ws.tables:
        raise TableError(f"Table '{table_name}' does not exist on sheet '{sheet_name}'")
    
    # Get the table reference
    table = ws.tables[table_name]
    table_range = table.ref
    
    # Parse the range
    min_row, min_col, max_row, max_col = ExcelRange.parse_range(table_range)
    
    # Adjust to 1-based for openpyxl
    min_row += 1
    min_col += 1
    max_row += 1
    max_col += 1
    
    # Extract headers (first row)
    headers = []
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=min_row, column=col)
        headers.append(cell.value or f"Column{col}")
    
    # Extract data (rows after the header)
    data = []
    for row in range(min_row + 1, max_row + 1):
        row_data = {}
        for col_idx, col in enumerate(range(min_col, max_col + 1)):
            cell = ws.cell(row=row, column=col)
            header = headers[col_idx]
            row_data[header] = cell.value
        data.append(row_data)
    
    return data

def list_charts(wb: Any, sheet_name: str) -> List[Dict[str, Any]]:
    """
    List all charts on an Excel sheet.
    
    Args:
        wb: Openpyxl workbook object
        sheet_name: Sheet name
        
    Returns:
        List of dictionaries with chart information
        
    Raises:
        SheetNotFoundError: If the sheet does not exist
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # List to store chart information
    charts_info = []
    
    # Check if the sheet has charts
    if hasattr(ws, '_charts'):
        for chart_id, chart_rel in enumerate(ws._charts):
            chart = chart_rel[0]  # Element 0 is the chart object, 1 is position
            
            # Determine the chart type
            chart_type = "unknown"
            if isinstance(chart, BarChart):
                chart_type = "bar" if chart.type == "bar" else "column"
            elif isinstance(chart, LineChart):
                chart_type = "line"
            elif isinstance(chart, PieChart):
                chart_type = "pie"
            elif isinstance(chart, ScatterChart):
                chart_type = "scatter"
            elif isinstance(chart, AreaChart):
                chart_type = "area"
            
            # Gather chart information
            chart_info = {
                'id': chart_id,
                'type': chart_type,
                'title': chart.title if hasattr(chart, 'title') and chart.title else f"Chart {chart_id}",
                'position': chart_rel[1] if len(chart_rel) > 1 else None,
                'series_count': len(chart.series) if hasattr(chart, 'series') else 0
            }
            
            charts_info.append(chart_info)
    
    return charts_info

# 3. Escritura y formato de datos (de excel_writer_mcp.py)
def write_sheet_data(ws: Any, start_cell: str, data: List[List[Any]]) -> None:
    """
    Write a two-dimensional array of values or formulas.
     **Emojis must never be included in text written to cells, labels, titles or Excel charts.**


    To ensure the output remains readable when the function is used by a language
    model, it is recommended to apply styles after writing and check the length
    of the resulting cells. If any column contains very long text, its width
    should be increased to avoid cutting off the content. This way the generated
    files will look professional.

    Args:
        ws: Openpyxl worksheet object
        start_cell (str): Anchor cell (e.g. "A1")
        data (List[List]): Values or strings "=FORMULA(...)"
        
    Raises:
        CellReferenceError: If the cell reference is invalid
    """
    if not ws:
        raise ExcelMCPError("Worksheet cannot be None")
    
    if not data or not isinstance(data, list):
        raise ExcelMCPError("Data must be a non-empty list")
    
    try:
        # Parsear la celda inicial para obtener fila y columna base
        start_row, start_col = ExcelRange.parse_cell_ref(start_cell)

        # Escribir los datos
        for i, row_data in enumerate(data):
            if row_data is None:
                continue

            if not isinstance(row_data, list):
                # If it's not a list, treat it as a single value
                row_data = [row_data]

            for j, value in enumerate(row_data):
                # Calcular coordenadas de celda (base 1 para openpyxl)
                row = start_row + i + 1
                col = start_col + j + 1

                # Escribir el valor
                cell = ws.cell(row=row, column=col)
                cell.value = value

        # ----------------------------------------------------
        # Auto-fit columns and rows of the written range
        # ----------------------------------------------------
        end_row = start_row + len(data) - 1
        max_len_row = 0
        for row_data in data:
            if row_data is None:
                continue
            if isinstance(row_data, list):
                max_len_row = max(max_len_row, len(row_data))
            else:
                max_len_row = max(max_len_row, 1)
        end_col = start_col + max_len_row - 1
        cell_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
        try:
            autofit_table(ws, cell_range)
        except Exception:
            # No interrumpir escritura por un fallo de ajuste
            pass
    
    except ValueError as e:
        raise CellReferenceError(f"Invalid cell reference '{start_cell}': {e}")
    except Exception as e:
        raise ExcelMCPError(f"Error writing data: {e}")

def append_rows(ws: Any, data: List[List[Any]]) -> None:
    """
    Append rows at the end with the given values.
     **Emojis must never be included in text written to cells, labels, titles or Excel charts.**

    
    Args:
        ws: Openpyxl worksheet object
        data (List[List]): Values or strings "=FORMULA(...)"
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    if not data or not isinstance(data, list):
        raise ExcelMCPError("Data must be a non-empty list")
    
    try:
        for row_data in data:
            if not isinstance(row_data, list):
                # Si no es una lista, convertir a lista con un solo elemento
                row_data = [row_data]
            
            ws.append(row_data)
    
    except Exception as e:
        raise ExcelMCPError(f"Error adding rows: {e}")

def update_cell(ws: Any, cell: str, value_or_formula: Any) -> None:
    """
    Update a single cell.
     **Emojis must never be included in text written to cells, labels, titles or Excel charts.**

    
    Args:
        ws: Openpyxl worksheet object
        cell (str): Cell reference (e.g. "A1")
        value_or_formula: Value or formula to assign
        
    Raises:
        CellReferenceError: If the cell reference is invalid
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Assign value to the cell
        cell_obj = ws[cell]
        cell_obj.value = value_or_formula

        # ----------------------------------------------
        # Auto-fit if long text is written
        # ----------------------------------------------
        if isinstance(value_or_formula, str):
            text = value_or_formula
            lines = text.splitlines()
            max_len = max(len(line) for line in lines)

            column_letter = cell_obj.column_letter
            current_w = ws.column_dimensions[column_letter].width or 8.43
            desired_w = min(max_len + 2, 80)
            if desired_w > current_w:
                ws.column_dimensions[column_letter].width = desired_w

            if len(lines) > 1 or max_len > current_w:
                cell_obj.alignment = Alignment(wrap_text=True)
                est_lines = max(len(lines), math.ceil(max_len / max(desired_w, 1)))
                current_h = ws.row_dimensions[cell_obj.row].height or 15
                desired_h = est_lines * 15
                if desired_h > current_h:
                    ws.row_dimensions[cell_obj.row].height = desired_h


    except KeyError:
        raise CellReferenceError(f"Invalid cell reference: '{cell}'")
    except Exception as e:
        raise ExcelMCPError(f"Error updating cell: {e}")

def autofit_table(ws: Any, cell_range: str) -> None:
    """Adjust column widths and row heights for a tabular range."""
    start_row, start_col, end_row, end_col = ExcelRange.parse_range(cell_range)

    col_widths: Dict[int, int] = {}
    row_heights: Dict[int, int] = {}

    for row in range(start_row, end_row + 1):
        max_lines = 1
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row + 1, column=col + 1)
            value = cell.value
            if value is None:
                continue
            text = str(value)
            lines = text.splitlines()
            longest = max(len(line) for line in lines)
            col_widths[col] = max(col_widths.get(col, 0), longest)
            est_lines = max(len(lines), math.ceil(longest / 40))
            if est_lines > 1:
                cell.alignment = Alignment(wrap_text=True)
            max_lines = max(max_lines, est_lines)
        if max_lines > 1:
            row_heights[row] = max_lines * 15

    for col, width in col_widths.items():
        column_letter = get_column_letter(col + 1)
        current = ws.column_dimensions[column_letter].width or 8.43
        desired = min(width + 2, 80)
        if desired > current:
            ws.column_dimensions[column_letter].width = desired

    for row, height in row_heights.items():
        current = ws.row_dimensions[row + 1].height or 15
        if height > current:
            ws.row_dimensions[row + 1].height = height

def apply_style(ws: Any, cell_range: str, style_dict: Dict[str, Any]) -> None:
    """
    Apply cell styles to a range.

    Args:
        ws: Openpyxl worksheet object.
        cell_range (str): Range in ``A1:B5`` format or a single cell like ``"A1"``.
        style_dict (dict): Dictionary with styles to apply:
            - font_name (str): Font name.
            - font_size (int): Font size.
            - bold (bool): Bold text.
            - italic (bool): Italic text.
            - fill_color (str): Background color in hex, e.g. ``"FF0000"``.
            - border_style (str): Border style (``"thin"``, ``"medium"``, ``"thick"``, etc.).
            - alignment (str): Alignment (``"center"``, ``"left"``, ``"right"``, etc.).

    Raises:
        RangeError: If the range is invalid.
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Parse the range
        if ':' in cell_range:
            # Cell range
            start_cell, end_cell = cell_range.split(':')
            start_coord = ws[start_cell].coordinate
            end_coord = ws[end_cell].coordinate
            range_str = f"{start_coord}:{end_coord}"
        else:
            # A single cell
            range_str = cell_range
        
        # Preparar los estilos
        font_kwargs = {}
        if 'font_name' in style_dict:
            font_kwargs['name'] = style_dict['font_name']
        if 'font_size' in style_dict:
            font_kwargs['size'] = style_dict['font_size']
        if 'bold' in style_dict:
            font_kwargs['bold'] = style_dict['bold']
        if 'italic' in style_dict:
            font_kwargs['italic'] = style_dict['italic']
        if 'font_color' in style_dict:
            font_kwargs['color'] = style_dict['font_color']
        
        fill = None
        if 'fill_color' in style_dict:
            fill = PatternFill(start_color=style_dict['fill_color'], 
                              end_color=style_dict['fill_color'],
                              fill_type='solid')
        
        border = None
        if 'border_style' in style_dict:
            side = Side(style=style_dict['border_style'])
            border = Border(left=side, right=side, top=side, bottom=side)
        
        alignment = None
        if 'alignment' in style_dict:
            alignment_value = style_dict['alignment'].lower()
            horizontal = None
            
            # Map horizontal alignment values
            if alignment_value in ['left', 'center', 'right', 'justify']:
                horizontal = alignment_value
            
            alignment = Alignment(horizontal=horizontal)
        
        # Apply styles to all cells in the range
        for row in ws[range_str]:
            for cell in row:
                if font_kwargs:
                    cell.font = Font(**font_kwargs)
                if fill:
                    cell.fill = fill
                if border:
                    cell.border = border
                if alignment:
                    cell.alignment = alignment
    
    except KeyError:
        raise RangeError(f"Invalid range: '{cell_range}'")
    except Exception as e:
        raise ExcelMCPError(f"Error applying styles: {e}")

def apply_number_format(ws: Any, cell_range: str, fmt: str) -> None:
    """
    Apply a number format to a range of cells.

    Args:
        ws: Openpyxl worksheet object.
        cell_range (str): Range in ``A1:B5`` format or a single cell like ``"A1"``.
        fmt (str): Number format (``"#,##0.00"``, ``"0%"``, ``"dd/mm/yyyy"``, etc.).

    Raises:
        RangeError: If the range is invalid.
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Parse the range
        if ':' in cell_range:
            # Cell range
            start_cell, end_cell = cell_range.split(':')
            start_coord = ws[start_cell].coordinate
            end_coord = ws[end_cell].coordinate
            range_str = f"{start_coord}:{end_coord}"
        else:
            # A single cell
            range_str = cell_range
        
        # Apply the format to all cells in the range
        for row in ws[range_str]:
            for cell in row:
                cell.number_format = fmt
    
    except KeyError:
        raise RangeError(f"Invalid range: '{cell_range}'")
    except Exception as e:
        raise ExcelMCPError(f"Error applying number format: {e}")

# 4. Tables and formulas (from advanced_excel_mcp.py)
def add_table(ws: Any, table_name: str, cell_range: str, style=None) -> Any:
    """
    Define a range as a styled table.

    Args:
        ws: Openpyxl worksheet object.
        table_name (str): Unique name for the table.
        cell_range (str): Range in ``A1:B5`` format.
        style (str, optional): Predefined style name or a custom dict.

    Returns:
        The created :class:`Table` object.

    Raises:
        TableError: If there is a problem with the table, e.g. duplicate name.
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Check if a table with that name already exists
        if hasattr(ws, 'tables') and table_name in ws.tables:
            raise TableError(f"A table named '{table_name}' already exists")
        
        # Create table object
        table = Table(displayName=table_name, ref=cell_range)
        
        # Apply style
        if style:
            if isinstance(style, dict):
                # Custom style
                style_info = TableStyleInfo(**style)
            else:
                # Predefined style
                style_info = TableStyleInfo(
                    name=style,
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
            table.tableStyleInfo = style_info
        
        # Add table to the sheet
        ws.add_table(table)

        # ------------------------------
        # Auto-fit table columns and rows
        # ------------------------------
        try:
            autofit_table(ws, cell_range)
        except Exception:
            pass

        return table
    
    except Exception as e:
        if "Duplicate name" in str(e):
            raise TableError(f"A table named '{table_name}' already exists")
        elif "Invalid coordinate" in str(e) or "Invalid cell" in str(e):
            raise RangeError(f"Invalid range: '{cell_range}'")
        else:
            raise TableError(f"Error adding table: {e}")

def set_formula(ws: Any, cell: str, formula: str) -> Any:
    """
    Set a formula in a cell.

    Args:
        ws: Openpyxl worksheet object.
        cell (str): Cell reference (e.g. ``"A1"``).
        formula (str): Excel formula, with or without the ``=`` sign.

    Returns:
        The updated cell.

    Raises:
        FormulaError: If there is a problem with the formula.
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Add '=' sign if not present
        if formula and not formula.startswith('='):
            formula = f"={formula}"
        
        # Set the formula
        ws[cell] = formula
        return ws[cell]
    
    except KeyError:
        raise RangeError(f"Invalid cell: '{cell}'")
    except Exception as e:
        raise FormulaError(f"Error setting formula: {e}")

# 5. Charts and pivot tables (from advanced_excel_mcp.py)
def add_chart(
    wb: Any,
    sheet_name: str,
    chart_type: str,
    data_range: str,
    title=None,
    position=None,
    style=None,
    theme=None,
    custom_palette=None,
) -> Tuple[int, Any]:
    """Insert a native chart using the data from the given range.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    ``data_range`` should reference a rectangular table with no blank cells in
    the value area. The first row or first column is interpreted as headers and
    categories according to :func:`determine_orientation`. All series must
    contain only numbers and be the same length as the category vector. If total
    rows or mixed columns exist, the chart may be created incorrectly.

    Validate beforehand that the range belongs to the correct sheet and that
    headers are present, since the series are added with
    ``titles_from_data=True``. Categories must not contain blank or duplicate
    values and numeric columns must not contain text.

    Args:
        wb: Openpyxl ``Workbook`` object.
        sheet_name: Name of the sheet where the chart will be inserted.
        chart_type: Chart type (``'column'``, ``'bar'``, ``'line'``, ``'pie'``, etc.).
        data_range: Data range in ``A1:B5`` format.
        title: Optional chart title.
        position: Cell where the chart will be placed (e.g. ``"E5"``).
        style: Chart style (number ``1``–``48`` or descriptive name).
        theme: Name of the color theme.
        custom_palette: List of custom colors.

    Returns:
        Tuple ``(chart id, chart object)``.

    Raises:
        ChartError: If a problem occurs while creating the chart.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    try:
        # Get the sheet
        ws = get_sheet(wb, sheet_name)
        
        # Validate and normalize the position
        if position:
            # If the position is a range, take only the first cell
            if ':' in position:
                position = position.split(':')[0]
            
            # Validate that it is a valid cell reference
            try:
                # Try parsing to verify it is a valid reference
                ExcelRange.parse_cell_ref(position)
            except ValueError:
                raise ValueError(f"Invalid position '{position}'. Must be a cell reference (e.g. 'E4')")
        
        # Create the chart object according to the type
        chart = None
        if chart_type.lower() == 'column':
            chart = BarChart()
            chart.type = "col"
        elif chart_type.lower() == 'bar':
            chart = BarChart()
            chart.type = "bar"
        elif chart_type.lower() == 'line':
            chart = LineChart()
        elif chart_type.lower() == 'pie':
            chart = PieChart()
        elif chart_type.lower() == 'scatter':
            chart = ScatterChart()
        elif chart_type.lower() == 'area':
            chart = AreaChart()
        else:
            raise ChartError(f"Chart type not supported: '{chart_type}'")
        
        # Set title if provided
        if title:
            chart.title = title
            
        # Determine if the range references another sheet
        data_sheet_name, sr, sc, er, ec = ExcelRange.parse_range_with_sheet(data_range)
        if data_sheet_name is None:
            data_sheet_name = sheet_name
        data_ws = get_sheet(wb, data_sheet_name)

        # Normalize the range for Reference (escaping the sheet name)
        if " " in data_sheet_name or any(c in data_sheet_name for c in "![]{}?"):
            sheet_prefix = f"'{data_sheet_name}'!"
        else:
            sheet_prefix = f"{data_sheet_name}!"
        clean_range = ExcelRange.range_to_a1(sr, sc, er, ec)
        
        # Parse data range
        try:
            # Use the previously calculated limits
            min_row = sr + 1
            min_col = sc + 1
            max_row = er + 1
            max_col = ec + 1

            # Trim empty rows or columns at the end
            min_row, min_col, max_row, max_col = _trim_range_to_data(data_ws, min_row, min_col, max_row, max_col)
            if max_row < min_row or max_col < min_col:
                raise ChartError("The specified range contains no data")

            # Determine orientation by analyzing the range contents
            is_column_oriented = determine_orientation(data_ws, min_row, min_col, max_row, max_col)
            
            # For charts that need categories (most except scatter)
            if chart_type.lower() != 'scatter':
                if is_column_oriented:
                    if _range_has_blank(data_ws, min_row + 1, min_col + 1, max_row, max_col):
                        raise ChartError("The data range contains blank cells")
                    categories = Reference(data_ws, min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=min_col)
                    data = Reference(data_ws, min_row=min_row, max_row=max_row, min_col=min_col + 1, max_col=max_col)
                    try:
                        chart.add_data(data, titles_from_data=True)
                    except TypeError:
                        chart.add_data(data)
                    chart.set_categories(categories)
                else:
                    if _range_has_blank(data_ws, min_row + 1, min_col, max_row, max_col):
                        raise ChartError("The data range contains blank cells")
                    categories = Reference(data_ws, min_row=min_row, max_row=min_row, min_col=min_col, max_col=max_col)
                    data = Reference(data_ws, min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=max_col)
                    try:
                        chart.add_data(data, titles_from_data=True)
                    except TypeError:
                        chart.add_data(data)
                    chart.set_categories(categories)
            else:
                if _range_has_blank(data_ws, min_row, min_col, max_row, max_col):
                    raise ChartError("The data range contains blank cells")
                data_ref = Reference(data_ws, min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col)
                chart.add_data(data_ref)
        
        except Exception as e:
            raise RangeError(f"Error processing data range '{data_range}': {e}")
        
        # Apply styles
        if style is not None:
            # Convert specified style (number, name, etc.)
            style_number = parse_chart_style(style)
            if style_number is not None:
                # Apply the style including the color palette
                apply_chart_style(chart, style_number)
            else:
                logger.warning(f"Invalid chart style: '{style}'. Using default style.")
        
        # Apply color theme if provided
        # (here we would use the theme but omit it for simplicity)
        
        # Apply custom palette if provided
        if custom_palette and isinstance(custom_palette, list):
            from openpyxl.chart.shapes import GraphicalProperties
            from openpyxl.drawing.fill import ColorChoice
            
            for i, series in enumerate(chart.series):
                if i < len(custom_palette):
                    # Ensure graphical properties exist
                    if not hasattr(series, 'graphicalProperties'):
                        series.graphicalProperties = GraphicalProperties()
                    elif series.graphicalProperties is None:
                        series.graphicalProperties = GraphicalProperties()
                    
                    # Assign color ensuring it doesn't have the # prefix
                    color = custom_palette[i]
                    if isinstance(color, str) and color.startswith('#'):
                        color = color[1:]
                    
                    # Apply the color explicitly
                    series.graphicalProperties.solidFill = ColorChoice(srgbClr=color)
        
        # Position the chart on the sheet
        if position:
            ws.add_chart(chart, position)
        else:
            ws.add_chart(chart)
        
        # Determine the chart ID (based on its position in the list)
        chart_id = len(ws._charts) - 1
        
        return chart_id, chart
    
    except SheetNotFoundError:
        raise
    except ChartError:
        raise
    except RangeError:
        raise
    except Exception as e:
        raise ChartError(f"Error creating chart: {e}")

def add_pivot_table(wb: Any, source_sheet: str, source_range: str, target_sheet: str,
                   target_cell: str, rows: List[str], cols: List[str], data_fields: List[str]) -> Any:
    """
    Create a pivot table.

    Args:
        wb: Openpyxl workbook object.
        source_sheet (str): Sheet containing the source data.
        source_range (str): Range of source data (e.g. ``A1:E10``).
        target_sheet (str): Sheet where the pivot table will be created.
        target_cell (str): Anchor cell (e.g. ``"A1"``).
        rows (list): Row fields.
        cols (list): Column fields.
        data_fields (list): Value fields and functions.

    Returns:
        The created :class:`PivotTable` object.

    Raises:
        PivotTableError: If there are issues with the pivot table.
    """
    if not wb:
        raise ExcelMCPError("Workbook cannot be None")
    
    try:
        # Get source sheet
        source_ws = get_sheet(wb, source_sheet)

        # Get target sheet
        target_ws = get_sheet(wb, target_sheet)

        logger.warning("Pivot tables in openpyxl have limited functionality and may not work as expected.")

        # Try creating the data cache (this is a required step)
        try:
            # Parse the range
            min_row, min_col, max_row, max_col = ExcelRange.parse_range(source_range)

            # Adjust to 1-based for Reference
            min_row += 1
            min_col += 1
            max_row += 1
            max_col += 1

            # Create data reference for the cache
            data_reference = Reference(source_ws, min_row=min_row, min_col=min_col,
                                     max_row=max_row, max_col=max_col)

            # Create pivot cache
            pivot_cache = PivotCache(cacheSource=data_reference, cacheDefinition={'refreshOnLoad': True})

            # Generate a unique ID for the pivot table
            pivot_name = f"PivotTable{len(wb._pivots) + 1 if hasattr(wb, '_pivots') else 1}"

            # Create the pivot table
            pivot_table = PivotTable(name=pivot_name, cache=pivot_cache,
                                    location=target_cell, rowGrandTotals=True, colGrandTotals=True)

            # Add row fields
            for row_field in rows:
                pivot_table.rowFields.append(PivotField(data=row_field))

            # Add column fields
            for col_field in cols:
                pivot_table.colFields.append(PivotField(data=col_field))

            # Add data fields
            for data_field in data_fields:
                pivot_table.dataFields.append(PivotField(data=data_field))

            # Add the pivot table to the target sheet
            target_ws.add_pivot_table(pivot_table)
            
            return pivot_table
            
        except Exception as pivot_error:
            logger.error(f"Error creating pivot table: {pivot_error}")
            raise PivotTableError(f"Error creating pivot table: {pivot_error}")
    
    except SheetNotFoundError:
        raise
    except PivotTableError:
        raise
    except Exception as e:
        raise PivotTableError(f"Error creating pivot table: {e}")

# ----------------------------------------
# NEW COMBINED HIGH-LEVEL FUNCTIONS
# ----------------------------------------

def create_sheet_with_data(wb: Any, sheet_name: str, data: List[List[Any]],
                           index: Optional[int] = None, overwrite: bool = False) -> Any:
    """
    Create a new sheet and write data in a single step.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        wb: Openpyxl workbook object.
        sheet_name (str): Name for the new sheet.
        data (List[List]): Data to write.
        index (int, optional): Position of the sheet in the workbook.
        overwrite (bool): If ``True`` overwrite an existing sheet with the same name.

    Returns:
        The created worksheet object.

    Raises:
        SheetExistsError: If the sheet already exists and ``overwrite`` is ``False``.
    """
    # Handle existing sheet case
    if sheet_name in list_sheets(wb):
        if overwrite:
            # Delete the existing sheet
            delete_sheet(wb, sheet_name)
        else:
            raise SheetExistsError(f"Sheet '{sheet_name}' already exists. Use overwrite=True to overwrite it.")
    
    # Create new sheet
    ws = add_sheet(wb, sheet_name, index)
    
    # Write data
    if data:
        write_sheet_data(ws, "A1", data)
    
    return ws

def create_formatted_table(wb: Any, sheet_name: str, start_cell: str, data: List[List[Any]],
                            table_name: str, table_style: Optional[str] = None,
                            formats: Optional[Dict[str, Union[str, Dict]]] = None) -> Tuple[Any, Any]:
    """
    Create a formatted table in a single step.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        wb: Openpyxl workbook object.
        sheet_name (str): Name of the sheet where the table will be created.
        start_cell (str): Starting cell for the data (e.g. ``"A1"``).
        data (List[List]): Data for the table including headers.
        table_name (str): Unique name for the table.
        table_style (str, optional): Predefined table style (e.g. ``"TableStyleMedium9"``).
        formats (dict, optional): Dictionary of formats to apply:
            - Keys: Relative ranges (e.g. ``"A2:A10"``) or cells.
            - Values: Number format or a style dictionary.

    Returns:
        Tuple ``(table object, worksheet)``.

    Example format::
        
        formats = {
            "B2:B10": "#,##0.00",  # Currency format
            "A1:Z1": {"bold": True, "fill_color": "DDEBF7"}  # Header style
        }
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # Get data range dimensions
    rows = len(data)
    cols = max([len(row) if isinstance(row, list) else 1 for row in data], default=0)
    
    # Write the data
    write_sheet_data(ws, start_cell, data)
    
    # Calculate the full table range
    start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1
    full_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
    
    # Create the table
    table = add_table(ws, table_name, full_range, table_style)
    
    # Apply additional formats if provided
    if formats:
        for range_str, format_value in formats.items():
            # Convert relative range to absolute if needed
            if not any(c in range_str for c in [':', '!']):
                # It's a single cell, add offset
                cell_row, cell_col = ExcelRange.parse_cell_ref(range_str)
                abs_row = start_row + cell_row
                abs_col = start_col + cell_col
                abs_range = ExcelRange.cell_to_a1(abs_row, abs_col)
            elif ':' in range_str and '!' not in range_str:
                # Range without a specific sheet, add offset
                range_start, range_end = range_str.split(':')
                start_row_rel, start_col_rel = ExcelRange.parse_cell_ref(range_start)
                end_row_rel, end_col_rel = ExcelRange.parse_cell_ref(range_end)

                # Calculate absolute positions
                abs_start_row = start_row + start_row_rel
                abs_start_col = start_col + start_col_rel
                abs_end_row = start_row + end_row_rel
                abs_end_col = start_col + end_col_rel

                # Create absolute range
                abs_range = ExcelRange.range_to_a1(abs_start_row, abs_start_col, abs_end_row, abs_end_col)
            else:
                # It's already an absolute range or includes the sheet
                abs_range = range_str

            # Apply format according to type
            if isinstance(format_value, str):
                # It's a number format
                apply_number_format(ws, abs_range, format_value)
            elif isinstance(format_value, dict):
                # It's a style dictionary
                apply_style(ws, abs_range, format_value)
    
    return table, ws

def create_chart_from_table(
    wb: Any,
    sheet_name: str,
    table_name: str,
    chart_type: str,
    title: Optional[str] = None,
    position: Optional[str] = None,
    style: Optional[Any] = None,
    use_headers: bool = True,
) -> Tuple[int, Any]:
    """Generate a chart from an existing table.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    The table must contain valid headers and must not include total rows. Data
    cells are assumed to form a rectangular range with no blanks. When
    ``use_headers`` is ``True`` the first row of the table is used as series
    titles and categories. All data columns must be numeric and the same length
    to avoid errors when creating the chart.

    Ensure the table has no blank cells or text columns where numbers are
    expected. Any mismatch in series length or category count can lead to
    incomplete or empty charts.

    Args:
        wb: Openpyxl ``Workbook`` object.
        sheet_name: Name of the sheet containing the table.
        table_name: Name of the table used as the source.
        chart_type: Chart type (``'column'``, ``'bar'``, ``'line'``, ``'pie'``, etc.).
        title: Optional chart title.
        position: Anchor cell for the chart.
        style: Chart style (number ``1``–``48`` or descriptive name).
        use_headers: If ``True`` use the first row as headers and categories.

    Returns:
        Tuple ``(chart ID, chart object)``.
    """
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # Get table information
    tables = list_tables(wb, sheet_name)
    table_info = None
    for table in tables:
        if table['name'] == table_name:
            table_info = table
            break
    
    if not table_info:
        raise TableError(f"Table '{table_name}' not found in sheet '{sheet_name}'")
    
    # Use the table range to create the chart
    table_range = table_info['ref']
    
    # Create the chart
    chart_id, chart = add_chart(wb, sheet_name, chart_type, table_range, 
                               title, position, style)
    
    return chart_id, chart

def create_chart_from_data(
    wb: Any,
    sheet_name: str,
    data: List[List[Any]],
    chart_type: str,
    position: Optional[str] = None,
    title: Optional[str] = None,
    style: Optional[Any] = None,
    create_table: bool = False,
    table_name: Optional[str] = None,
    table_style: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a chart from ``data`` by writing the values first.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    ``data`` must be a list of lists forming a rectangular structure with no
    empty cells. The first row or column is interpreted as headers and
    categories; therefore every row must have the same length and numeric columns
    should not contain text. Avoid including total rows or records that should
    not be charted.

    Before calling the function check for blank cells, duplicated headers or
    mismatched lengths between categories and series. ``add_chart`` uses
    ``titles_from_data=True`` to assign the series names. If the series are
    inconsistent the resulting chart may be incomplete or show errors.

    Args:
        wb: Openpyxl ``Workbook`` object.
        sheet_name: Name of the sheet where the chart will be created.
        data: Data matrix including the headers.
        chart_type: Chart type (``'column'``, ``'bar'``, ``'line'``, ``'pie'``, etc.).
        position: Cell where to place the chart.
        title: Chart title.
        style: Chart style (number ``1``–``48`` or descriptive name).
        create_table: If ``True`` a table with the written data will be created.
        table_name: Name of the table (required if ``create_table`` is ``True``).
        table_style: Optional table style.

    Returns:
        Dictionary with information about the chart and, if created, the table.
    """
    # Create sheet if it does not exist
    if sheet_name not in list_sheets(wb):
        add_sheet(wb, sheet_name)
    
    # Get the sheet
    ws = get_sheet(wb, sheet_name)
    
    # Determine a suitable location for the data
    # By default, place the data at A1
    data_start_cell = "A1"
    
    # Write the data
    write_sheet_data(ws, data_start_cell, data)
    
    # Calculate the full data range
    rows = len(data)
    cols = max([len(row) if isinstance(row, list) else 1 for row in data], default=0)
    start_row, start_col = ExcelRange.parse_cell_ref(data_start_cell)
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1
    data_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
    
    result = {
        "data_range": data_range,
        "rows": rows,
        "columns": cols
    }
    
    # Create table if requested
    if create_table:
        if not table_name:
            # Generate table name if not provided
            table_name = f"Table_{sheet_name}_{int(time.time())}"
            
        try:
            table = add_table(ws, table_name, data_range, table_style)
            result["table"] = {
                "name": table_name,
                "range": data_range,
                "style": table_style
            }
        except Exception as e:
            logger.warning(f"Could not create the table: {e}")
    
    # Create the chart
    try:
        chart_id, chart = add_chart(wb, sheet_name, chart_type, data_range, 
                                  title, position, style)
        
        result["chart"] = {
            "id": chart_id,
            "type": chart_type,
            "title": title,
            "position": position,
            "style": style
        }
    except Exception as e:
        logger.error(f"Error creating chart: {e}")
        raise ChartError(f"Error creating chart: {e}")

    return result

def create_chart_from_dataframe(
    wb: Any,
    sheet_name: str,
    df: 'pd.DataFrame',
    chart_type: str,
    position: Optional[str] = None,
    title: Optional[str] = None,
    style: Optional[Any] = None,
    create_table: bool = False,
    table_name: Optional[str] = None,
    table_style: Optional[str] = None,
) -> Dict[str, Any]:
    """Generate a chart from a ``pandas.DataFrame``.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    The ``DataFrame`` must contain numeric columns without missing values in the
    series and should not include total rows. The headers are used as series
    titles, so duplicates or blank cells should be avoided. The contents of the
    ``DataFrame`` are written to the sheet and delegated to
    :func:`create_chart_from_data`, therefore the same validation rules apply.

    Args:
        wb: Openpyxl ``Workbook`` object.
        sheet_name: Name of the sheet where the chart will be created.
        df: Data in ``pandas.DataFrame`` format.
        chart_type: Chart type (``'column'``, ``'bar'``, ``'line'``, ``'pie'``, etc.).
        position: Anchor cell for the chart.
        title: Chart title.
        style: Optional chart style.
        create_table: If ``True`` create a table with the written data.
        table_name: Name of the table to create.
        table_style: Table style.

    Returns:
        Dictionary with information about the chart and, if created, the table.
    """

    if df is None:
        raise ExcelMCPError("The provided DataFrame is None")

    # Convert the DataFrame to a list of lists including headers
    data = [df.columns.tolist()] + df.values.tolist()

    return create_chart_from_data(
        wb,
        sheet_name,
        data,
        chart_type,
        position=position,
        title=title,
        style=style,
        create_table=create_table,
        table_name=table_name,
        table_style=table_style,
    )

def create_report(wb: Any, data: Dict[str, List[List[Any]]], tables: Optional[Dict[str, Dict[str, Any]]] = None,
                 charts: Optional[Dict[str, Dict[str, Any]]] = None, formats: Optional[Dict[str, Dict[str, Any]]] = None,
                 overwrite_sheets: bool = False) -> Dict[str, Any]:
    """
    Create a full report with multiple sheets, tables and charts in a single step.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    This function acts as a generic template for automated report generators.
    All created sheets should be orderly and styled. Check the available space
    before inserting charts so they do not end up on top of any table or block
    of text. After creating a table, verify which column contains the longest
    strings and adjust its width so the content is visible without manual
    editing.

    Args:
        wb: Openpyxl workbook object.
        data: Dictionary with data per sheet: ``{"Sheet1": [[data]], "Sheet2": [[data]]}``
        tables: Dictionary with table configuration:
            ``{"SalesTable": {"sheet": "Sales", "range": "A1:B10", "style": "TableStyleMedium9"}}``
        charts: Dictionary with chart configuration:
            ``{"SalesChart": {"sheet": "Sales", "type": "column", "data": "SalesTable",
                              "title": "Sales", "position": "D2", "style": "dark-blue"}}``
        formats: Dictionary of formats to apply:
            ``{"Sales": {"B2:B10": "#,##0.00", "A1:Z1": {"bold": True}}}``
        overwrite_sheets: If ``True`` overwrite existing sheets.

    Returns:
        Dictionary with information about the created elements.
    """
    result = {
        "sheets": [],
        "tables": [],
        "charts": []
    }
    
    # Create/update sheets with data
    for sheet_name, sheet_data in data.items():
        if sheet_name in list_sheets(wb):
            if overwrite_sheets:
                # Use the existing sheet
                ws = wb[sheet_name]
                # Write the data
                write_sheet_data(ws, "A1", sheet_data)
            else:
                # Add numeric suffix if the sheet already exists
                base_name = sheet_name
                counter = 1
                while f"{base_name}_{counter}" in list_sheets(wb):
                    counter += 1
                new_name = f"{base_name}_{counter}"
                ws = create_sheet_with_data(wb, new_name, sheet_data)
                sheet_name = new_name
        else:
            # Create new sheet
            ws = create_sheet_with_data(wb, sheet_name, sheet_data)
        
        result["sheets"].append({"name": sheet_name, "rows": len(sheet_data)})
        
        # Apply specific formats for this sheet
        if formats and sheet_name in formats:
            for range_str, format_value in formats[sheet_name].items():
                if isinstance(format_value, str):
                    apply_number_format(ws, range_str, format_value)
                elif isinstance(format_value, dict):
                    apply_style(ws, range_str, format_value)
    
    # Crear tablas
    if tables:
        for table_name, table_config in tables.items():
            sheet_name = table_config.get("sheet")
            range_str = table_config.get("range")
            style = table_config.get("style")
            
            if not sheet_name or not range_str:
                logger.warning(f"Incomplete configuration for table '{table_name}'. Sheet and range are required.")
                continue
            
            try:
                # Verify that the sheet exists
                if sheet_name not in list_sheets(wb):
                    logger.warning(f"Sheet '{sheet_name}' not found for table '{table_name}'. Skipping.")
                    continue
                
                ws = wb[sheet_name]
                table = add_table(ws, table_name, range_str, style)
                
                result["tables"].append({
                    "name": table_name,
                    "sheet": sheet_name,
                    "range": range_str,
                    "style": style
                })
                
                # Apply specific formats for this table
                if "formats" in table_config:
                    for range_str, format_value in table_config["formats"].items():
                        if isinstance(format_value, str):
                            apply_number_format(ws, range_str, format_value)
                        elif isinstance(format_value, dict):
                            apply_style(ws, range_str, format_value)
            
            except Exception as e:
                logger.warning(f"Error al crear tabla '{table_name}': {e}")
    
    # Create charts
    if charts:
        for chart_name, chart_config in charts.items():
            sheet_name = chart_config.get("sheet")
            chart_type = chart_config.get("type")
            data_source = chart_config.get("data")
            title = chart_config.get("title", chart_name)
            position = chart_config.get("position")
            style = chart_config.get("style")
            
            if not sheet_name or not chart_type or not data_source:
                logger.warning(f"Incomplete configuration for chart '{chart_name}'. Sheet, type and data are required.")
                continue
            
            try:
                # Verificar que la hoja existe
                if sheet_name not in list_sheets(wb):
                    logger.warning(f"Sheet '{sheet_name}' not found for chart '{chart_name}'. Skipping.")
                    continue
                
                # Determinar si data_source es una tabla o un rango
                data_range = data_source
                if data_source in [t["name"] for t in result["tables"]]:
                    # Es una tabla, obtener su rango
                    for table in result["tables"]:
                        if table["name"] == data_source:
                            data_range = table["range"]
                            break
                
                # Create the chart
                chart_id, chart = add_chart(wb, sheet_name, chart_type, data_range, 
                                           title, position, style)
                
                result["charts"].append({
                    "name": chart_name,
                    "id": chart_id,
                    "sheet": sheet_name,
                    "type": chart_type,
                    "data_source": data_source,
                    "position": position,
                    "style": style
                })
            
            except Exception as e:
                logger.warning(f"Error creating chart '{chart_name}': {e}")
    
    return result

def create_dashboard(wb: Any, dashboard_config: Dict[str, Any],
                    create_new: bool = True) -> Dict[str, Any]:
    """
    Create a complete dashboard with tables, charts and interactive filters.

     **Emojis must never be included in text written to cells, labels, titles or charts.**

    It is intended for an automated agent to build an attractive sheet without
    overlaps. Each chart should be placed with space from previous tables or
    text. After writing the data for each section check the column widths and
    enlarge them when some cell is particularly long so the sheet remains easy
    to read without manual editing.

    Args:
        wb: Openpyxl workbook object.
        dashboard_config: Dictionary with the complete dashboard configuration
            {
                "title": "Sales Dashboard",
                "sheet": "Dashboard",
                "data_sheet": "Data",
                "data": [[data]],
                "sections": [
                    {
                        "title": "Sales by Region",
                        "type": "chart",
                        "chart_type": "column",
                        "data_range": "A1:B10",
                        "position": "E1",
                        "style": "dark-blue"
                    },
                    {
                        "title": "Product Table",
                        "type": "table",
                        "data_range": "D1:F10",
                        "name": "ProductTable",
                        "style": "TableStyleMedium9"
                    }
                ]
            }
        create_new: If ``True`` create a new sheet for the dashboard.

    Returns:
        Dictionary with information about the created elements.
    """
    # Basic configuration
    title = dashboard_config.get("title", "Dashboard")
    sheet_name = dashboard_config.get("sheet", "Dashboard")
    data_sheet = dashboard_config.get("data_sheet")
    data = dashboard_config.get("data")
    
    result = {
        "title": title,
        "sheet": sheet_name,
        "sections": []
    }
    
    # Crear o usar la hoja del dashboard
    if sheet_name in list_sheets(wb):
        if create_new:
            # Add numeric suffix
            base_name = sheet_name
            counter = 1
            while f"{base_name}_{counter}" in list_sheets(wb):
                counter += 1
            sheet_name = f"{base_name}_{counter}"
            ws = add_sheet(wb, sheet_name)
            result["sheet"] = sheet_name
        else:
            # Usar la hoja existente
            ws = wb[sheet_name]
    else:
        # Crear nueva hoja
        ws = add_sheet(wb, sheet_name)
    
    # Crear hoja de datos si se proporciona
    if data_sheet and data:
        if data_sheet in list_sheets(wb):
            # Usar la hoja existente
            data_ws = wb[data_sheet]
            # Clear existing data
            # (This could be improved to avoid deleting everything)
            max_row = data_ws.max_row
            max_col = data_ws.max_column
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    data_ws.cell(row=row, column=col).value = None
        else:
            # Crear nueva hoja de datos
            data_ws = add_sheet(wb, data_sheet)
        
        # Escribir los datos
        write_sheet_data(data_ws, "A1", data)
        result["data_sheet"] = data_sheet
    
    # Add title to the dashboard
    update_cell(ws, "A1", title)
    apply_style(ws, "A1", {
        "font_size": 16,
        "bold": True,
        "alignment": "center"
    })
    

    
    # Space after the title
    current_row = 3
    
    # Process dashboard sections
    sections = dashboard_config.get("sections", [])
    for i, section in enumerate(sections):
        section_type = section.get("type")
        section_title = section.get("title", f"Section {i+1}")
        
        # Information for the result
        section_result = {
            "title": section_title,
            "type": section_type,
            "row": current_row
        }
        
        # Add section title
        update_cell(ws, f"A{current_row}", section_title)
        apply_style(ws, f"A{current_row}", {
            "font_size": 12,
            "bold": True
        })
        current_row += 1
        
        # Process according to the section type
        if section_type == "chart":
            chart_type = section.get("chart_type", "column")
            data_range = section.get("data_range")
            
            # If the range has no specific sheet, use the data sheet
            if data_range and '!' not in data_range and data_sheet:
                if ' ' in data_sheet or any(c in data_sheet for c in "![]{}?"):
                    data_range = f"'{data_sheet}'!{data_range}"
                else:
                    data_range = f"{data_sheet}!{data_range}"
            
            chart_position = section.get("position", f"A{current_row}")
            chart_title = section.get("title", section_title)
            chart_style = section.get("style")
            
            try:
                chart_id, chart = add_chart(wb, sheet_name, chart_type, data_range, 
                                          chart_title, chart_position, chart_style)
                
                section_result["chart_id"] = chart_id
                section_result["data_range"] = data_range
                
                # Move down rows according to position and estimated chart size
                # (this is a simple estimate)
                current_row += 15
            except Exception as e:
                logger.warning(f"Error creating chart in section '{section_title}': {e}")
                current_row += 2  # Move down a few rows in case of error
        
        elif section_type == "table":
            table_range = section.get("data_range")
            table_name = section.get("name", f"Table_{i}")
            table_style = section.get("style")
            
            # If the range has no specific sheet, use the data sheet
            if table_range and '!' not in table_range and data_sheet:
                if ' ' in data_sheet or any(c in data_sheet for c in "![]{}?"):
                    full_table_range = f"'{data_sheet}'!{table_range}"
                else:
                    full_table_range = f"{data_sheet}!{table_range}"
            else:
                full_table_range = table_range
                
            try:
                # Extract table data to display on the dashboard
                if data_sheet:
                    source_ws = wb[data_sheet]
                    # Extract range without sheet name
                    if '!' in table_range:
                        pure_range = table_range.split('!')[1]
                    else:
                        pure_range = table_range
                    
                    # Read data from the source
                    table_data = read_sheet_data(wb, data_sheet, pure_range)
                    
                    # Determine dimensions
                    table_rows = len(table_data)
                    table_cols = max([len(row) if isinstance(row, list) else 1 for row in table_data], default=0)
                    
                    # Write data on the dashboard
                    write_sheet_data(ws, f"A{current_row}", table_data)
                    
                    # Create local table on the dashboard
                    local_range = f"A{current_row}:{get_column_letter(table_cols)}:{current_row + table_rows - 1}"
                    table = add_table(ws, table_name, local_range, table_style)
                    
                    section_result["table_name"] = table_name
                    section_result["source_range"] = full_table_range
                    section_result["dashboard_range"] = local_range
                    
                    # Move down rows according to table size
                    current_row += table_rows + 2
                else:
                    # Si no hay hoja de datos, crear tabla directamente en el dashboard
                    table = add_table(ws, table_name, table_range, table_style)
                    
                    section_result["table_name"] = table_name
                    section_result["range"] = table_range
                    
                    # Estimar filas para avanzar
                    try:
                        min_row, min_col, max_row, max_col = ExcelRange.parse_range(table_range)
                        current_row += (max_row - min_row) + 3
                    except:
                        current_row += 10  # Default value if calculation fails
            except Exception as e:
                logger.warning(f"Error creating table in section '{section_title}': {e}")
                current_row += 2
        
        elif section_type == "text":
            text_content = section.get("content", "")
            cell_ref = f"A{current_row}"
            
            update_cell(ws, cell_ref, text_content)
            
            # Aplicar formato si se especifica
            text_format = section.get("format", {})
            if text_format:
                apply_style(ws, cell_ref, text_format)
            
            section_result["content"] = text_content
            section_result["cell"] = cell_ref
            
            current_row += 2
        
        # Add the section to the result
        result["sections"].append(section_result)
        
        # Space between sections
        current_row += 1
    
    return result

def apply_excel_template(wb: Any, template_name: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Apply a predefined template to an Excel workbook.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        wb: Openpyxl workbook object.
        template_name (str): Name of the template to apply (e.g. ``"sales_report"``, ``"dashboard"``).
        data: Dictionary with data specific to the template.

    Returns:
        Dictionary with information about the created elements.

    Available templates:
        - ``"basic_report"``: Basic report with table and chart.
        - ``"financial_dashboard"``: Financial dashboard with multiple KPIs and charts.
        - ``"sales_analysis"``: Sales analysis by region and product.
        - ``"project_tracker"``: Project tracker with progress tables and charts.
    """
    result = {
        "template": template_name,
        "sheets": [],
        "elements": []
    }
    
    # Implementation of predefined templates
    if template_name == "basic_report":
        # Basic report template
        title = data.get("title", "Basic Report")
        subtitle = data.get("subtitle", "")
        report_date = data.get("date", time.strftime("%d/%m/%Y"))
        sheet_name = data.get("sheet", "Report")
        report_data = data.get("data", [])
        
        # Crear hoja para el informe si no existe
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Title and basic information
        update_cell(ws, "A1", title)
        apply_style(ws, "A1", {
            "font_size": 16,
            "bold": True,
            "alignment": "center"
        })
        
        if subtitle:
            update_cell(ws, "A2", subtitle)
            apply_style(ws, "A2", {
                "font_size": 12,
                "alignment": "center"
            })
        
        update_cell(ws, "A3", f"Fecha: {report_date}")
        
        # Crear tabla con los datos
        start_row = 5
        if report_data:
            write_sheet_data(ws, f"A{start_row}", report_data)
            
            # Determinar dimensiones
            rows = len(report_data)
            cols = max([len(row) if isinstance(row, list) else 1 for row in report_data], default=0)
            
            # Crear tabla
            table_range = f"A{start_row}:{get_column_letter(cols)}{start_row + rows - 1}"
            table_name = data.get("table_name", "ReportTable")
            table_style = data.get("table_style", "TableStyleMedium9")
            
            try:
                table = add_table(ws, table_name, table_range, table_style)
                result["elements"].append({
                    "type": "table",
                    "name": table_name,
                    "range": table_range
                })
            except Exception as e:
                logger.warning(f"Error al crear tabla: {e}")
            
            # Create chart
            chart_type = data.get("chart_type", "column")
            chart_position = data.get("chart_position", f"G{start_row}")
            chart_title = data.get("chart_title", "Report Chart")
            chart_style = data.get("chart_style", "colorful-1")
            
            try:
                chart_id, chart = add_chart(wb, sheet_name, chart_type, table_range, 
                                          chart_title, chart_position, chart_style)
                
                result["elements"].append({
                    "type": "chart",
                    "id": chart_id,
                    "position": chart_position
                })
            except Exception as e:
                logger.warning(f"Error creating chart: {e}")
        
        result["sheets"].append({"name": sheet_name, "type": "report"})
    
    elif template_name == "financial_dashboard":
        # More advanced template for a financial dashboard
        title = data.get("title", "Financial Dashboard")
        sheet_name = data.get("sheet", "Dashboard")
        financial_data = data.get("financial_data", {})
        
        # Configuration to create the full dashboard
        dashboard_config = {
            "title": title,
            "sheet": sheet_name,
            "sections": []
        }
        
        # 1. Financial KPI section
        if "kpis" in financial_data:
            kpis = financial_data["kpis"]
            kpi_section = {
                "title": "Key Financial Indicators",
                "type": "text",
                "content": "Financial KPIs"
            }
            dashboard_config["sections"].append(kpi_section)

            # Each KPI could be added as text or a formatted cell
            for kpi_name, kpi_value in kpis.items():
                kpi_section = {
                    "title": kpi_name,
                    "type": "text",
                    "content": f"{kpi_name}: {kpi_value}",
                    "format": {
                        "bold": True,
                        "font_size": 12
                    }
                }
                dashboard_config["sections"].append(kpi_section)
        
        # 2. Financial charts section
        if "charts" in financial_data:
            for chart_config in financial_data["charts"]:
                chart_section = {
                    "title": chart_config.get("title", "Financial Chart"),
                    "type": "chart",
                    "chart_type": chart_config.get("type", "column"),
                    "data_range": chart_config.get("data_range", ""),
                    "position": chart_config.get("position", ""),
                    "style": chart_config.get("style", "dark-blue")
                }
                dashboard_config["sections"].append(chart_section)
        
        # 3. Data tables section
        if "tables" in financial_data:
            for table_config in financial_data["tables"]:
                table_section = {
                    "title": table_config.get("title", "Financial Table"),
                    "type": "table",
                    "data_range": table_config.get("data_range", ""),
                    "name": table_config.get("name", "FinanceTable"),
                    "style": table_config.get("style", "TableStyleMedium9")
                }
                dashboard_config["sections"].append(table_section)
        
        # Create the dashboard
        dashboard_result = create_dashboard(wb, dashboard_config)
        
        # Add result
        result["sheets"].append({"name": sheet_name, "type": "dashboard"})
        result["dashboard"] = dashboard_result
    
    elif template_name == "sales_analysis":
        # Template for sales analysis
        title = data.get("title", "Sales Analysis")
        sheet_data = data.get("sales_data", [])
        sheet_name = data.get("sheet", "Sales")
        
        # Crear hoja de datos si no existe
        data_sheet = f"{sheet_name}_Datos"
        if data_sheet not in list_sheets(wb):
            data_ws = add_sheet(wb, data_sheet)
        else:
            data_ws = wb[data_sheet]
        
        # Escribir datos de ventas
        if sheet_data:
            write_sheet_data(data_ws, "A1", sheet_data)
            
            # Crear tabla para los datos
            rows = len(sheet_data)
            cols = max([len(row) if isinstance(row, list) else 1 for row in sheet_data], default=0)
            data_range = f"A1:{get_column_letter(cols)}{rows}"
            
            try:
                table = add_table(data_ws, "SalesDataTable", data_range, "TableStyleMedium9")
                result["elements"].append({
                    "type": "table",
                    "name": "SalesDataTable",
                    "sheet": data_sheet,
                    "range": data_range
                })
            except Exception as e:
                logger.warning(f"Error al crear tabla de datos: {e}")
        
        # Create analysis sheet
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Analysis title
        update_cell(ws, "A1", title)
        apply_style(ws, "A1", {
            "font_size": 16,
            "bold": True,
            "alignment": "center"
        })
        

            
        # Create analysis sections according to the data structure
        current_row = 3
        
        # 1. Sales by Region (assuming there is a region column)
        update_cell(ws, f"A{current_row}", "Sales by Region")
        apply_style(ws, f"A{current_row}", {"bold": True, "font_size": 12})
        current_row += 1
        
        try:
            # Create chart for sales by region
            chart_id, chart = add_chart(wb, sheet_name, "column",
                                       f"{data_sheet}!A1:{get_column_letter(cols)}{rows}",
                                       "Sales by Region", f"A{current_row}", "colorful-1")
            
            result["elements"].append({
                "type": "chart",
                "name": "SalesByRegionChart",
                "sheet": sheet_name,
                "id": chart_id
            })
            
            current_row += 15  # Space for the chart
        except Exception as e:
            logger.warning(f"Error creating sales by region chart: {e}")
            current_row += 2
        
        # 2. Sales Trend (if there is time data)
        update_cell(ws, f"A{current_row}", "Sales Trend")
        apply_style(ws, f"A{current_row}", {"bold": True, "font_size": 12})
        current_row += 1
        
        try:
            # Create chart for sales trend
            chart_id, chart = add_chart(wb, sheet_name, "line",
                                       f"{data_sheet}!A1:{get_column_letter(cols)}{rows}",
                                       "Sales Trend", f"A{current_row}", "line-markers")
            
            result["elements"].append({
                "type": "chart",
                "name": "SalesTrendChart",
                "sheet": sheet_name,
                "id": chart_id
            })
            
            current_row += 15  # Space for the chart
        except Exception as e:
            logger.warning(f"Error creating sales trend chart: {e}")
            current_row += 2
        
        result["sheets"].append({"name": sheet_name, "type": "analysis"})
        result["sheets"].append({"name": data_sheet, "type": "data"})
        
    elif template_name == "project_tracker":
        # Template for project tracking
        title = data.get("title", "Project Tracking")
        projects = data.get("projects", [])
        sheet_name = data.get("sheet", "Projects")
        
        # Prepare project data
        if not projects:
            # Create sample data if none is provided
            projects = [
                ["ID", "Project", "Owner", "Start", "Deadline", "Status", "Progress"],
                ["P001", "Project A", "Juan Pérez", "01/01/2023", "30/06/2023", "In progress", 75],
                ["P002", "Project B", "Ana López", "15/02/2023", "31/07/2023", "In progress", 40],
                ["P003", "Project C", "Carlos Ruiz", "01/03/2023", "31/08/2023", "Delayed", 20]
            ]
        
        # Create sheet for projects if it does not exist
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Tracker title
        update_cell(ws, "A1", title)
        apply_style(ws, "A1", {
            "font_size": 16,
            "bold": True,
            "alignment": "center"
        })
        

        
        # Write project data
        write_sheet_data(ws, "A3", projects)
        
        # Crear tabla para los datos
        rows = len(projects)
        cols = len(projects[0]) if rows > 0 else 7
        table_range = f"A3:{get_column_letter(cols)}{rows+2}"
        
        try:
            table = add_table(ws, "ProjectsTable", table_range, "TableStyleMedium9")
            result["elements"].append({
                "type": "table",
                "name": "ProjectsTable",
                "sheet": sheet_name,
                "range": table_range
            })
            
            # Apply percentage format to the progress column
            avance_col = get_column_letter(cols)
            apply_number_format(ws, f"{avance_col}4:{avance_col}{rows+2}", "0%")
        except Exception as e:
            logger.warning(f"Error al crear tabla de proyectos: {e}")
        
        # Create progress chart
        try:
            chart_id, chart = add_chart(wb, sheet_name, "column",
                                       table_range,
                                       "Project Progress", "I3", "colorful-3")
            
            result["elements"].append({
                "type": "chart",
                "name": "ProgressChart",
                "sheet": sheet_name,
                "id": chart_id
            })
        except Exception as e:
            logger.warning(f"Error creating progress chart: {e}")
        
        result["sheets"].append({"name": sheet_name, "type": "tracker"})
    
    else:
        logger.warning(f"Plantilla '{template_name}' no reconocida.")
        result["error"] = f"Plantilla '{template_name}' no disponible"
    
    return result

def update_report(wb: Any, report_config: Dict[str, Any],
                 recalculate: bool = True) -> Dict[str, Any]:
    """
    Update an existing report with new data.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        wb: Openpyxl workbook object.
        report_config: Configuration for the report update
            {
                "data_updates": {
                    "Sales": {"range": "A2:C10", "data": [[new data]]},
                    "Customers": {"range": "A2:D20", "data": [[new data]]}
                },
                "recalculate_formulas": True,
                "refresh_charts": True
            }
        recalculate: If ``True`` recalculate formulas after updating.

    Returns:
        Dictionary with information about the updated elements.
    """
    result = {
        "updated_sheets": [],
        "updated_tables": [],
        "updated_charts": [],
        "recalculated": recalculate
    }
    
    # Actualizar datos en hojas
    data_updates = report_config.get("data_updates", {})
    for sheet_name, update_info in data_updates.items():
        if sheet_name not in list_sheets(wb):
            logger.warning(f"Sheet '{sheet_name}' not found. Skipping update.")
            continue
        
        ws = wb[sheet_name]
        range_str = update_info.get("range")
        data = update_info.get("data")
        
        if not range_str or not data:
            logger.warning(f"Incomplete configuration to update sheet '{sheet_name}'. Range and data are required.")
            continue
        
        try:
            # Obtener solo la primera celda del rango
            if ':' in range_str:
                start_cell = range_str.split(':')[0]
            else:
                start_cell = range_str
            
            # Escribir nuevos datos
            write_sheet_data(ws, start_cell, data)
            
            result["updated_sheets"].append({
                "name": sheet_name,
                "range": range_str
            })
        except Exception as e:
            logger.warning(f"Error al actualizar datos en hoja '{sheet_name}': {e}")
    
    # Actualizar/refrescar tablas
    refresh_tables = report_config.get("refresh_tables", [])
    for table_info in refresh_tables:
        sheet_name = table_info.get("sheet")
        table_name = table_info.get("name")
        new_range = table_info.get("new_range")
        
        if not sheet_name or not table_name:
            logger.warning("Incomplete table information. Sheet and name are required.")
            continue
        
        if sheet_name not in list_sheets(wb):
            logger.warning(f"Sheet '{sheet_name}' not found. Skipping table update.")
            continue
        
        ws = wb[sheet_name]
        
        try:
            # Verificar si la tabla existe
            if not hasattr(ws, 'tables') or table_name not in ws.tables:
                logger.warning(f"Table '{table_name}' not found in sheet '{sheet_name}'.")
                continue
            
            # Get current reference
            current_range = ws.tables[table_name].ref
            
            # Update range if a new one is provided
            if new_range:
                ws.tables[table_name].ref = new_range
                
                result["updated_tables"].append({
                    "name": table_name,
                    "sheet": sheet_name,
                    "old_range": current_range,
                    "new_range": new_range
                })
            else:
                result["updated_tables"].append({
                    "name": table_name,
                    "sheet": sheet_name,
                    "refreshed": True
                })
        except Exception as e:
            logger.warning(f"Error updating table '{table_name}': {e}")
    
    # Recalculate formulas if requested
    if recalculate:
        # OpenPyXL does not directly recalculate formulas
        # This is a placeholder that could be implemented in future versions
        # or via Excel's COM API if available
        result["recalculation_note"] = "Formula recalculation in OpenPyXL is limited"
    
    # Update charts
    refresh_charts = report_config.get("refresh_charts", [])
    for chart_info in refresh_charts:
        sheet_name = chart_info.get("sheet")
        chart_id = chart_info.get("id")
        new_data_range = chart_info.get("new_data_range")
        
        if not sheet_name or chart_id is None:
            logger.warning("Incomplete chart information. Sheet and id are required.")
            continue
        
        if sheet_name not in list_sheets(wb):
            logger.warning(f"Sheet '{sheet_name}' not found. Skipping chart update.")
            continue
        
        ws = wb[sheet_name]
        
        try:
            # Verify if the chart exists
            if not hasattr(ws, '_charts') or chart_id >= len(ws._charts) or chart_id < 0:
                logger.warning(f"Chart with ID {chart_id} not found in sheet '{sheet_name}'.")
                continue
            
            # In OpenPyXL updating a chart is not straightforward
            # One option is to delete the chart and create a new one
            if new_data_range:
                # Get current chart properties
                chart_rel = ws._charts[chart_id]
                chart = chart_rel[0]
                position = chart_rel[1] if len(chart_rel) > 1 else None
                
                # Determine chart type
                chart_type = "column"  # Default value
                if isinstance(chart, BarChart):
                    chart_type = "bar" if chart.type == "bar" else "column"
                elif isinstance(chart, LineChart):
                    chart_type = "line"
                elif isinstance(chart, PieChart):
                    chart_type = "pie"
                elif isinstance(chart, ScatterChart):
                    chart_type = "scatter"
                elif isinstance(chart, AreaChart):
                    chart_type = "area"
                
                # Get title if it exists
                title = chart.title if hasattr(chart, 'title') and chart.title else None
                
                # Delete the old chart
                del ws._charts[chart_id]
                
                # Create a new chart with the same parameters but new range
                new_chart_id, new_chart = add_chart(wb, sheet_name, chart_type, new_data_range,
                                                 title, position)
                
                result["updated_charts"].append({
                    "id": chart_id,
                    "new_id": new_chart_id,
                    "sheet": sheet_name,
                    "old_data_range": "unknown",  # No easy way to get the original range
                    "new_data_range": new_data_range
                })
            else:
                # Without a new range it cannot be easily updated
                result["updated_charts"].append({
                    "id": chart_id,
                    "sheet": sheet_name,
                    "note": "No new range provided. Updating data requires Excel COM."
                })
        except Exception as e:
            logger.warning(f"Error updating chart {chart_id}: {e}")
    
    return result

def import_data(wb: Any, import_config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Import data from various sources into Excel.

    Args:
        wb: Openpyxl workbook object.
        import_config: Import configuration
            {
                "source": "csv",  # csv, json, pandas, etc.
                "source_path": "data.csv",
                "sheet": "Data",
                "start_cell": "A1",
                "options": {
                    "delimiter": ",",
                    "has_header": true
                }
            }

    Returns:
        Dictionary with information about the imported data.

    Note: This function is a simplified example that only imports data from CSV.
    """
    result = {
        "source": import_config.get("source"),
        "imported_rows": 0,
        "imported_columns": 0
    }
    
    source_type = import_config.get("source", "").lower()
    source_path = import_config.get("source_path")
    sheet_name = import_config.get("sheet", "Data")
    start_cell = import_config.get("start_cell", "A1")
    options = import_config.get("options", {})
    
    if not source_path:
        logger.warning("No source path specified for importing data.")
        result["error"] = "No source path specified"
        return result
        
    # Create the sheet if it does not exist
    if sheet_name not in list_sheets(wb):
        ws = add_sheet(wb, sheet_name)
    else:
        ws = wb[sheet_name]
    
    if source_type == "csv":
        try:
            import csv
            
            delimiter = options.get("delimiter", ",")
            has_header = options.get("has_header", True)
            
            data = []
            with open(source_path, 'r', encoding='utf-8', newline='') as f:
                csv_reader = csv.reader(f, delimiter=delimiter)
                for row in csv_reader:
                    data.append(row)
            
            # Write the data
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error importing CSV: {e}")
            result["error"] = f"Error importing CSV: {e}"
    
    elif source_type == "json":
        try:
            import json
            
            with open(source_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # Convert JSON to list of lists
            data = []
            
            if isinstance(json_data, list):
                # It is a list of objects
                if json_data and isinstance(json_data[0], dict):
                    # Get headers (claves del primer objeto)
                    headers = list(json_data[0].keys())
                    data.append(headers)
                    
                    # Add data rows
                    for item in json_data:
                        row = [item.get(header, "") for header in headers]
                        data.append(row)
                else:
                    # It is a simple list
                    for item in json_data:
                        data.append([item])
            elif isinstance(json_data, dict):
                # It is a dictionary
                for key, value in json_data.items():
                    data.append([key, value])
            
            # Write the data
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error importing JSON: {e}")
            result["error"] = f"Error importing JSON: {e}"
    
    elif source_type == "pandas":
        try:
            import pandas as pd
            
            # Options for pandas
            file_ext = os.path.splitext(source_path)[1].lower()
            
            if file_ext == '.csv':
                df = pd.read_csv(source_path)
            elif file_ext in ['.xls', '.xlsx']:
                df = pd.read_excel(source_path)
            elif file_ext == '.json':
                df = pd.read_json(source_path)
            else:
                raise ValueError(f"Unsupported file format: {file_ext}")
            
            # Convert DataFrame to list of lists
            data = [df.columns.tolist()]  # Encabezados
            data.extend(df.values.tolist())  # Datos
            
            # Write the data
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error importing with pandas: {e}")
            result["error"] = f"Error importing with pandas: {e}"
    
    else:
        logger.warning(f"Unsupported source type: {source_type}")
        result["error"] = f"Unsupported source type: {source_type}"
    
    return result

def export_data(wb: Any, export_config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Export data from Excel to different formats.

    Args:
        wb: Openpyxl workbook object.
        export_config: Export configuration
            {
                "format": "csv",  # csv, json, pdf, html, etc.
                "sheet": "Data",
                "range": "A1:D10",
                "output_path": "exported_data.csv",
                "options": {
                    "delimiter": ",",
                    "include_header": true
                }
            }

    Returns:
        Dictionary with information about the exported data.

    Note: This function is a simplified example that only exports to CSV and JSON.
    """
    result = {
        "format": export_config.get("format"),
        "exported_rows": 0,
        "exported_columns": 0
    }
    
    export_format = export_config.get("format", "").lower()
    sheet_name = export_config.get("sheet")
    range_str = export_config.get("range")
    output_path = export_config.get("output_path")
    options = export_config.get("options", {})
    
    if not sheet_name:
        logger.warning("No sheet specified para exportar datos.")
        result["error"] = "No sheet specified"
        return result
        
    if sheet_name not in list_sheets(wb):
        logger.warning(f"Hoja '{sheet_name}' no encontrada.")
        result["error"] = f"Hoja '{sheet_name}' no encontrada"
        return result
    
    # Leer los datos del rango especificado
    data = read_sheet_data(wb, sheet_name, range_str)
    
    if not data:
        logger.warning(f"No data found in range {range_str} de la hoja {sheet_name}")
        return []
    
    # Filter the data based on the criteria
    result = []
    headers = data[0] if data else []
    
    # If there is data with headers
    if len(data) > 1:
        # Convert to record format (list of dictionaries)
        records = []
        for row in data[1:]:
            record = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    record[header] = row[i]
                else:
                    record[header] = None
            records.append(record)
        
        # Apply filters if provided
        if filters:
            filtered_records = []
            for record in records:
                include = True
                for field, value in filters.items():
                    if field in record:
                        # If the filter value is a list, check if the value is in the list
                        if isinstance(value, list):
                            if record[field] not in value:
                                include = False
                                break
                        # If the filter value is a dictionary, apply operators
                        elif isinstance(value, dict):
                            for op, op_value in value.items():
                                if op == 'eq' and record[field] != op_value:
                                    include = False
                                    break
                                elif op == 'ne' and record[field] == op_value:
                                    include = False
                                    break
                                elif op == 'gt' and (not isinstance(record[field], (int, float)) or record[field] <= op_value):
                                    include = False
                                    break
                                elif op == 'lt' and (not isinstance(record[field], (int, float)) or record[field] >= op_value):
                                    include = False
                                    break
                                elif op == 'contains' and (not isinstance(record[field], str) or op_value not in record[field]):
                                    include = False
                                    break
                        # If the filter value is a simple value, perform an equality comparison
                        elif record[field] != value:
                            include = False
                            break
                if include:
                    filtered_records.append(record)
            records = filtered_records
        
        # Devolver los registros filtrados
        result = records
    
    return result

def create_report_from_template(template_file, output_file, data_mappings, chart_mappings=None, format_mappings=None):
    """
    Create a report based on an Excel template, replacing data, updating charts and applying formats.
     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        template_file (str): Path to the Excel template.

        output_file (str): Path where the generated report will be saved.
        data_mappings (dict): Dictionary with data mappings:
            {
                "sheet_name": {
                    "range1": data_list1,
                    "range2": data_list2,
                    ...
                }
            }
        chart_mappings (dict, optional): Dictionary with chart updates:
            {
                "sheet_name": {
                    "chart_id": {
                        "title": "New title",
                        "data_range": "New range",
                        ...
                    }
                }
            }
        format_mappings (dict, optional): Dictionary with formats to apply:
            {
                "sheet_name": {
                    "range1": {"number_format": "#,##0.00"},
                    "range2": {"style": {"bold": True, "fill_color": "FFFF00"}},
                    ...
                }
            }

    Returns:
        dict: Result of the operation
    """
    try:
        # Verificar que el archivo de plantilla existe
        if not os.path.exists(template_file):
            raise FileNotFoundError(f"La plantilla no existe: {template_file}")
        
        # Copiar la plantilla al archivo de salida
        import shutil
        shutil.copy2(template_file, output_file)
        
        # Abrir el nuevo archivo
        wb = openpyxl.load_workbook(output_file)
        
        # Aplicar mapeos de datos
        if data_mappings:
            for sheet_name, ranges in data_mappings.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"Sheet '{sheet_name}' does not exist in the template")
                    continue
                
                ws = wb[sheet_name]
                for range_str, data in ranges.items():
                    # Si el rango es una sola celda, extraer la celda de inicio
                    if ':' not in range_str:
                        start_cell = range_str
                    else:
                        start_cell = range_str.split(':')[0]
                    
                    # Write the data
                    write_sheet_data(ws, start_cell, data)
        
        # Apply chart mappings
        if chart_mappings:
            for sheet_name, charts in chart_mappings.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"Sheet '{sheet_name}' does not exist in the template")
                    continue
                
                ws = wb[sheet_name]
                existing_charts = list_charts(ws)
                
                for chart_id, chart_updates in charts.items():
                    # Check if chart_id is an index or a name
                    chart_idx = None
                    if isinstance(chart_id, int) or (isinstance(chart_id, str) and chart_id.isdigit()):
                        chart_idx = int(chart_id)
                    else:
                        # Look up the chart by title
                        for i, chart in enumerate(existing_charts):
                            if chart.get('title') == chart_id:
                                chart_idx = i
                                break
                    
                    if chart_idx is None or chart_idx >= len(existing_charts):
                        logger.warning(f"Chart not found '{chart_id}' en la hoja '{sheet_name}'")
                        continue
                    
                    # Update chart properties
                    chart = ws._charts[chart_idx][0]
                    
                    if 'title' in chart_updates:
                        chart.title = chart_updates['title']
                    
                    if 'data_range' in chart_updates:
                        # Updating the data range is complex and depends on the chart type
                        # For now just log that this feature is not implemented
                        logger.warning("Chart data range update is not fully implemented")
        
        # Aplicar mapeos de formato
        if format_mappings:
            for sheet_name, ranges in format_mappings.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"Sheet '{sheet_name}' does not exist in the template")
                    continue
                
                ws = wb[sheet_name]
                for range_str, formats in ranges.items():
                    if 'number_format' in formats:
                        apply_number_format(ws, range_str, formats['number_format'])
                    
                    if 'style' in formats:
                        apply_style(ws, range_str, formats['style'])
        
        # Guardar el archivo
        wb.save(output_file)
        
        return {
            "success": True,
            "template_file": template_file,
            "output_file": output_file,
            "message": f"Informe creado correctamente: {output_file}"
        }
    
    except Exception as e:
        logger.error(f"Error al crear informe desde plantilla: {e}")
        return {
            "success": False,
            "error": str(e),
            "message": f"Error al crear informe desde plantilla: {e}"
        }

def create_dynamic_dashboard(file_path, data, dashboard_config, overwrite=False):
    """
    Create a dynamic dashboard with multiple visualizations in a single step.

    Models using this function should ensure that tables and charts do not overlap.
     **Emojis must never be included in text written to cells, labels, titles or charts.**
    Leaving empty rows between sections and adjusting column widths after writing
    the data helps produce a clean, professional result.

    Args:
        file_path (str): Path to the Excel file to create or modify.
        data (dict): Dictionary with data per sheet:
            {
                "sheet_name": [
                    ["Header1", "Header2", ...],
                    [value1, value2, ...],
                    ...
                ]
            }
        dashboard_config (dict): Dashboard configuration:
            {
                "tables": [
                    {
                        "sheet": "sheet_name",
                        "name": "TableName",
                        "range": "A1:C10",
                        "style": "TableStyleMedium9"
                    }
                ],
                "charts": [
                    {
                        "sheet": "sheet_name",
                        "type": "column",
                        "data_range": "A1:C10",
                        "title": "Chart Title",
                        "position": "E1",
                        "style": "style1"
                    }
                ],
                "slicers": [
                    {
                        "sheet": "sheet_name",
                        "table": "TableName",
                        "column": "Category",
                        "position": "H1"
                    }
                ]
            }
        overwrite (bool): If ``True`` overwrite the file if it exists.

    Returns:
        dict: Result of the operation
    """
    try:
        # Verificar si el archivo existe
        file_exists = os.path.exists(file_path)
        
        if file_exists and not overwrite:
            raise FileExistsError(f"El archivo '{file_path}' ya existe. Use overwrite=True para sobrescribir.")
        
        # Crear o abrir el archivo
        if not file_exists or overwrite:
            wb = openpyxl.Workbook()
            # Eliminar la hoja predeterminada si existe
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
        else:
            wb = openpyxl.load_workbook(file_path)
        
        # Crear o actualizar hojas con datos
        for sheet_name, sheet_data in data.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.create_sheet(sheet_name)
            
            # Escribir datos
            if sheet_data:
                write_sheet_data(ws, "A1", sheet_data)
        
        # Crear tablas
        for table_config in dashboard_config.get("tables", []):
            sheet_name = table_config["sheet"]
            table_name = table_config["name"]
            range_str = table_config["range"]
            style = table_config.get("style", "TableStyleMedium9")
            
            if sheet_name not in wb.sheetnames:
                logger.warning(f"Sheet '{sheet_name}' does not exist to create table '{table_name}'")
                continue
            
            ws = wb[sheet_name]
            
            # Verificar si la tabla ya existe
            table_exists = False
            if hasattr(ws, 'tables') and table_name in ws.tables:
                table_exists = True
                logger.warning(f"Table '{table_name}' already exists, it will be updated")
            
            if table_exists:
                # Actualizar tabla existente
                refresh_table(ws, table_name, range_str)
            else:
                # Crear nueva tabla
                add_table(ws, table_name, range_str, style)
            
            # Aplicar formatos si se especifican
            if "formats" in table_config:
                for cell_range, fmt in table_config["formats"].items():
                    if isinstance(fmt, str):
                        # Numeric format
                        apply_number_format(ws, cell_range, fmt)
                    elif isinstance(fmt, dict):
                        # Es un estilo
                        apply_style(ws, cell_range, fmt)
        
        # Create charts
        for chart_config in dashboard_config.get("charts", []):
            sheet_name = chart_config["sheet"]
            chart_type = chart_config["type"]
            data_range = chart_config["data_range"]
            title = chart_config.get("title", f"Chart {len(wb[sheet_name]._charts) + 1}")
            position = chart_config.get("position", "E1")
            style = chart_config.get("style")
            
            if sheet_name not in wb.sheetnames:
                logger.warning(f"Sheet '{sheet_name}' does not exist to create the chart '{title}'")
                continue
            
            # Create chart
            chart_id, _ = add_chart(wb, sheet_name, chart_type, data_range, title, position, style)
        
        # Set column widths for optimal display
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Set a minimum width for date columns
            for i in range(1, ws.max_column + 1):
                column_letter = get_column_letter(i)
                # Verificar si hay celdas con formato de fecha en la columna
                date_format = False
                for cell in ws[column_letter]:
                    if cell.number_format and ('yy' in cell.number_format.lower() or 'mm' in cell.number_format.lower() or 'dd' in cell.number_format.lower()):
                        date_format = True
                        break
                
                if date_format:
                    # Set minimum width for date columns
                    ws.column_dimensions[column_letter].width = max(ws.column_dimensions[column_letter].width or 0, 10)
        
        # Guardar el archivo
        wb.save(file_path)
        
        return {
            "success": True,
            "file_path": file_path,
            "message": f"Dashboard creado correctamente: {file_path}"
        }
    
    except Exception as e:
        logger.error(f"Error creating dashboard: {e}")
        return {
            "success": False,
            "error": str(e),
            "message": f"Error creating dashboard: {e}"
        }

def import_multi_source_data(excel_file, import_config, sheet_name=None, start_cell="A1", create_tables=False):
    """
    Import data from multiple sources (CSV, JSON, SQL) into an Excel file in one step.

     **Emojis must never be included in text written to cells, labels, titles or charts.**

    Args:
        excel_file (str): Path to the Excel file where the data will be imported.
        import_config (dict): Import configuration:
            {
                "csv": [
                    {
                        "file_path": "data.csv",
                        "sheet_name": "SheetName",
                        "start_cell": "A1",
                        "delimiter": ",",
                        "encoding": "utf-8"
                    }
                ],
                "json": [
                    {
                        "file_path": "data.json",
                        "sheet_name": "SheetName",
                        "start_cell": "A1",
                        "fields": ["field1", "field2"]
                    }
                ],
                "sql": [
                    {
                        "query": "SELECT * FROM table",
                        "sheet_name": "SheetName",
                        "start_cell": "A1",
                        "connection_string": "..."
                    }
                ]
            }
        sheet_name (str, optional): Default sheet name if not specified in the configuration
        start_cell (str, optional): Default starting cell if not specified in the configuration
        create_tables (bool, optional): If True, create Excel tables for each dataset
    
    Returns:
        dict: Result of the operation
    """
    try:
        import csv
        import json
        
        # Try to import pandas if available (optional)
        try:
            import pandas as pd
            HAS_PANDAS = True
        except ImportError:
            HAS_PANDAS = False
            logger.warning("Pandas is not available. Some features will be limited.")
        
        # Verificar si el archivo Excel existe, si no, crearlo
        if not os.path.exists(excel_file):
            wb = openpyxl.Workbook()
            if sheet_name and "Sheet" in wb.sheetnames:
                # Renombrar la hoja predeterminada si se proporciona sheet_name
                wb["Sheet"].title = sheet_name
        else:
            wb = openpyxl.load_workbook(excel_file)
        
        imported_data = []
        
        # Procesar importaciones CSV
        for csv_config in import_config.get("csv", []):
            csv_file = csv_config["file_path"]
            csv_sheet = csv_config.get("sheet_name", sheet_name)
            csv_cell = csv_config.get("start_cell", start_cell)
            delimiter = csv_config.get("delimiter", ",")
            encoding = csv_config.get("encoding", "utf-8")
            
            if not os.path.exists(csv_file):
                logger.warning(f"El archivo CSV no existe: {csv_file}")
                continue
            
            # Crear la hoja si no existe
            if csv_sheet not in wb.sheetnames:
                ws = wb.create_sheet(csv_sheet)
            else:
                ws = wb[csv_sheet]
            
            # Leer datos CSV
            if HAS_PANDAS:
                # Use pandas if available
                df = pd.read_csv(csv_file, delimiter=delimiter, encoding=encoding)
                data = [df.columns.tolist()]  # Encabezados
                data.extend(df.values.tolist())  # Datos
            else:
                # Use standard csv if pandas is not available
                data = []
                with open(csv_file, 'r', encoding=encoding) as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        data.append(row)
            
            # Escribir datos en la hoja
            write_sheet_data(ws, csv_cell, data)
            
            # Crear tabla si se solicita
            if create_tables:
                # Determinar el rango de la tabla
                start_row, start_col = ExcelRange.parse_cell_ref(csv_cell)
                end_row = start_row + len(data) - 1
                end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
                table_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
                
                # Create a unique name for the table
                table_name = f"Table_{csv_sheet}_{len(imported_data) + 1}"
                table_name = table_name.replace(" ", "_")
                
                try:
                    add_table(ws, table_name, table_range, "TableStyleMedium9")
                except Exception as table_error:
                    logger.warning(f"Could not create the table for {csv_file}: {table_error}")
            
            imported_data.append({
                "source": "csv",
                "file": csv_file,
                "sheet": csv_sheet,
                "rows": len(data)
            })
        
        # Procesar importaciones JSON
        for json_config in import_config.get("json", []):
            json_file = json_config["file_path"]
            json_sheet = json_config.get("sheet_name", sheet_name)
            json_cell = json_config.get("start_cell", start_cell)
            fields = json_config.get("fields", [])
            
            if not os.path.exists(json_file):
                logger.warning(f"El archivo JSON no existe: {json_file}")
                continue
            
            # Crear la hoja si no existe
            if json_sheet not in wb.sheetnames:
                ws = wb.create_sheet(json_sheet)
            else:
                ws = wb[json_sheet]
            
            # Leer datos JSON
            with open(json_file, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # Convertir JSON a formato tabular
            if isinstance(json_data, list):
                # Si es una lista de objetos, extraer los campos
                if fields:
                    # Usar campos especificados
                    headers = fields
                elif json_data and isinstance(json_data[0], dict):
                    # Usar todas las claves del primer objeto
                    headers = list(json_data[0].keys())
                else:
                    headers = []
                
                # Crear datos tabulares
                data = [headers]
                for item in json_data:
                    if isinstance(item, dict):
                        row = [item.get(field, "") for field in headers]
                        data.append(row)
                    else:
                        # If the item is not a dictionary, add it as a single column
                        data.append([item])
            else:
                # Si es un solo objeto, usar sus claves y valores
                if isinstance(json_data, dict):
                    if fields:
                        # Usar campos especificados
                        headers = fields
                        data = [headers, [json_data.get(field, "") for field in headers]]
                    else:
                        # Usar todas las claves
                        headers = list(json_data.keys())
                        data = [headers, list(json_data.values())]
                else:
                    # If it is neither a dictionary nor a list, use a simple representation
                    data = [["Value"], [json_data]]
            
            # Escribir datos en la hoja
            write_sheet_data(ws, json_cell, data)
            
            # Crear tabla si se solicita
            if create_tables and data:
                # Determinar el rango de la tabla
                start_row, start_col = ExcelRange.parse_cell_ref(json_cell)
                end_row = start_row + len(data) - 1
                end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
                table_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
                
                # Create a unique name for the table
                table_name = f"Table_{json_sheet}_{len(imported_data) + 1}"
                table_name = table_name.replace(" ", "_")
                
                try:
                    add_table(ws, table_name, table_range, "TableStyleMedium9")
                except Exception as table_error:
                    logger.warning(f"Could not create the table for {json_file}: {table_error}")
            
            imported_data.append({
                "source": "json",
                "file": json_file,
                "sheet": json_sheet,
                "rows": len(data)
            })
        
        # Process SQL queries (requires database connection)
        if "sql" in import_config and import_config["sql"]:
            try:
                import pyodbc
                HAS_PYODBC = True
            except ImportError:
                HAS_PYODBC = False
                logger.warning("pyodbc is not available. SQL data cannot be imported.")
            
            if HAS_PYODBC or HAS_PANDAS:
                for sql_config in import_config.get("sql", []):
                    query = sql_config["query"]
                    sql_sheet = sql_config.get("sheet_name", sheet_name)
                    sql_cell = sql_config.get("start_cell", start_cell)
                    connection_string = sql_config.get("connection_string", "")
                    
                    if not query or not connection_string:
                        logger.warning("Se requiere query y connection_string para importar datos SQL")
                        continue
                    
                    # Crear la hoja si no existe
                    if sql_sheet not in wb.sheetnames:
                        ws = wb.create_sheet(sql_sheet)
                    else:
                        ws = wb[sql_sheet]
                    
                    try:
                        data = []
                        
                        if HAS_PANDAS:
                            # Use pandas if available
                            import urllib.parse
                            params = urllib.parse.quote_plus(connection_string)
                            connection_url = f"mssql+pyodbc:///?odbc_connect={params}"
                            
                            from sqlalchemy import create_engine
                            engine = create_engine(connection_url)
                            df = pd.read_sql(query, engine)
                            
                            data = [df.columns.tolist()]  # Encabezados
                            data.extend(df.values.tolist())  # Datos
                        else:
                            # Usar pyodbc directamente
                            conn = pyodbc.connect(connection_string)
                            cursor = conn.cursor()
                            cursor.execute(query)
                            
                            # Obtener nombres de columnas
                            columns = [column[0] for column in cursor.description]
                            data.append(columns)
                            
                            # Obtener datos
                            for row in cursor.fetchall():
                                data.append(list(row))
                            
                            conn.close()
                        
                        # Escribir datos en la hoja
                        write_sheet_data(ws, sql_cell, data)
                        
                        # Crear tabla si se solicita
                        if create_tables and data:
                            # Determinar el rango de la tabla
                            start_row, start_col = ExcelRange.parse_cell_ref(sql_cell)
                            end_row = start_row + len(data) - 1
                            end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
                            table_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
                            
                            # Create a unique name for the table
                            table_name = f"Table_{sql_sheet}_{len(imported_data) + 1}"
                            table_name = table_name.replace(" ", "_")
                            
                            try:
                                add_table(ws, table_name, table_range, "TableStyleMedium9")
                            except Exception as table_error:
                                logger.warning(f"Could not create the table for SQL query: {table_error}")
                        
                        imported_data.append({
                            "source": "sql",
                            "query": query[:50] + "..." if len(query) > 50 else query,
                            "sheet": sql_sheet,
                            "rows": len(data)
                        })
                    
                    except Exception as sql_error:
                        logger.error(f"Error al importar datos SQL: {sql_error}")
                        continue
        
        # Guardar el archivo Excel
        wb.save(excel_file)
        
        return {
            "success": True,
            "file_path": excel_file,
            "imported_data": imported_data,
            "message": f"Datos importados correctamente a {excel_file}"
        }
    
    except Exception as e:
        logger.error(f"Error al importar datos: {e}")
        return {
            "success": False,
            "error": str(e),
            "message": f"Error al importar datos: {e}"
        }

def export_excel_data(excel_file, export_config):
    """
    Export Excel data to multiple formats (CSV, JSON, PDF) in one step.
    
    Args:
        excel_file (str): Path to the source Excel file
        export_config (dict): Export configuration:
            {
                "csv": [
                    {
                        "sheet_name": "SheetName",
                        "range": "A1:C10",
                        "output_file": "output.csv",
                        "delimiter": ",",
                        "encoding": "utf-8"
                    }
                ],
                "json": [
                    {
                        "sheet_name": "SheetName",
                        "range": "A1:C10",
                        "output_file": "output.json",
                        "format": "records"  # "records", "object", "table"
                    }
                ],
                "pdf": {
                      "output_file": "output.pdf",
                      "sheets": ["Sheet1", "Sheet2"]  # or null for all
                }
            }
    
    Returns:
        dict: Result of the operation
    """
    try:
        import csv
        import json
        
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"El archivo Excel no existe: {excel_file}")
        
        # Cargar el archivo Excel
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        
        exported_files = []
        
        # Exportar a CSV
        for csv_config in export_config.get("csv", []):
            sheet_name = csv_config["sheet_name"]
            range_str = csv_config.get("range")
            output_file = csv_config["output_file"]
            delimiter = csv_config.get("delimiter", ",")
            encoding = csv_config.get("encoding", "utf-8")
            
            if sheet_name not in wb.sheetnames:
                logger.warning(f"La hoja '{sheet_name}' no existe")
                continue
            
            # Leer los datos del rango especificado
            data = read_sheet_data(wb, sheet_name, range_str)
            
            # Write the data en CSV
            with open(output_file, 'w', newline='', encoding=encoding) as csvfile:
                writer = csv.writer(csvfile, delimiter=delimiter)
                for row in data:
                    writer.writerow(row)
            
            exported_files.append({
                "format": "csv",
                "file": output_file,
                "sheet": sheet_name,
                "rows": len(data)
            })
        
        # Exportar a JSON
        for json_config in export_config.get("json", []):
            sheet_name = json_config["sheet_name"]
            range_str = json_config.get("range")
            output_file = json_config["output_file"]
            format_type = json_config.get("format", "records")
            
            if sheet_name not in wb.sheetnames:
                logger.warning(f"La hoja '{sheet_name}' no existe")
                continue
            
            # Leer los datos del rango especificado
            data = read_sheet_data(wb, sheet_name, range_str)
            
            if not data:
                logger.warning(f"No hay datos para exportar en la hoja '{sheet_name}'")
                continue
            
            # Convert data to JSON format according to the specified type
            headers = data[0]
            json_data = None
            
            if format_type == "records":
                # Formato de registros [{campo1: valor1, campo2: valor2}, {...}]
                json_data = []
                for row in data[1:]:
                    record = {}
                    for i, header in enumerate(headers):
                        if i < len(row):
                            record[header] = row[i]
                    json_data.append(record)
            
            elif format_type == "object":
                # Formato de objeto {id1: {campo1: valor1}, id2: {campo1: valor2}}
                json_data = {}
                id_field = headers[0]  # Usar la primera columna como ID
                for row in data[1:]:
                    if not row:
                        continue
                    record = {}
                    for i, header in enumerate(headers[1:], 1):  # Empezar desde la segunda columna
                        if i < len(row):
                            record[header] = row[i]
                    json_data[row[0]] = record
            
            elif format_type == "table":
                # Formato de tabla {headers: [...], data: [[...], [...]]}
                json_data = {
                    "headers": headers,
                    "data": [row for row in data[1:]]
                }
            
            # Write the data en JSON
            with open(output_file, 'w', encoding='utf-8') as jsonfile:
                json.dump(json_data, jsonfile, indent=2)
            
            exported_files.append({
                "format": "json",
                "file": output_file,
                "sheet": sheet_name,
                "rows": len(data) - 1  # Sin contar encabezados
            })
        
        # Exportar a PDF (requiere biblioteca adicional)
        if "pdf" in export_config:
            pdf_config = export_config["pdf"]
            output_file = pdf_config["output_file"]
            sheets = pdf_config.get("sheets")
            
            try:
                # Try to use win32com for Excel if available
                import win32com.client
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                
                # Abrir el archivo
                workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
                
                # Determinar las hojas a exportar
                sheets_to_export = []
                if sheets:
                    for sheet_name in sheets:
                        try:
                            sheet = workbook.Sheets(sheet_name)
                            sheets_to_export.append(sheet)
                        except:
                            logger.warning(f"La hoja '{sheet_name}' no existe para exportar a PDF")
                else:
                    # Exportar todas las hojas
                    sheets_to_export = workbook.Sheets
                
                # Exportar a PDF
                if sheets_to_export:
                    workbook.ExportAsFixedFormat(0, os.path.abspath(output_file))
                    
                    exported_files.append({
                        "format": "pdf",
                        "file": output_file,
                        "sheets": sheets if sheets else [sheet.Name for sheet in sheets_to_export]
                    })
                
                # Cerrar Excel
                workbook.Close(False)
                excel.Quit()
            
            except ImportError:
                logger.warning("win32com is not available. Cannot export to PDF.")
                pass  # If win32com is not available, simply skip the PDF export
            except Exception as pdf_error:
                logger.error(f"Error al exportar a PDF: {pdf_error}")
                pass
        
        return {
            "success": True,
            "file_path": excel_file,
            "exported_files": exported_files,
            "message": f"Datos exportados correctamente desde {excel_file}"
        }
    
    except Exception as e:
        logger.error(f"Error al exportar datos: {e}")
        return {
            "success": False,
            "error": str(e),
            "message": f"Error al exportar datos: {e}"
        }

def export_single_visible_sheet_pdf(excel_file: str, output_pdf: Optional[str] = None) -> Dict[str, Any]:
    """Export an Excel workbook to PDF only if it has a single visible sheet.

    Args:
        excel_file: Path to the Excel file to export.
        output_pdf: Path of the resulting PDF file. If not provided, the same
            name as ``excel_file`` with ``.pdf`` extension is used.

    Returns:
        dict: Result of the operation.
    """
    try:
        import shutil
        import subprocess

        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"El archivo Excel no existe: {excel_file}")

        wb = openpyxl.load_workbook(excel_file, data_only=True)
        visible_sheets = [ws.title for ws in wb.worksheets if getattr(ws, "sheet_state", "visible") == "visible"]

        if len(visible_sheets) != 1:
            msg = f"The file must have a single visible sheet. Visible sheets: {len(visible_sheets)}"
            logger.warning(msg)
            return {
                "success": False,
                "file_path": excel_file,
                "visible_sheets": visible_sheets,
                "message": msg,
            }

        if not output_pdf:
            output_pdf = os.path.splitext(excel_file)[0] + ".pdf"
        output_pdf = os.path.abspath(output_pdf)

        # Intentar exportar con win32com (Windows)
        try:
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
            workbook.ExportAsFixedFormat(0, output_pdf)
            workbook.Close(False)
            excel.Quit()

            msg = f"File successfully exported to PDF: {output_pdf}"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_file": output_pdf,
                "message": msg,
            }
        except ImportError:
            logger.info("win32com not available, LibreOffice will be tried")
        except Exception as e:
            logger.error(f"Error al exportar con win32com: {e}")

        # Fallback a LibreOffice en sistemas no Windows
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            outdir = os.path.dirname(output_pdf)
            cmd = [soffice, "--headless", "--convert-to", "pdf", os.path.abspath(excel_file), "--outdir", outdir]
            subprocess.run(cmd, check=True)

            generated = os.path.join(outdir, Path(excel_file).stem + ".pdf")
            if generated != output_pdf:
                os.replace(generated, output_pdf)

            msg = f"Archivo exportado correctamente a PDF: {output_pdf}"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_file": output_pdf,
                "message": msg,
            }

        msg = "No available method found to export to PDF."
        logger.error(msg)
        return {
            "success": False,
            "file_path": excel_file,
            "message": msg,
        }

    except Exception as e:
        logger.error(f"Error al exportar a PDF: {e}")
        return {
            "success": False,
            "file_path": excel_file,
            "error": str(e),
            "message": f"Error al exportar a PDF: {e}",
        }

def export_sheets_to_pdf(
    excel_file: str,
    sheets: Optional[Union[str, List[str]]] = None,
    output_dir: Optional[str] = None,
    single_file: bool = False,
) -> Dict[str, Any]:
    """Export one or more sheets of an Excel workbook to PDF.

    Parameters
    ----------
    excel_file : str
        Path to the Excel file to export.
    sheets : Union[str, List[str]], optional
        Name of the sheet or list of sheets to export. If ``None`` all sheets
        in the workbook are exported one by one.
    output_dir : str, optional
        Folder where the PDFs will be stored. By default the original file
        directory is used.
    single_file : bool, optional
        If ``True`` and several sheets are specified a single PDF is generated
        with all of them if supported. If ``False`` a PDF is created per sheet.

    Returns
    -------
    dict
        Operation result with the list of generated PDFs. If any sheet does not
        exist a warning is included in ``warnings``.
    """

    try:
        import shutil
        import subprocess

        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"El archivo Excel no existe: {excel_file}")

        wb = openpyxl.load_workbook(excel_file, data_only=True)
        all_sheets = wb.sheetnames
        wb.close()

        if sheets is None:
            target_sheets = all_sheets
        elif isinstance(sheets, str):
            target_sheets = [sheets]
        else:
            target_sheets = list(sheets)

        warnings = []
        valid_sheets = []
        for s in target_sheets:
            if s in all_sheets:
                valid_sheets.append(s)
            else:
                warnings.append(f"La hoja '{s}' no existe")

        if not valid_sheets:
            msg = "No valid sheets found to export"
            logger.warning(msg)
            return {
                "success": False,
                "file_path": excel_file,
                "warnings": warnings,
                "message": msg,
            }

        if output_dir is None:
            output_dir = os.path.dirname(os.path.abspath(excel_file))

        pdf_files: List[str] = []

        # Try to use win32com if available
        try:
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(os.path.abspath(excel_file))

            if single_file and len(valid_sheets) > 1:
                workbook.Worksheets(valid_sheets).Select()
                output_pdf = os.path.join(
                    output_dir, Path(excel_file).stem + ".pdf"
                )
                workbook.ActiveSheet.ExportAsFixedFormat(0, output_pdf)
                pdf_files.append(output_pdf)
            else:
                for s in valid_sheets:
                    ws = workbook.Worksheets(s)
                    output_pdf = os.path.join(
                        output_dir, f"{Path(excel_file).stem}_{s}.pdf"
                    )
                    ws.ExportAsFixedFormat(0, output_pdf)
                    pdf_files.append(output_pdf)

            workbook.Close(False)
            excel.Quit()

            msg = "PDF export completed successfully"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_files": pdf_files,
                "warnings": warnings,
                "message": msg,
            }
        except ImportError:
            logger.info("win32com not available, trying LibreOffice")
        except Exception as e:
            logger.error(f"Error al exportar con win32com: {e}")

        # Fallback a LibreOffice
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            with tempfile.TemporaryDirectory() as tmpdir:
                if single_file and len(valid_sheets) > 1:
                    tmp_xlsx = os.path.join(tmpdir, "tmp.xlsx")
                    wb = openpyxl.load_workbook(excel_file)
                    for sheet in wb.sheetnames:
                        wb[sheet].sheet_state = (
                            "visible" if sheet in valid_sheets else "hidden"
                        )
                    wb.save(tmp_xlsx)
                    wb.close()
                    cmd = [
                        soffice,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        os.path.abspath(tmp_xlsx),
                        "--outdir",
                        tmpdir,
                    ]
                    subprocess.run(cmd, check=True)
                    generated = os.path.join(tmpdir, "tmp.pdf")
                    final = os.path.join(
                        output_dir, Path(excel_file).stem + ".pdf"
                    )
                    shutil.move(generated, final)
                    pdf_files.append(final)
                else:
                    for s in valid_sheets:
                        tmp_xlsx = os.path.join(tmpdir, f"{s}.xlsx")
                        wb = openpyxl.load_workbook(excel_file)
                        for sheet in wb.sheetnames:
                            wb[sheet].sheet_state = (
                                "visible" if sheet == s else "hidden"
                            )
                        wb.save(tmp_xlsx)
                        wb.close()
                        cmd = [
                            soffice,
                            "--headless",
                            "--convert-to",
                            "pdf",
                            os.path.abspath(tmp_xlsx),
                            "--outdir",
                            tmpdir,
                        ]
                        subprocess.run(cmd, check=True)
                        generated = os.path.join(tmpdir, f"{s}.pdf")
                        final = os.path.join(
                            output_dir, f"{Path(excel_file).stem}_{s}.pdf"
                        )
                        shutil.move(generated, final)
                        pdf_files.append(final)

            msg = "PDF export completed successfully"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_files": pdf_files,
                "warnings": warnings,
                "message": msg,
            }

        msg = "No available method found to export to PDF."
        logger.error(msg)
        return {
            "success": False,
            "file_path": excel_file,
            "warnings": warnings,
            "message": msg,
        }

    except Exception as e:
        logger.error(f"Error al exportar a PDF: {e}")
        return {
            "success": False,
            "file_path": excel_file,
            "error": str(e),
            "message": f"Error al exportar a PDF: {e}",
        }

# Crear el servidor MCP como variable global
mcp = None
if HAS_MCP:
    # Esta es la variable global que el sistema MCP busca
    mcp = FastMCP("Master Excel MCP", 
                 dependencies=["openpyxl", "pandas", "numpy"])
    logger.info("Servidor MCP unificado iniciado correctamente")
    
    # Register basic workbook management functions
    @mcp.tool(description="Creates a new empty Excel file")
    def create_workbook_tool(filename, overwrite=False):
        """Create a new empty Excel workbook.

        This function creates an empty ``.xlsx`` file at the specified location.
        It is the recommended first step when generating a new document from scratch.

        Args:
            filename (str): Full path and name of the file to create. Must have a ``.xlsx`` extension.
            overwrite (bool, optional): If ``True`` overwrite the file if it already exists. Default is ``False``.

        Returns:
            dict: Information about the operation result including the created file path.

        Example:
            create_workbook_tool("C:/data/new_book.xlsx")
        """
        try:
            wb = create_workbook(filename, overwrite)
            save_workbook(wb, filename)
            
            return {
                "success": True,
                "file_path": filename,
                "message": f"Excel file successfully created: {filename}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating Excel file: {e}"
            }
    
    @mcp.tool(description="Abre un fichero Excel existente")
    def open_workbook_tool(filename):
        """Open an existing Excel file.

        This function opens an existing ``.xlsx`` or ``.xls`` file so it can be manipulated.
        Use this before performing any operation on an existing file.

        Args:
            filename (str): Full path and name of the Excel file to open.

        Returns:
            dict: Information about the opened file including sheet count and other properties.

        Raises:
            FileNotFoundError: If the specified file does not exist.

        Example:
            open_workbook_tool("C:/data/sales_report.xlsx")
        """
        try:
            wb = open_workbook(filename)
            sheet_names = list_sheets(wb)
            close_workbook(wb)
            
            return {
                "success": True,
                "file_path": filename,
                "sheets": sheet_names,
                "sheet_count": len(sheet_names),
                "message": f"Excel file successfully opened: {filename}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error opening Excel file: {e}"
            }
    
    @mcp.tool(description="Guarda el Workbook en disco")
    def save_workbook_tool(filename, new_filename=None):
        """Save the workbook to disk.

        Use this function after modifying a workbook to persist the changes.

        Args:
            filename (str): Full path and name of the Excel file to save.
            new_filename (str, optional): If provided, save the file under a new name ("Save As"). Defaults to ``None``.

        Returns:
            dict: Information about the operation result including the saved file path.

        Example:
            save_workbook_tool("C:/data/report.xlsx")
            save_workbook_tool("C:/data/report.xlsx", "C:/data/report_backup.xlsx")  # Save As
        """
        try:
            wb = open_workbook(filename)
            saved_path = save_workbook(wb, new_filename or filename)
            close_workbook(wb)
            
            return {
                "success": True,
                "original_file": filename,
                "saved_file": saved_path,
                "message": f"Excel file successfully saved: {saved_path}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error saving Excel file: {e}"
            }
    
    @mcp.tool(description="Lista las hojas disponibles en un archivo Excel")
    def list_sheets_tool(filename):
        """List the worksheets available in an Excel file.

        This function returns all worksheets contained in an Excel workbook and is useful
        to get an overview before working with the file.

        Args:
            filename (str): Full path and name of the Excel file to inspect.

        Returns:
            dict: Dictionary with the sheet names and their positions in the workbook.

        Raises:
            FileNotFoundError: If the specified file does not exist.

        Example:
            list_sheets_tool("C:/data/financial_report.xlsx")  # Returns: {"sheets": ["Sales", "Costs", "Summary"]}
        """
        try:
            wb = open_workbook(filename)
            sheets = list_sheets(wb)
            close_workbook(wb)
            
            return {
                "success": True,
                "file_path": filename,
                "sheets": sheets,
                "count": len(sheets),
                "message": f"Se encontraron {len(sheets)} hojas en el archivo Excel"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al listar hojas: {e}"
            }
    
    # Register basic worksheet manipulation functions
    @mcp.tool(description="Adds a new empty sheet")
    def add_sheet_tool(filename, sheet_name, index=None):
        """Add a new empty worksheet.

        This function inserts a new blank worksheet into an existing Excel workbook.
        You can specify the position where the sheet should be inserted.

        Args:
            filename (str): Full path and name of the Excel file.
            sheet_name (str): Name for the new sheet.
            index (int, optional): Position where the sheet will be inserted (``0`` is the first position).
                                 If ``None`` the sheet is added at the end. Default is ``None``.

        Returns:
            dict: Information about the operation including the updated list of sheets.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetExistsError: If a sheet with the same name already exists.

        Example:
            add_sheet_tool("C:/data/report.xlsx", "New Summary")  # Add at the end
            add_sheet_tool("C:/data/report.xlsx", "Cover", 0)  # Add as first sheet
        """
        try:
            wb = open_workbook(filename)
            ws = add_sheet(wb, sheet_name, index)
            save_workbook(wb, filename)
            
            sheets = list_sheets(wb)
            close_workbook(wb)
            
            return {
                "success": True,
                "file_path": filename,
                "sheet_name": sheet_name,
                "sheet_index": sheets.index(sheet_name),
                "all_sheets": sheets,
                "message": f"Sheet '{sheet_name}' added successfully"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error adding sheet: {e}"
            }
    
    @mcp.tool(description="Delete the indicated sheet")
    def delete_sheet_tool(filename, sheet_name):
        """Delete the indicated worksheet.

        This function removes a specific worksheet from an Excel workbook. Use with care
        because once the file is saved the deletion is permanent.

        Args:
            filename (str): Full path and name of the Excel file.
            sheet_name (str): Name of the worksheet to delete.

        Returns:
            dict: Information about the operation including the updated list of sheets.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If the specified sheet does not exist.
            ValueError: If attempting to delete the only sheet in the workbook.

        Example:
            delete_sheet_tool("C:/data/report.xlsx", "Draft")
        """
        try:
            wb = open_workbook(filename)
            delete_sheet(wb, sheet_name)
            save_workbook(wb, filename)
            
            remaining_sheets = list_sheets(wb)
            close_workbook(wb)
            
            return {
                "success": True,
                "file_path": filename,
                "deleted_sheet": sheet_name,
                "remaining_sheets": remaining_sheets,
                "remaining_count": len(remaining_sheets),
                "message": f"Sheet '{sheet_name}' successfully deleted"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error deleting sheet: {e}"
            }
    
    @mcp.tool(description="Rename a sheet")
    def rename_sheet_tool(filename, old_name, new_name):
        """Rename a worksheet.

        This function changes the name of an existing worksheet in an Excel workbook.

        Args:
            filename (str): Full path and name of the Excel file.
            old_name (str): Current name of the sheet to rename.
            new_name (str): New name for the sheet.

        Returns:
            dict: Information about the operation including the updated list of sheets.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If no sheet exists with the original name.
            SheetExistsError: If a sheet with the new name already exists.

        Example:
            rename_sheet_tool("C:/data/report.xlsx", "Sheet1", "Executive Summary")
        """
        try:
            wb = open_workbook(filename)
            rename_sheet(wb, old_name, new_name)
            save_workbook(wb, filename)
            
            sheets = list_sheets(wb)
            close_workbook(wb)
            
            return {
                "success": True,
                "file_path": filename,
                "old_name": old_name,
                "new_name": new_name,
                "all_sheets": sheets,
                "message": f"Sheet renamed from '{old_name}' to '{new_name}'"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error renaming sheet: {e}"
            }
    
    # Register basic writing functions
    @mcp.tool(description="Write a two-dimensional array of values or formulas")
    def write_sheet_data_tool(file_path, sheet_name, start_cell, data):
        """Write a two-dimensional array of values or formulas to a worksheet.

        This function writes data to a range of cells starting at ``start_cell``.
        It is ideal for inserting tables of data or matrices of values.

        Args:
            file_path (str): Full path and name of the Excel file.
            sheet_name (str): Name of the worksheet where data will be written.
            start_cell (str): Starting cell (e.g. ``"A1"``).
            data (list): Two-dimensional list with the data to write.
                        Example: [["Name", "Age"], ["John", 25], ["Mary", 30]]

        Returns:
            dict: Information about the operation, including the modified range.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If the specified sheet does not exist.
            CellReferenceError: If the cell reference is not valid.

        Example:
            write_sheet_data_tool(
                "C:/data/report.xlsx",
                "Data",
                "B2",
                [["Quarter", "Sales", "Costs"], ["Q1", 5000, 3000], ["Q2", 6200, 3100]]
            )
        """
        try:
            # Validate arguments
            if not isinstance(data, list):
                raise ValueError("The 'data' parameter must be a list")

            # Open the file and get the sheet
            wb = openpyxl.load_workbook(file_path)
            ws = get_sheet(wb, sheet_name)

            # Write the data
            write_sheet_data(ws, start_cell, data)

            # Save and close
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "start_cell": start_cell,
                "rows_written": len(data),
                "columns_written": max([len(row) if isinstance(row, list) else 1 for row in data], default=0),
                "message": f"Data successfully written starting at {start_cell}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error writing data: {e}"
            }
    
    @mcp.tool(description="Update a single cell")
    def update_cell_tool(file_path, sheet_name, cell, value_or_formula):
        """Update the value or formula of a specific cell.

        This function modifies a single cell in a worksheet. It can be used for both values and formulas.

        Args:
            file_path (str): Full path and name of the Excel file.
            sheet_name (str): Name of the sheet containing the cell to update.
            cell (str): Reference of the cell to update (e.g. ``"B5"``).
            value_or_formula (str | int | float | bool): Value or formula to set. Formulas must start with ``=``.

        Returns:
            dict: Information about the operation including the modified cell.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If the specified sheet does not exist.
            CellReferenceError: If the cell reference is not valid.

        Example:
            update_cell_tool("C:/data/report.xlsx", "Sales", "C4", 5280.50)  # Numeric value
            update_cell_tool("C:/data/report.xlsx", "Sales", "D4", "=SUM(A1:A10)")  # Formula
        """
        try:
            # Open the file and get the sheet
            wb = openpyxl.load_workbook(file_path)
            ws = get_sheet(wb, sheet_name)
            
            # Update the cell
            update_cell(ws, cell, value_or_formula)
            
            # Save and close
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "cell": cell,
                "value": value_or_formula,
                "message": f"Cell {cell} successfully updated in sheet {sheet_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error updating cell: {e}"
            }
    
    # Register advanced functions
    @mcp.tool(description="Define a range as a formatted table on an Excel sheet")
    def add_table_tool(file_path, sheet_name, table_name, cell_range, style=None):
        """Define a range as a formatted table in Excel.

        This function converts a cell range into an Excel table with formatting so the
        data can be filtered and sorted automatically.

        Args:
            file_path (str): Full path and name of the Excel file.
            sheet_name (str): Name of the sheet where the table will be created.
            table_name (str): Name for the table (must be unique within the workbook).
            cell_range (str): Cell range for the table in Excel format (e.g. ``"A1:D10"``).
            style (str, optional): Table style to apply (e.g. ``"TableStyleMedium9"``). If ``None`` the default style is used.

        Returns:
            dict: Information about the operation including table details.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If the specified sheet does not exist.
            RangeError: If the provided range is not valid.
            TableError: If a table with the same name already exists or another table issue occurs.

        Example:
            add_table_tool(
                "C:/data/sales.xlsx",
                "Data",
                "RegionalPrices",
                "B3:F15",
                "TableStyleMedium2"
            )
        """
        try:
            # Open the file
            wb = openpyxl.load_workbook(file_path)
            
            # Get the sheet
            ws = get_sheet(wb, sheet_name)
            
            # Add the table
            table = add_table(ws, table_name, cell_range, style)
            
            # Save changes
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "table_name": table_name,
                "range": cell_range,
                "style": style,
                "message": f"Table '{table_name}' successfully created in range {cell_range}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating table: {e}"
            }
    
    @mcp.tool(description="Insert a native chart into an Excel sheet with multiple customization options")
    def add_chart_tool(file_path, sheet_name, chart_type, data_range, title=None, position=None, style=None, theme=None, custom_palette=None):
        """Insert a professional native chart into a worksheet.

        This function creates a chart based on worksheet data with multiple customization
        options to build professional visualizations directly in Excel.

        Args:
            file_path (str): Full path and name of the Excel file.
            sheet_name (str): Name of the sheet where the chart will be inserted.
            chart_type (str): Type of chart to create. Options include ``'line'``, ``'bar'``, ``'column'``, ``'pie'``, ``'scatter'``,
                             ``'area'``, ``'doughnut'``, ``'radar'``, ``'surface'``, ``'stock'``.
            data_range (str): Range with the data for the chart in Excel format (e.g. ``"A1:D10"``).
            title (str, optional): Title for the chart. Defaults to ``None``.
            position (str, optional): Position to insert the chart in ``A1:F15`` format. Defaults to ``None`` for an automatic position.
            style (int, optional): Numeric chart style (1-48). Defaults to ``None``.
            theme (str, optional): Color theme for the chart. Defaults to ``None``.
            custom_palette (list, optional): List of custom colors in hex (``#RRGGBB``). Defaults to ``None``.

        Returns:
            dict: Information about the operation including details of the created chart.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            SheetNotFoundError: If the specified sheet does not exist.
            RangeError: If the data range is not valid.
            ChartError: If there is a problem creating the chart.

        Example:
            add_chart_tool(
                "C:/data/sales.xlsx",
                "Data",
                "column",
                "A1:B10",
                title="Quarterly Sales",
                position="E1:J15",
                style=12,
                custom_palette=["#4472C4", "#ED7D31", "#A5A5A5"]
            )
        """
        try:
            # Open the file
            wb = openpyxl.load_workbook(file_path)
            
            # Create chart
            chart_id, chart = add_chart(wb, sheet_name, chart_type, data_range, title, position, style, theme, custom_palette)
            
            # Save changes
            wb.save(file_path)
            
            # Extract chart type for a better response message
            chart_type_display = chart_type
            if chart_type.lower() == "col":
                chart_type_display = "column"
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "chart_id": chart_id,
                "chart_type": chart_type_display,
                "data_range": data_range,
                "title": title,
                "position": position,
                "message": f"Chart '{chart_type_display}' successfully created with ID {chart_id}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating chart: {e}"
            }
    
    # Register new combined functions
    @mcp.tool(description="Create a sheet with data in one step")
    def create_sheet_with_data_tool(file_path, sheet_name, data, overwrite=False):
        """Create an Excel file with a single sheet and data in one step.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            file_path (str): Path to the Excel file to create.
            sheet_name (str): Name of the sheet to create.
            data (list): Two-dimensional array with the data.
            overwrite (bool): If ``True`` overwrite the file if it already exists.

        Returns:
            dict: Result of the operation.
        """
        try:
            # Check if the file exists
            file_exists = os.path.exists(file_path)
            
            if file_exists and not overwrite:
                raise FileExistsError(f"The file '{file_path}' already exists. Use overwrite=True to overwrite.")
            
            # Create or open the file
            if not file_exists or overwrite:
                wb = openpyxl.Workbook()
                # Remove the default sheet if it exists
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]
            else:
                wb = openpyxl.load_workbook(file_path)
            
            # Check if the sheet already exists
            if sheet_name in wb.sheetnames:
                if overwrite:
                    # Delete the existing sheet
                    del wb[sheet_name]
                else:
                    raise SheetExistsError(f"The sheet '{sheet_name}' already exists. Use overwrite=True to overwrite.")
            
            # Create the sheet
            ws = wb.create_sheet(sheet_name)
            
            # Write the data
            if data:
                write_sheet_data(ws, "A1", data)
            
            # Save the file
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "rows_written": len(data) if data else 0,
                "columns_written": max([len(row) if isinstance(row, list) else 1 for row in data], default=0) if data else 0,
                "message": f"File created with sheet '{sheet_name}' and data"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating sheet with data: {e}"
            }
    
    @mcp.tool(description="Create a formatted table with data in one step")
    def create_formatted_table_tool(file_path, sheet_name, start_cell, data, table_name, table_style="TableStyleMedium9", formats=None):
        """Create a formatted table with data in one step.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            file_path (str): Path to the Excel file.
            sheet_name (str): Name of the sheet.
            start_cell (str): Starting cell (e.g. ``"A1"``).
            data (list): Two-dimensional array with the data.
            table_name (str): Name for the table.
            table_style (str): Table style.
            formats (dict): Dictionary with formats to apply:
                {
                    "A2:A10": "#,##0.00",  # Numeric format
                    "B2:B10": {"bold": True, "fill_color": "FFFF00"}  # Style
                }

        Returns:
            dict: Result of the operation.
        """
        try:
            # Check if the file exists, if not, create it
            if not os.path.exists(file_path):
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames and sheet_name != "Sheet":
                    # Rename the default sheet
                    wb["Sheet"].title = sheet_name
            else:
                wb = openpyxl.load_workbook(file_path)
                
                # Create the sheet if it doesn't exist
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)
            
            # Get the sheet
            ws = wb[sheet_name]
            
            # Write the data
            write_sheet_data(ws, start_cell, data)
            
            # Determine the table range
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row = start_row + len(data) - 1
            end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
            table_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
            
            # Create the table
            add_table(ws, table_name, table_range, table_style)
            
            # Apply formats if provided
            if formats:
                for cell_range, fmt in formats.items():
                    if isinstance(fmt, dict):
                        # It's a style
                        apply_style(ws, cell_range, fmt)
                    else:
                        # It's a number format
                        apply_number_format(ws, cell_range, fmt)
            
            # Save the file
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "table_name": table_name,
                "table_range": table_range,
                "table_style": table_style,
                "message": f"Table '{table_name}' created and formatted successfully"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating formatted table: {e}"
            }
    
    @mcp.tool(description="Create a chart from new data in one step")
    def create_chart_from_data_tool(file_path, sheet_name, data, chart_type, position=None, title=None, style=None):
        """Create a chart from new data in one step.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            file_path (str): Path to the Excel file.
            sheet_name (str): Name of the sheet.
            data (list): Two-dimensional array with the data for the chart.
            chart_type (str): Chart type (``'column'``, ``'bar'``, ``'line'``, ``'pie'``, etc.).
            position (str): Cell where the chart will be placed (e.g. ``"E1"``).
            title (str): Chart title.
            style: Chart style.

        Returns:
            dict: Result of the operation.
        """
        try:
            # Check if the file exists, if not, create it
            if not os.path.exists(file_path):
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames and sheet_name != "Sheet":
                    # Rename the default sheet
                    wb["Sheet"].title = sheet_name
            else:
                wb = openpyxl.load_workbook(file_path)
                
                # Create the sheet if it doesn't exist
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)
            
            # Get the sheet
            ws = wb[sheet_name]
            
            # Find a free area for the data
            # Search from the left to place the source data
            # (The common convention is to put data on the left and charts on the right)
            start_cell = "A1"
            
            # Check if there is already data in that area
            if ws["A1"].value is not None:
                # Find the first empty column
                col = 1
                while ws.cell(row=1, column=col).value is not None:
                    col += 1
                start_cell = f"{get_column_letter(col)}1"
            
            # Write the data
            write_sheet_data(ws, start_cell, data)
            
            # Determine the data range for the chart
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row = start_row + len(data) - 1
            end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
            data_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
            
            # Determine chart position if not provided
            if not position:
                # Place the chart to the right of the data with a space
                chart_col = end_col + 2  # Dejar una columna de espacio
                position = f"{get_column_letter(chart_col + 1)}1"
            
            # Create the chart
            chart_id, _ = add_chart(wb, sheet_name, chart_type, data_range, title, position, style)
            
            # Save the file
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "data_range": data_range,
                "chart_id": chart_id,
                "chart_type": chart_type,
                "position": position,
                "message": f"Chart '{chart_type}' successfully created from new data"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error creating chart with data: {e}"
            }
    
    
    @mcp.tool(description="Update an existing report with new data")
    def update_report_tool(file_path, data_updates, config_updates=None, recalculate=True):
        """Update an existing report with new data and configuration changes.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            file_path (str): Path to the Excel file to update.
            data_updates (dict): Dictionary with data updates:
                {
                    "sheet_name": {
                        "range1": data_list1,
                        "range2": data_list2,
                        ...
                    }
                }
            config_updates (dict, optional): Configuration updates:
                {
                    "charts": [
                        {
                            "sheet": "sheet_name",
                            "id": 0,  # or "title"
                            "title": "New Title",
                            "style": "new_style"
                        }
                    ],
                    "tables": [
                        {
                            "sheet": "sheet_name",
                            "name": "TableName",
                            "range": "A1:D20"  # New range
                        }
                    ]
                }
            recalculate (bool): If ``True`` recalculate all formulas.

        Returns:
            dict: Result of the operation.
        """
        try:
            # Verify that the file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file does not exist: {file_path}")
            
            # Open the file
            wb = openpyxl.load_workbook(file_path)
            
            # Update data
            for sheet_name, ranges in data_updates.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"Sheet '{sheet_name}' does not exist, it will be skipped")
                    continue
                
                ws = wb[sheet_name]
                
                for range_str, data in ranges.items():
                    # If the range is a single cell, extract the start cell
                    if ':' not in range_str:
                        start_cell = range_str
                    else:
                        start_cell = range_str.split(':')[0]
                    
                    # Write the data
                    write_sheet_data(ws, start_cell, data)
            
            # Update settings
            if config_updates:
                # Update tables
                for table_config in config_updates.get("tables", []):
                    sheet_name = table_config["sheet"]
                    table_name = table_config["name"]
                    
                    if sheet_name not in wb.sheetnames:
                        logger.warning(f"Sheet '{sheet_name}' does not exist to update the table '{table_name}'")
                        continue
                    
                    ws = wb[sheet_name]
                    
                    # Check if the table exists
                    if not hasattr(ws, 'tables') or table_name not in ws.tables:
                        logger.warning(f"The table '{table_name}' does not exist on sheet '{sheet_name}'")
                        continue
                    
                    # Update the table range if provided
                    if "range" in table_config:
                        refresh_table(ws, table_name, table_config["range"])
                
                # Update charts
                for chart_config in config_updates.get("charts", []):
                    sheet_name = chart_config["sheet"]
                    chart_id = chart_config["id"]
                    
                    if sheet_name not in wb.sheetnames:
                        logger.warning(f"Sheet '{sheet_name}' does not exist to update the chart")
                        continue
                    
                    ws = wb[sheet_name]
                    
                    # Check if chart_id is an index or a title
                    if isinstance(chart_id, (int, str)) and str(chart_id).isdigit():
                        chart_idx = int(chart_id)
                    else:
                        # Search the chart by title
                        chart_idx = None
                        for i, chart_rel in enumerate(ws._charts):
                            chart = chart_rel[0]
                            if hasattr(chart, 'title') and chart.title == chart_id:
                                chart_idx = i
                                break
                    
                    if chart_idx is None or chart_idx >= len(ws._charts):
                        logger.warning(f"Chart with ID/title '{chart_id}' not found on sheet '{sheet_name}'")
                        continue
                    
                    # Update chart properties
                    chart = ws._charts[chart_idx][0]
                    
                    if "title" in chart_config:
                        chart.title = chart_config["title"]
                    
                    if "style" in chart_config:
                        try:
                            apply_chart_style(chart, chart_config["style"])
                        except Exception as style_error:
                            logger.warning(f"Error applying style to chart: {style_error}")
            
            # Recalculate formulas if requested
            if recalculate:
                # openpyxl has no direct method to recalculate
                # In Excel, this happens automatically when opening the file
                # Here we simply log that recalculation was requested
                logger.info("Formula recalculation was requested (this will happen when the file is opened in Excel)")
            
            # Save the file
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheets_updated": list(data_updates.keys()),
                "message": f"Report successfully updated: {file_path}"
            }
        except Exception as e:
            logger.error(f"Error updating report: {e}")
            return {
                "success": False,
                "error": str(e),
                "message": f"Error updating report: {e}"
            }
    
    @mcp.tool(description="Create a dynamic dashboard with multiple visualizations in one step")
    def create_dashboard_tool(file_path, data, dashboard_config, overwrite=False):
        """Create a dynamic dashboard with multiple visualizations in one step.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            file_path (str): Path to the Excel file to create.
            data (dict): Dictionary with data per sheet (see docs for format).
            dashboard_config (dict): Dashboard configuration (see docs for format).
            overwrite (bool): If ``True`` overwrite the file if it exists.

        Returns:
            dict: Result of the operation.
        """
        return create_dynamic_dashboard(file_path, data, dashboard_config, overwrite)
    
    @mcp.tool(description="Create a report based on an Excel template, replacing data and updating charts")
    def create_report_from_template_tool(template_file, output_file, data_mappings, chart_mappings=None, format_mappings=None):
        """Create a report from an Excel template, replacing data and updating charts.

        Args:
             **Emojis must never be included in text written to cells, labels, titles or charts.**

            template_file (str): Path to the Excel template.
            output_file (str): Path where the generated report will be saved.
            data_mappings (dict): Data mappings dictionary (see docs for format).
            chart_mappings (dict, optional): Dictionary with chart updates.
            format_mappings (dict, optional): Dictionary with formats to apply.

        Returns:
            dict: Result of the operation.
        """
        return create_report_from_template(template_file, output_file, data_mappings, chart_mappings, format_mappings)
    
    @mcp.tool(description="Import data from multiple sources (CSV, JSON, SQL) into an Excel file")
    def import_data_tool(excel_file, import_config, sheet_name=None, start_cell="A1", create_tables=False):
        """Import data from multiple sources (CSV, JSON, SQL) into an Excel file.

        Args:
            excel_file (str): Path to the Excel file where the data will be imported.
            import_config (dict): Import configuration (see documentation for format).
            sheet_name (str, optional): Default sheet name.
            start_cell (str, optional): Default starting cell.
            create_tables (bool, optional): If ``True`` create Excel tables for each dataset.

        Returns:
            dict: Result of the operation.
        """
        return import_multi_source_data(excel_file, import_config, sheet_name, start_cell, create_tables)
    
    @mcp.tool(description="Export Excel data to multiple formats (CSV, JSON, PDF)")
    def export_data_tool(excel_file, export_config):
        """Export Excel data to multiple formats (CSV, JSON, PDF).

        Args:
            excel_file (str): Path to the source Excel file.
            export_config (dict): Export configuration (see documentation for format).

        Returns:
            dict: Result of the operation.
        """
        return export_excel_data(excel_file, export_config)
    
    @mcp.tool(description="Filter and extract data from a table or range as records")
    def filter_data_tool(file_path, sheet_name, range_str=None, table_name=None, filters=None):
        """Filter and extract data from a table or range as records.

        Args:
            file_path (str): Path to the Excel file.
            sheet_name (str): Name of the sheet.
            range_str (str, optional): Range in ``A1:B5`` format (required if ``table_name`` is not provided).
            table_name (str, optional): Table name (required if ``range_str`` is not provided).
            filters (dict, optional): Filters to apply to the data:
                {
                    "field1": value1,               # Simple equality
                    "field2": [value1, value2],      # List of allowed values
                    "field3": {"gt": 100},           # Greater than
                    "field4": {"lt": 50},            # Less than
                    "field5": {"contains": "text"}   # Contains text
                }

        Returns:
            dict: Result of the operation with the filtered data.
        """
        try:
            # Validate arguments
            if not range_str and not table_name:
                raise ValueError("You must provide 'range_str' or 'table_name'")

            # Open the file
            wb = openpyxl.load_workbook(file_path, data_only=True)

            # Verify that the sheet exists
            if sheet_name not in wb.sheetnames:
                raise SheetNotFoundError(f"Sheet '{sheet_name}' does not exist in the file")
            
            ws = wb[sheet_name]
            
            # If table_name is provided, get its range
            if table_name:
                if not hasattr(ws, 'tables') or table_name not in ws.tables:
                    raise TableNotFoundError(f"Table '{table_name}' does not exist on sheet '{sheet_name}'")
                
                range_str = ws.tables[table_name].ref
            
            # Filter the data
            filtered_data = filter_sheet_data(wb, sheet_name, range_str, filters)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "source": f"Table '{table_name}'" if table_name else f"Range {range_str}",
                "filtered_data": filtered_data,
                "record_count": len(filtered_data),
                "message": f"Found {len(filtered_data)} records that meet the criteria"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error filtering data: {e}"
            }

    @mcp.tool(description="Export a workbook to PDF only if it has a single visible sheet")
    def export_single_sheet_pdf_tool(excel_file, output_pdf=None):
        """Export an Excel file to PDF only if it has a single visible sheet."""
        return export_single_visible_sheet_pdf(excel_file, output_pdf)

    @mcp.tool(description="Export one or more sheets to PDF")
    def export_sheets_pdf_tool(excel_file, sheets=None, output_dir=None, single_file=False):
        """Export the specified sheets of an Excel workbook to PDF.

        ``sheets`` may be a sheet name or a list of names. If ``None`` every existing
        sheet is exported individually. If ``single_file`` is ``True`` and several
        sheets are specified, the function attempts to create a single PDF with all of them.
        """
        return export_sheets_to_pdf(excel_file, sheets, output_dir, single_file)

if __name__ == "__main__":
    logger.info("Master Excel MCP - Usage example")
    logger.info("This module brings together all Excel functionalities in one place.")
