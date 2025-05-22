#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel MCP Master (Model Context Protocol for Excel)
-------------------------------------------------------
Biblioteca unificada para manipular archivos Excel con funcionalidades avanzadas:
- Combina todos los módulos Excel MCP en una interfaz unificada
- Proporciona funciones de alto nivel para operaciones comunes
- Optimiza el flujo de trabajo con Excel

Este módulo integra:
- excel_mcp_complete.py: Lectura y exploración de datos
- workbook_manager_mcp.py: Gestión de libros y hojas
- excel_writer_mcp.py: Escritura y formato de celdas
- advanced_excel_mcp.py: Tablas, fórmulas, gráficos y tablas dinámicas

Author: MCP Team
Version: 1.0

Guía de uso para LLM y agentes
------------------------------
Todas las funciones de esta biblioteca están pensadas para ser utilizadas por
modelos de lenguaje o herramientas automáticas que generan archivos Excel.
Para obtener los mejores resultados se deben seguir estas recomendaciones de
contexto en cada operación:

- **Aplicar estilos en todo momento** para que las hojas resultantes sean
  visualmente agradables. Utiliza las funciones de esta librería para asignar
  estilos a celdas, tablas y gráficos.
- **Evitar la superposición de elementos**. Coloca los gráficos en celdas libres
  y deja al menos un par de filas de separación respecto a tablas o bloques de
  texto. Nunca sitúes gráficos encima de texto.
- **Ajustar automáticamente el ancho de columnas**. Tras escribir tablas o
  conjuntos de datos revisa qué celdas contienen textos largos y aumenta la
  anchura de la columna para que todo sea legible sin romper el diseño.
- **Buscar siempre la disposición más clara y ordenada**, separando secciones y
  agrupando los elementos relacionados para que el fichero final sea fácil de
  entender.
- **Revisar la orientación de los datos**. Si las tablas no son obvias, indica
  explícitamente si las categorías están en filas o columnas para que las
  funciones de gráficos las interpreten correctamente.
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

# Configuración de logging
logger = logging.getLogger("excel_mcp_master")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stderr)
handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
logger.addHandler(handler)

# Importar MCP
try:
    from mcp.server.fastmcp import FastMCP
    HAS_MCP = True
except ImportError:
    logger.warning("No se pudo importar FastMCP. Las funcionalidades de servidor MCP no estarán disponibles.")
    HAS_MCP = False

# Intentar importar las bibliotecas necesarias
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
    logger.warning(f"Error al importar bibliotecas esenciales: {e}")
    logger.warning("Es posible que algunas funcionalidades no estén disponibles")
    HAS_OPENPYXL = False

# Importar módulos Excel MCP existentes
# Nota: En una implementación real, importaríamos las funciones de los módulos existentes
# Sin embargo, para este caso, vamos a reimplementar las funciones clave directamente

# Clases base de excepciones (unificadas)
class ExcelMCPError(Exception):
    """Excepción base para todos los errores de Excel MCP."""
    pass

class FileNotFoundError(ExcelMCPError):
    """Se lanza cuando no se encuentra un archivo Excel."""
    pass

class FileExistsError(ExcelMCPError):
    """Se lanza cuando se intenta crear un archivo que ya existe."""
    pass

class SheetNotFoundError(ExcelMCPError):
    """Se lanza cuando no se encuentra una hoja en el archivo Excel."""
    pass

class SheetExistsError(ExcelMCPError):
    """Se lanza cuando se intenta crear una hoja que ya existe."""
    pass

class CellReferenceError(ExcelMCPError):
    """Se lanza cuando hay un problema con una referencia de celda."""
    pass

class RangeError(ExcelMCPError):
    """Se lanza cuando hay un problema con un rango de celdas."""
    pass

class TableError(ExcelMCPError):
    """Se lanza cuando hay un problema con una tabla de Excel."""
    pass

class ChartError(ExcelMCPError):
    """Se lanza cuando hay un problema con un gráfico."""
    pass

class FormulaError(ExcelMCPError):
    """Se lanza cuando hay un problema con una fórmula."""
    pass

class PivotTableError(ExcelMCPError):
    """Se lanza cuando hay un problema con una tabla dinámica."""
    pass

# Utilidades comunes 
class ExcelRange:
    """
    Clase para manipular y convertir rangos de Excel.
    
    Esta clase proporciona métodos para convertir entre notación de Excel (A1:B5)
    y coordenadas de Python (0-based), además de validar rangos.
    """
    
    @staticmethod
    def parse_cell_ref(cell_ref: str) -> Tuple[int, int]:
        """
        Convierte una referencia de celda en estilo A1 a coordenadas (fila, columna) 0-based.
        
        Args:
            cell_ref: Referencia de celda en formato Excel (ej: 'A1', 'B5')
            
        Returns:
            Tupla (fila, columna) con índices base 0
            
        Raises:
            ValueError: Si la referencia de celda no es válida
        """
        if not cell_ref or not isinstance(cell_ref, str):
            raise ValueError(f"Referencia de celda inválida: {cell_ref}")
        
        # Extraer la parte de columna (letras)
        col_str = ''.join(c for c in cell_ref if c.isalpha())
        # Extraer la parte de fila (números)
        row_str = ''.join(c for c in cell_ref if c.isdigit())
        
        if not col_str or not row_str:
            raise ValueError(f"Formato de celda inválido: {cell_ref}")
        
        # Convertir columna a índice (A->0, B->1, etc.)
        col_idx = 0
        for c in col_str.upper():
            col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
        col_idx -= 1  # Ajustar a base 0
        
        # Convertir fila a índice (base 0)
        row_idx = int(row_str) - 1
        
        return row_idx, col_idx
    
    @staticmethod
    def parse_range(range_str: str) -> Tuple[int, int, int, int]:
        """
        Convierte un rango en estilo A1:B5 a coordenadas (row1, col1, row2, col2) 0-based.
        
        Args:
            range_str: Rango en formato Excel (ej: 'A1:B5')
            
        Returns:
            Tupla (fila_inicio, col_inicio, fila_fin, col_fin) con índices base 0
            
        Raises:
            ValueError: Si el rango no es válido
        """
        if not range_str or not isinstance(range_str, str):
            raise ValueError(f"Rango inválido: {range_str}")
        
        # Manejar rangos con referencia a hoja
        if '!' in range_str:
            parts = range_str.split('!')
            if len(parts) != 2:
                raise ValueError(f"Formato de rango con hoja inválido: {range_str}")
            range_str = parts[1]  # Usar solo la parte del rango
        
        # Dividir el rango en celdas de inicio y fin
        if ':' in range_str:
            start_cell, end_cell = range_str.split(':')
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row, end_col = ExcelRange.parse_cell_ref(end_cell)
        else:
            # Si es una sola celda, inicio y fin son iguales
            start_row, start_col = ExcelRange.parse_cell_ref(range_str)
            end_row, end_col = start_row, start_col
        
        return start_row, start_col, end_row, end_col
    
    @staticmethod
    def cell_to_a1(row: int, col: int) -> str:
        """
        Convierte coordenadas (fila, columna) 0-based a referencia de celda A1.
        
        Args:
            row: Índice de fila (base 0)
            col: Índice de columna (base 0)
            
        Returns:
            Referencia de celda en formato A1
        """
        if row < 0 or col < 0:
            raise ValueError(f"Índices negativos no válidos: fila={row}, columna={col}")
        
        # Convertir columna a letras
        col_str = ""
        col_val = col + 1  # Convertir a base 1 para cálculo
        
        while col_val > 0:
            remainder = (col_val - 1) % 26
            col_str = chr(65 + remainder) + col_str
            col_val = (col_val - 1) // 26
        
        # Convertir fila a número (base 1 para Excel)
        row_val = row + 1
        
        return f"{col_str}{row_val}"
    
    @staticmethod
    def range_to_a1(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
        """
        Convierte coordenadas de rango 0-based a rango A1:B5.
        
        Args:
            start_row: Fila inicial (base 0)
            start_col: Columna inicial (base 0)
            end_row: Fila final (base 0)
            end_col: Columna final (base 0)
            
        Returns:
            Rango en formato A1:B5
        """
        start_cell = ExcelRange.cell_to_a1(start_row, start_col)
        end_cell = ExcelRange.cell_to_a1(end_row, end_col)
        
        if start_cell == end_cell:
            return start_cell
        return f"{start_cell}:{end_cell}"

    @staticmethod
    def parse_range_with_sheet(range_str: str) -> Tuple[Optional[str], int, int, int, int]:
        """Convierte un rango que puede incluir hoja a tupla ``(sheet, row1, col1, row2, col2)``.

        Args:
            range_str: Cadena de rango, posiblemente con prefijo de hoja ``Hoja!A1:B2``.

        Returns:
            Tupla ``(sheet, start_row, start_col, end_row, end_col)`` donde ``sheet``
            es ``None`` si no se especificó hoja.
        """
        if not range_str or not isinstance(range_str, str):
            raise ValueError(f"Rango inválido: {range_str}")

        sheet_name = None
        pure_range = range_str
        if "!" in range_str:
            parts = range_str.split("!", 1)
            if len(parts) != 2:
                raise ValueError(f"Formato de rango con hoja inválido: {range_str}")
            sheet_name = parts[0].strip("'")
            pure_range = parts[1]

        start_row, start_col, end_row, end_col = ExcelRange.parse_range(pure_range)
        return sheet_name, start_row, start_col, end_row, end_col

# Constantes y mapeos
# Mapeo de nombres de estilo a números de estilo de Excel
CHART_STYLE_NAMES = {
    # Estilos claros
    'light-1': 1, 'light-2': 2, 'light-3': 3, 'light-4': 4, 'light-5': 5, 'light-6': 6,
    'office-1': 1, 'office-2': 2, 'office-3': 3, 'office-4': 4, 'office-5': 5, 'office-6': 6,
    'white': 1, 'minimal': 2, 'soft': 3, 'gradient': 4, 'muted': 5, 'outlined': 6,
    
    # Estilos oscuros
    'dark-1': 7, 'dark-2': 8, 'dark-3': 9, 'dark-4': 10, 'dark-5': 11, 'dark-6': 12, 
    'dark-blue': 7, 'dark-gray': 8, 'dark-green': 9, 'dark-red': 10, 'dark-purple': 11, 'dark-orange': 12,
    'navy': 7, 'charcoal': 8, 'forest': 9, 'burgundy': 10, 'indigo': 11, 'rust': 12,
    
    # Estilos coloridos
    'colorful-1': 13, 'colorful-2': 14, 'colorful-3': 15, 'colorful-4': 16, 
    'colorful-5': 17, 'colorful-6': 18, 'colorful-7': 19, 'colorful-8': 20,
    'bright': 13, 'vivid': 14, 'rainbow': 15, 'multi': 16, 'contrast': 17, 'vibrant': 18,
    
    # Temas de Office
    'ion-1': 21, 'ion-2': 22, 'ion-3': 23, 'ion-4': 24,
    'wisp-1': 25, 'wisp-2': 26, 'wisp-3': 27, 'wisp-4': 28,
    'aspect-1': 29, 'aspect-2': 30, 'aspect-3': 31, 'aspect-4': 32,
    'badge-1': 33, 'badge-2': 34, 'badge-3': 35, 'badge-4': 36,
    'gallery-1': 37, 'gallery-2': 38, 'gallery-3': 39, 'gallery-4': 40,
    'median-1': 41, 'median-2': 42, 'median-3': 43, 'median-4': 44,
    
    # Estilos para tipos específicos de gráficos
    'column-default': 1, 'column-dark': 7, 'column-colorful': 13, 
    'bar-default': 1, 'bar-dark': 7, 'bar-colorful': 13,
    'line-default': 1, 'line-dark': 7, 'line-markers': 3, 'line-dash': 5,
    'pie-default': 1, 'pie-dark': 7, 'pie-explosion': 4, 'pie-3d': 10,
    'area-default': 1, 'area-dark': 7, 'area-transparent': 5, 'area-stacked': 9,
    'scatter-default': 1, 'scatter-dark': 7, 'scatter-bubble': 4, 'scatter-smooth': 9,
}

# Mapeo entre estilos y paletas de colores recomendadas
STYLE_TO_PALETTE = {
    # Estilos claros (1-6)
    1: 'office', 2: 'office', 3: 'colorful', 4: 'colorful', 5: 'pastel', 6: 'pastel',
    # Estilos oscuros (7-12)
    7: 'dark-blue', 8: 'dark-gray', 9: 'dark-green', 10: 'dark-red', 11: 'dark-purple', 12: 'dark-orange',
    # Estilos coloridos (13-20)
    13: 'colorful', 14: 'colorful', 15: 'colorful', 16: 'colorful', 
    17: 'colorful', 18: 'colorful', 19: 'colorful', 20: 'colorful',
}

# CHART_COLOR_SCHEMES (con esquemas de colores) - Esta constante estaría normalmente definida 
# en el módulo original, pero la incluyo simplificada
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

# Función auxiliar para obtener una hoja de trabajo (unificada)
def get_sheet(wb, sheet_name_or_index) -> Any:
    """
    Obtiene una hoja de Excel por nombre o índice.
    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name_or_index: Nombre o índice de la hoja
        
    Returns:
        Objeto worksheet
        
    Raises:
        SheetNotFoundError: Si la hoja no existe
    """
    if wb is None:
        raise ExcelMCPError("El workbook no puede ser None")
    
    if isinstance(sheet_name_or_index, int):
        # Si es un índice, intentar acceder por posición
        if 0 <= sheet_name_or_index < len(wb.worksheets):
            return wb.worksheets[sheet_name_or_index]
        else:
            raise SheetNotFoundError(f"No existe una hoja con el índice {sheet_name_or_index}")
    else:
        # Si es un nombre, intentar acceder por nombre
        try:
            return wb[sheet_name_or_index]
        except KeyError:
            sheets_info = ", ".join(wb.sheetnames)
            raise SheetNotFoundError(f"La hoja '{sheet_name_or_index}' no existe en el archivo. Hojas disponibles: {sheets_info}")
        except Exception as e:
            raise ExcelMCPError(f"Error al acceder a la hoja: {e}")

# Función auxiliar para convertir estilo a número
def parse_chart_style(style):
    """
    Convierte diferentes formatos de estilo a un número de estilo de Excel (1-48).
    
    Args:
        style: Estilo en formato int, str numérico, 'styleN', o nombre descriptivo
        
    Returns:
        Número de estilo entre 1-48, o None si no es un estilo válido
    """
    if isinstance(style, int) and 1 <= style <= 48:
        return style
        
    if isinstance(style, str):
        # Caso 1: String numérico '5'
        if style.isdigit():
            style_num = int(style)
            if 1 <= style_num <= 48:
                return style_num
                
        # Caso 2: Formato 'styleN' o 'Style N'
        style_lower = style.lower()
        if style_lower.startswith('style'):
            try:
                # Extraer el número después de 'style'
                num_part = ''.join(c for c in style_lower[5:] if c.isdigit())
                if num_part:
                    style_num = int(num_part)
                    if 1 <= style_num <= 48:
                        return style_num
            except (ValueError, IndexError):
                pass
                
        # Caso 3: Nombre descriptivo ('dark-blue', etc.)
        if style_lower in CHART_STYLE_NAMES:
            return CHART_STYLE_NAMES[style_lower]
            
    return None

# Función auxiliar para aplicar estilos a gráficos y colores a la vez
def apply_chart_style(chart, style):
    """
    Aplica un estilo predefinido al gráfico, incluyendo paleta de colores adecuada.
    
    Args:
        chart: Objeto de gráfico openpyxl
        style: Estilo en cualquier formato soportado (número, nombre, etc.)
    
    Returns:
        True si se aplicó correctamente, False en caso contrario
    """
    # Convertir a número de estilo si no lo es ya
    style_number = parse_chart_style(style)
    
    if style_number is None:
        style_str = str(style) if style else "None"
        logger.warning(f"Estilo de gráfico inválido: '{style_str}'. Debe ser un número entre 1-48 o un nombre de estilo válido.")
        logger.info("Nombres de estilo válidos incluyen: 'dark-blue', 'light-1', 'colorful-3', etc.")
        return False
        
    if not (1 <= style_number <= 48):
        logger.warning(f"Estilo de gráfico inválido: {style_number}. Debe estar entre 1 y 48.")
        return False
    
    # Paso 1: Aplicar número de estilo a atributos del estilo nativo de Excel
    try:
        # La propiedad style en openpyxl se corresponde con el número de estilo de Excel
        chart.style = style_number
        logger.info(f"Aplicado estilo nativo {style_number} al gráfico")
    except Exception as e:
        logger.warning(f"Error al aplicar estilo {style_number}: {e}")
    
    # Paso 2: Aplicar paleta de colores asociada según el tema correspondiente al estilo
    palette_name = STYLE_TO_PALETTE.get(style_number, 'default')
    colors = CHART_COLOR_SCHEMES.get(palette_name, CHART_COLOR_SCHEMES['default'])
    
    # Aplicar colores a las series
    try:
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.fill import ColorChoice
        
        for i, series in enumerate(chart.series):
            if i < len(colors):
                # Asegurarse de que existen propiedades gráficas
                if not hasattr(series, 'graphicalProperties') or series.graphicalProperties is None:
                    series.graphicalProperties = GraphicalProperties()
                    
                # Asignar color usando ColorChoice para mejor compatibilidad
                color = colors[i % len(colors)]
                if isinstance(color, str) and color.startswith('#'):
                    color = color[1:]
                    
                series.graphicalProperties.solidFill = ColorChoice(srgbClr=color)
                
        logger.info(f"Aplicado estilo {style_number} con paleta '{palette_name}' al gráfico")
        return True
        
    except Exception as e:
        logger.warning(f"Error al aplicar colores para estilo {style_number}: {e}")
        return False

def determine_orientation(ws: Any, min_row: int, min_col: int, max_row: int, max_col: int) -> bool:
    """Intenta deducir la orientación de los datos.

    Devuelve ``True`` si las categorías parecen estar en la primera columna
    (orientación por columnas) y ``False`` si lo más probable es que se
    encuentren en la primera fila. El algoritmo compara la proporción de valores
    numéricos al interpretar los datos de ambas maneras y usa la forma del rango
    como desempate. Está pensado para que un LLM evite elegir encabezados
    equivocados en tablas cuadradas o poco claras.
    """

    def _is_number(value: Any) -> bool:
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    # Calcular ratio de números asumiendo categorías en la primera columna
    col_numeric = col_total = 0
    for c in range(min_col + 1, max_col + 1):
        for r in range(min_row, max_row + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                col_total += 1
                if _is_number(val):
                    col_numeric += 1

    col_ratio = (col_numeric / col_total) if col_total else 0

    # Calcular ratio de números asumiendo categorías en la primera fila
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
        return False  # encabezados en la primera fila
    if col_ratio > row_ratio:
        return True   # encabezados en la primera columna

    # Desempate por forma del rango
    return (max_row - min_row) >= (max_col - min_col)

def _trim_range_to_data(ws: Any, min_row: int, min_col: int, max_row: int, max_col: int) -> Tuple[int, int, int, int]:
    """Elimina filas y columnas vacías al final de un rango."""
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
# FUNCIONES BASE (reimplementadas de los módulos originales)
# ----------------------------------------

# 1. Gestión de Workbooks (de workbook_manager_mcp.py)
def create_workbook(filename: str, overwrite: bool = False) -> Any:
    """
    Crea un nuevo fichero Excel vacío.
    
    Args:
        filename (str): Ruta y nombre del archivo a crear.
        overwrite (bool, opcional): Si es True, sobreescribe archivo existente.
        
    Returns:
        Objeto Workbook.
        
    Raises:
        FileExistsError: Si el archivo existe y overwrite es False.
    """
    if os.path.exists(filename) and not overwrite:
        raise FileExistsError(f"El archivo '{filename}' ya existe. Use overwrite=True para sobreescribir.")
    
    wb = openpyxl.Workbook()
    if overwrite:
        save_workbook(wb, filename)
    return wb

def open_workbook(filename: str) -> Any:
    """
    Abre un fichero Excel existente.
    
    Args:
        filename (str): Ruta del archivo.
        
    Returns:
        Objeto Workbook.
        
    Raises:
        FileNotFoundError: Si el archivo no existe.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"El archivo '{filename}' no existe.")
    
    try:
        wb = openpyxl.load_workbook(filename)
        return wb
    except Exception as e:
        logger.error(f"Error al abrir el archivo '{filename}': {e}")
        raise ExcelMCPError(f"Error al abrir el archivo: {e}")

def save_workbook(wb: Any, filename: Optional[str] = None) -> str:
    """
    Guarda el Workbook en disco.
    
    Args:
        wb: Objeto Workbook.
        filename (str, opcional): Si se indica, guarda con otro nombre.
        
    Returns:
        Ruta del fichero guardado.
        
    Raises:
        ExcelMCPError: Si hay error al guardar.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    try:
        # Si no se proporciona filename, usar el filename original si existe
        if not filename and hasattr(wb, 'path'):
            filename = wb.path
        elif not filename:
            raise ExcelMCPError("Debe proporcionar un nombre de archivo")
        
        wb.save(filename)
        return filename
    except Exception as e:
        logger.error(f"Error al guardar el workbook en '{filename}': {e}")
        raise ExcelMCPError(f"Error al guardar el workbook: {e}")

def close_workbook(wb: Any) -> None:
    """
    Cierra el Workbook en memoria.
    
    Args:
        wb: Objeto Workbook.
        
    Returns:
        Ninguno.
    """
    if not wb:
        return
    
    try:
        # Openpyxl realmente no tiene un método close(),
        # pero podemos eliminar referencias para ayudar al GC
        if hasattr(wb, "_archive"):
            wb._archive.close()
    except Exception as e:
        logger.warning(f"Advertencia al cerrar workbook: {e}")

def list_sheets(wb: Any) -> List[str]:
    """
    Devuelve lista de nombres de hojas.
    
    Args:
        wb: Objeto Workbook.
        
    Returns:
        List[str]: Lista de nombres de hojas.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    if hasattr(wb, 'sheetnames'):
        return wb.sheetnames
    
    # Alternativa si no se puede acceder a sheetnames
    sheet_names = []
    for sheet in wb.worksheets:
        if hasattr(sheet, 'title'):
            sheet_names.append(sheet.title)
    
    return sheet_names

def add_sheet(wb: Any, sheet_name: str, index: Optional[int] = None) -> Any:
    """
    Añade una nueva hoja vacía.
    
    Args:
        wb: Objeto Workbook.
        sheet_name (str): Nombre de la hoja.
        index (int, opcional): Posición en pestañas.
        
    Returns:
        Hoja creada.
        
    Raises:
        SheetExistsError: Si ya existe una hoja con ese nombre.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    # Verificar si ya existe una hoja con ese nombre
    if sheet_name in list_sheets(wb):
        raise SheetExistsError(f"Ya existe una hoja con el nombre '{sheet_name}'")
    
    # Crear nueva hoja
    if index is not None:
        ws = wb.create_sheet(sheet_name, index)
    else:
        ws = wb.create_sheet(sheet_name)
    
    return ws

def delete_sheet(wb: Any, sheet_name: str) -> None:
    """
    Elimina la hoja indicada.
    
    Args:
        wb: Objeto Workbook.
        sheet_name (str): Nombre de la hoja a eliminar.
        
    Raises:
        SheetNotFoundError: Si la hoja no existe.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    # Verificar que la hoja existe
    if sheet_name not in list_sheets(wb):
        raise SheetNotFoundError(f"La hoja '{sheet_name}' no existe en el workbook")
    
    # Eliminar la hoja
    try:
        del wb[sheet_name]
    except Exception as e:
        logger.error(f"Error al eliminar la hoja '{sheet_name}': {e}")
        raise ExcelMCPError(f"Error al eliminar la hoja: {e}")

def rename_sheet(wb: Any, old_name: str, new_name: str) -> None:
    """
    Renombra una hoja.
    
    Args:
        wb: Objeto Workbook.
        old_name (str): Nombre actual de la hoja.
        new_name (str): Nuevo nombre para la hoja.
        
    Raises:
        SheetNotFoundError: Si la hoja original no existe.
        SheetExistsError: Si ya existe una hoja con el nuevo nombre.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    # Verificar que la hoja original existe
    if old_name not in list_sheets(wb):
        raise SheetNotFoundError(f"La hoja '{old_name}' no existe en el workbook")
    
    # Verificar que no exista una hoja con el nuevo nombre
    if new_name in list_sheets(wb) and old_name != new_name:
        raise SheetExistsError(f"Ya existe una hoja con el nombre '{new_name}'")
    
    # Renombrar la hoja
    try:
        wb[old_name].title = new_name
    except Exception as e:
        logger.error(f"Error al renombrar la hoja '{old_name}' a '{new_name}': {e}")
        raise ExcelMCPError(f"Error al renombrar la hoja: {e}")

# 2. Lectura y exploración de datos (de excel_mcp_complete.py)
def read_sheet_data(wb: Any, sheet_name: str, range_str: Optional[str] = None, 
                   formulas: bool = False) -> List[List[Any]]:
    """
    Lee valores y, opcionalmente, fórmulas de una hoja de Excel.
    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name: Nombre de la hoja
        range_str: Rango en formato A1:B5, o None para toda la hoja
        formulas: Si es True, devuelve fórmulas en lugar de valores calculados
    
    Returns:
        Lista de listas con los valores o fórmulas de las celdas
        
    Raises:
        SheetNotFoundError: Si la hoja no existe
        RangeError: Si el rango es inválido
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Si no se especifica rango, usar toda la hoja
    if not range_str:
        # Determinar el rango usado (min_row, min_col, max_row, max_col)
        min_row, min_col = 1, 1
        max_row = ws.max_row
        max_col = ws.max_column
    else:
        # Parsear el rango especificado
        try:
            min_row, min_col, max_row, max_col = ExcelRange.parse_range(range_str)
            # Convertir a base 1 para openpyxl
            min_row += 1
            min_col += 1
            max_row += 1
            max_col += 1
        except ValueError as e:
            raise RangeError(f"Rango inválido '{range_str}': {e}")
    
    # Extraer datos del rango
    data = []
    for row in range(min_row, max_row + 1):
        row_data = []
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            
            # Obtener el valor adecuado (fórmula o valor calculado)
            if formulas and cell.data_type == 'f':
                # Si queremos fórmulas y la celda tiene una fórmula
                value = cell.value  # Esto es la fórmula con '='
            else:
                # Valor normal o calculado
                value = cell.value
            
            row_data.append(value)
        data.append(row_data)
    
    return data

def list_tables(wb: Any, sheet_name: str) -> List[Dict[str, Any]]:
    """
    Lista todas las tablas definidas en una hoja de Excel.
    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name: Nombre de la hoja
        
    Returns:
        Lista de diccionarios con información de las tablas
        
    Raises:
        SheetNotFoundError: Si la hoja no existe
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Lista para almacenar información de las tablas
    tables_info = []
    
    # Verificar si la hoja tiene tablas
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
    Obtiene los datos de una tabla específica en formato de registros.
    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name: Nombre de la hoja
        table_name: Nombre de la tabla
        
    Returns:
        Lista de diccionarios, donde cada diccionario representa una fila
        
    Raises:
        SheetNotFoundError: Si la hoja no existe
        TableError: Si la tabla no existe
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Verificar si la tabla existe
    if not hasattr(ws, 'tables') or table_name not in ws.tables:
        raise TableError(f"La tabla '{table_name}' no existe en la hoja '{sheet_name}'")
    
    # Obtener la referencia de la tabla
    table = ws.tables[table_name]
    table_range = table.ref
    
    # Parsear el rango
    min_row, min_col, max_row, max_col = ExcelRange.parse_range(table_range)
    
    # Ajustar a base 1 para openpyxl
    min_row += 1
    min_col += 1
    max_row += 1
    max_col += 1
    
    # Extraer encabezados (primera fila)
    headers = []
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=min_row, column=col)
        headers.append(cell.value or f"Column{col}")
    
    # Extraer datos (filas después del encabezado)
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
    Lista todos los gráficos en una hoja de Excel.
    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name: Nombre de la hoja
        
    Returns:
        Lista de diccionarios con información de los gráficos
        
    Raises:
        SheetNotFoundError: Si la hoja no existe
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Lista para almacenar información de los gráficos
    charts_info = []
    
    # Verificar si la hoja tiene gráficos
    if hasattr(ws, '_charts'):
        for chart_id, chart_rel in enumerate(ws._charts):
            chart = chart_rel[0]  # El elemento 0 es el objeto chart, el 1 es position
            
            # Determinar el tipo de gráfico
            chart_type = "desconocido"
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
            
            # Recopilar información del gráfico
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
    Escribe un array bidimensional de valores o fórmulas.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    Para garantizar que la salida sea legible cuando la función es utilizada por
    un modelo de lenguaje se recomienda aplicar estilos tras la escritura y
    comprobar la longitud de las celdas resultantes. Si alguna columna contiene
    textos muy largos, se debe aumentar su ancho para evitar que el contenido se
    corte. De esta forma los ficheros generados tendrán un aspecto profesional.

    Args:
        ws: Objeto worksheet de openpyxl
        start_cell (str): Celda de anclaje (e.j. "A1")
        data (List[List]): Valores o cadenas "=FÓRMULA(...)"
        
    Raises:
        CellReferenceError: Si la referencia de celda es inválida
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    if not data or not isinstance(data, list):
        raise ExcelMCPError("Los datos deben ser una lista no vacía")
    
    try:
        # Parsear la celda inicial para obtener fila y columna base
        start_row, start_col = ExcelRange.parse_cell_ref(start_cell)

        # Escribir los datos
        for i, row_data in enumerate(data):
            if row_data is None:
                continue

            if not isinstance(row_data, list):
                # Si no es una lista, tratar como valor único
                row_data = [row_data]

            for j, value in enumerate(row_data):
                # Calcular coordenadas de celda (base 1 para openpyxl)
                row = start_row + i + 1
                col = start_col + j + 1

                # Escribir el valor
                cell = ws.cell(row=row, column=col)
                cell.value = value

        # ----------------------------------------------------
        # Ajuste automático de columnas y filas del rango escrito
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
        raise CellReferenceError(f"Referencia de celda inválida '{start_cell}': {e}")
    except Exception as e:
        raise ExcelMCPError(f"Error al escribir datos: {e}")

def append_rows(ws: Any, data: List[List[Any]]) -> None:
    """
    Añade filas al final con los valores dados.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        ws: Objeto worksheet de openpyxl
        data (List[List]): Valores o cadenas "=FÓRMULA(...)"
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    if not data or not isinstance(data, list):
        raise ExcelMCPError("Los datos deben ser una lista no vacía")
    
    try:
        for row_data in data:
            if not isinstance(row_data, list):
                # Si no es una lista, convertir a lista con un solo elemento
                row_data = [row_data]
            
            ws.append(row_data)
    
    except Exception as e:
        raise ExcelMCPError(f"Error al añadir filas: {e}")

def update_cell(ws: Any, cell: str, value_or_formula: Any) -> None:
    """
    Actualiza individualmente una celda.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        ws: Objeto worksheet de openpyxl
        cell (str): Referencia de celda (e.j. "A1")
        value_or_formula: Valor o fórmula a asignar
        
    Raises:
        CellReferenceError: Si la referencia de celda es inválida
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Asignar valor a la celda
        cell_obj = ws[cell]
        cell_obj.value = value_or_formula

        # ----------------------------------------------
        # Ajuste automático si se escribe texto largo
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
        raise CellReferenceError(f"Referencia de celda inválida: '{cell}'")
    except Exception as e:
        raise ExcelMCPError(f"Error al actualizar celda: {e}")

def autofit_table(ws: Any, cell_range: str) -> None:
    """Ajusta ancho de columnas y alto de filas para un rango tabular."""
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
    Aplica estilos de celda a un rango.
    
    Args:
        ws: Objeto worksheet de openpyxl
        cell_range (str): Rango en formato A1:B5, o una sola celda (e.j. "A1")
        style_dict (dict): Diccionario con estilos a aplicar:
            - font_name (str): Nombre de la fuente
            - font_size (int): Tamaño de la fuente
            - bold (bool): Negrita
            - italic (bool): Cursiva
            - fill_color (str): Color de fondo (formato hex: "FF0000")
            - border_style (str): Estilo de borde ('thin', 'medium', 'thick', etc.)
            - alignment (str): Alineación ('center', 'left', 'right', etc.)
            
    Raises:
        RangeError: Si el rango es inválido
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Parsear el rango
        if ':' in cell_range:
            # Rango de celdas
            start_cell, end_cell = cell_range.split(':')
            start_coord = ws[start_cell].coordinate
            end_coord = ws[end_cell].coordinate
            range_str = f"{start_coord}:{end_coord}"
        else:
            # Una sola celda
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
            
            # Mapear valores de alineación horizontal
            if alignment_value in ['left', 'center', 'right', 'justify']:
                horizontal = alignment_value
            
            alignment = Alignment(horizontal=horizontal)
        
        # Aplicar estilos a todas las celdas del rango
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
        raise RangeError(f"Rango inválido: '{cell_range}'")
    except Exception as e:
        raise ExcelMCPError(f"Error al aplicar estilos: {e}")

def apply_number_format(ws: Any, cell_range: str, fmt: str) -> None:
    """
    Aplica formato numérico a un rango de celdas.
    
    Args:
        ws: Objeto worksheet de openpyxl
        cell_range (str): Rango en formato A1:B5, o una sola celda (e.j. "A1")
        fmt (str): Formato numérico ("#,##0.00", "0%", "dd/mm/yyyy", etc.)
        
    Raises:
        RangeError: Si el rango es inválido
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Parsear el rango
        if ':' in cell_range:
            # Rango de celdas
            start_cell, end_cell = cell_range.split(':')
            start_coord = ws[start_cell].coordinate
            end_coord = ws[end_cell].coordinate
            range_str = f"{start_coord}:{end_coord}"
        else:
            # Una sola celda
            range_str = cell_range
        
        # Aplicar formato a todas las celdas del rango
        for row in ws[range_str]:
            for cell in row:
                cell.number_format = fmt
    
    except KeyError:
        raise RangeError(f"Rango inválido: '{cell_range}'")
    except Exception as e:
        raise ExcelMCPError(f"Error al aplicar formato numérico: {e}")

# 4. Tablas y fórmulas (de advanced_excel_mcp.py)
def add_table(ws: Any, table_name: str, cell_range: str, style=None) -> Any:
    """
    Define un rango como Tabla con estilo.
    
    Args:
        ws: Objeto worksheet de openpyxl
        table_name (str): Nombre único para la tabla
        cell_range (str): Rango en formato A1:B5
        style (str, opcional): Nombre de estilo predefinido o dict personalizado
        
    Returns:
        Objeto Table creado
        
    Raises:
        TableError: Si hay un problema con la tabla (ej. nombre duplicado)
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Verificar si ya existe una tabla con ese nombre
        if hasattr(ws, 'tables') and table_name in ws.tables:
            raise TableError(f"Ya existe una tabla con el nombre '{table_name}'")
        
        # Crear objeto de tabla
        table = Table(displayName=table_name, ref=cell_range)
        
        # Aplicar estilo
        if style:
            if isinstance(style, dict):
                # Estilo personalizado
                style_info = TableStyleInfo(**style)
            else:
                # Estilo predefinido
                style_info = TableStyleInfo(
                    name=style,
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
            table.tableStyleInfo = style_info
        
        # Añadir tabla a la hoja
        ws.add_table(table)

        # ------------------------------
        # Ajuste automático de columnas y filas de la tabla
        # ------------------------------
        try:
            autofit_table(ws, cell_range)
        except Exception:
            pass

        return table
    
    except Exception as e:
        if "Duplicate name" in str(e):
            raise TableError(f"Ya existe una tabla con el nombre '{table_name}'")
        elif "Invalid coordinate" in str(e) or "Invalid cell" in str(e):
            raise RangeError(f"Rango inválido: '{cell_range}'")
        else:
            raise TableError(f"Error al añadir tabla: {e}")

def set_formula(ws: Any, cell: str, formula: str) -> Any:
    """
    Establece una fórmula en una celda.
    
    Args:
        ws: Objeto worksheet de openpyxl
        cell (str): Referencia de celda (ej. "A1")
        formula (str): Fórmula Excel, con o sin signo =
        
    Returns:
        Celda actualizada
        
    Raises:
        FormulaError: Si hay un problema con la fórmula
    """
    if not ws:
        raise ExcelMCPError("El worksheet no puede ser None")
    
    try:
        # Añadir signo = si no está presente
        if formula and not formula.startswith('='):
            formula = f"={formula}"
        
        # Establecer la fórmula
        ws[cell] = formula
        return ws[cell]
    
    except KeyError:
        raise RangeError(f"Celda inválida: '{cell}'")
    except Exception as e:
        raise FormulaError(f"Error al establecer fórmula: {e}")

# 5. Gráficos y tablas dinámicas (de advanced_excel_mcp.py)
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
    """Inserta un gráfico nativo utilizando los datos del rango indicado.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    ``data_range`` debe referirse a una tabla rectangular sin celdas vacías en
    la zona de valores. La primera fila o la primera columna se interpretan como
    encabezados y categorías, según determine_orientation. Todas las series deben
    contener únicamente números y tener la misma longitud que el vector de
    categorías. Si existen filas de totales o columnas mezcladas, el gráfico
    podría crearse de forma incorrecta.

    Valida previamente que el rango pertenezca a la hoja correcta y que los
    encabezados estén presentes, ya que las series se añaden con
    ``titles_from_data=True``. Las categorías no deben contener valores en blanco
    ni duplicados, y las columnas numéricas no pueden incluir texto.

    Args:
        wb: Objeto ``Workbook`` de openpyxl.
        sheet_name: Nombre de la hoja donde insertar el gráfico.
        chart_type: Tipo de gráfico (``'column'``, ``'bar'``, ``'line'``,
            ``'pie'``, etc.).
        data_range: Rango de datos en formato ``A1:B5``.
        title: Título opcional del gráfico.
        position: Celda donde colocar el gráfico (por ejemplo ``"E5"``).
        style: Estilo del gráfico (número ``1``–``48`` o nombre descriptivo).
        theme: Nombre del tema de color.
        custom_palette: Lista de colores personalizados.

    Returns:
        Tupla ``(id del gráfico, objeto chart)``.

    Raises:
        ChartError: Si ocurre un problema al crear el gráfico.
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    try:
        # Obtener la hoja
        ws = get_sheet(wb, sheet_name)
        
        # Validar y normalizar la posición
        if position:
            # Si la posición es un rango, tomar solo la primera celda
            if ':' in position:
                position = position.split(':')[0]
            
            # Validar que es una referencia de celda válida
            try:
                # Intentar parsear para verificar que es una referencia válida
                ExcelRange.parse_cell_ref(position)
            except ValueError:
                raise ValueError(f"Posición inválida '{position}'. Debe ser una referencia de celda (ej: 'E4')")
        
        # Crear objeto de gráfico según el tipo
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
            raise ChartError(f"Tipo de gráfico no soportado: '{chart_type}'")
        
        # Configurar título si se proporciona
        if title:
            chart.title = title
            
        # Determinar si el rango hace referencia a otra hoja
        data_sheet_name, sr, sc, er, ec = ExcelRange.parse_range_with_sheet(data_range)
        if data_sheet_name is None:
            data_sheet_name = sheet_name
        data_ws = get_sheet(wb, data_sheet_name)

        # Normalizar el rango para Reference (con nombre de hoja escapado)
        if " " in data_sheet_name or any(c in data_sheet_name for c in "![]{}?"):
            sheet_prefix = f"'{data_sheet_name}'!"
        else:
            sheet_prefix = f"{data_sheet_name}!"
        clean_range = ExcelRange.range_to_a1(sr, sc, er, ec)
        
        # Parsear rango de datos
        try:
            # Usar los límites calculados previamente
            min_row = sr + 1
            min_col = sc + 1
            max_row = er + 1
            max_col = ec + 1

            # Recortar filas o columnas vacías al final
            min_row, min_col, max_row, max_col = _trim_range_to_data(data_ws, min_row, min_col, max_row, max_col)
            if max_row < min_row or max_col < min_col:
                raise ChartError("El rango indicado no contiene datos")

            # Determinar orientación analizando el contenido del rango
            is_column_oriented = determine_orientation(data_ws, min_row, min_col, max_row, max_col)
            
            # Para gráficos que necesitan categorías (la mayoría excepto scatter)
            if chart_type.lower() != 'scatter':
                if is_column_oriented:
                    if _range_has_blank(data_ws, min_row + 1, min_col + 1, max_row, max_col):
                        raise ChartError("El rango de datos contiene celdas vacías")
                    categories = Reference(data_ws, min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=min_col)
                    data = Reference(data_ws, min_row=min_row, max_row=max_row, min_col=min_col + 1, max_col=max_col)
                    try:
                        chart.add_data(data, titles_from_data=True)
                    except TypeError:
                        chart.add_data(data)
                    chart.set_categories(categories)
                else:
                    if _range_has_blank(data_ws, min_row + 1, min_col, max_row, max_col):
                        raise ChartError("El rango de datos contiene celdas vacías")
                    categories = Reference(data_ws, min_row=min_row, max_row=min_row, min_col=min_col, max_col=max_col)
                    data = Reference(data_ws, min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=max_col)
                    try:
                        chart.add_data(data, titles_from_data=True)
                    except TypeError:
                        chart.add_data(data)
                    chart.set_categories(categories)
            else:
                if _range_has_blank(data_ws, min_row, min_col, max_row, max_col):
                    raise ChartError("El rango de datos contiene celdas vacías")
                data_ref = Reference(data_ws, min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col)
                chart.add_data(data_ref)
        
        except Exception as e:
            raise RangeError(f"Error al procesar rango de datos '{data_range}': {e}")
        
        # Aplicar estilos
        if style is not None:
            # Convertir estilo especificado (número, nombre, etc.)
            style_number = parse_chart_style(style)
            if style_number is not None:
                # Aplicar el estilo incluyendo la paleta de colores
                apply_chart_style(chart, style_number)
            else:
                logger.warning(f"Estilo de gráfico inválido: '{style}'. Se usará estilo predeterminado.")
        
        # Aplicar tema de color si se proporciona
        # (aquí usaríamos el tema, pero por simplicidad lo omitimos en esta implementación)
        
        # Aplicar paleta personalizada si se proporciona
        if custom_palette and isinstance(custom_palette, list):
            from openpyxl.chart.shapes import GraphicalProperties
            from openpyxl.drawing.fill import ColorChoice
            
            for i, series in enumerate(chart.series):
                if i < len(custom_palette):
                    # Asegurarse de que existen propiedades gráficas
                    if not hasattr(series, 'graphicalProperties'):
                        series.graphicalProperties = GraphicalProperties()
                    elif series.graphicalProperties is None:
                        series.graphicalProperties = GraphicalProperties()
                    
                    # Asignar color asegurándonos de que no tiene el prefijo #
                    color = custom_palette[i]
                    if isinstance(color, str) and color.startswith('#'):
                        color = color[1:]
                    
                    # Aplicar el color de forma explícita
                    series.graphicalProperties.solidFill = ColorChoice(srgbClr=color)
        
        # Posicionar el gráfico en la hoja
        if position:
            ws.add_chart(chart, position)
        else:
            ws.add_chart(chart)
        
        # Determinar el ID del gráfico (basado en su posición en la lista)
        chart_id = len(ws._charts) - 1
        
        return chart_id, chart
    
    except SheetNotFoundError:
        raise
    except ChartError:
        raise
    except RangeError:
        raise
    except Exception as e:
        raise ChartError(f"Error al crear gráfico: {e}")

def add_pivot_table(wb: Any, source_sheet: str, source_range: str, target_sheet: str, 
                   target_cell: str, rows: List[str], cols: List[str], data_fields: List[str]) -> Any:
    """
    Crea una tabla dinámica.
    
    Args:
        wb: Objeto workbook de openpyxl
        source_sheet (str): Hoja con datos fuente
        source_range (str): Rango de datos fuente (A1:E10)
        target_sheet (str): Hoja donde crear la tabla dinámica
        target_cell (str): Celda de anclaje (ej. "A1")
        rows (list): Campos para filas
        cols (list): Campos para columnas
        data_fields (list): Campos para valores y funciones
        
    Returns:
        Objeto PivotTable creado
        
    Raises:
        PivotTableError: Si hay problemas con la tabla dinámica
    """
    if not wb:
        raise ExcelMCPError("El workbook no puede ser None")
    
    try:
        # Obtener hoja de origen
        source_ws = get_sheet(wb, source_sheet)
        
        # Obtener hoja de destino
        target_ws = get_sheet(wb, target_sheet)
        
        logger.warning("Las tablas dinámicas en openpyxl tienen funcionalidad limitada y pueden no funcionar como se espera.")
        
        # Intentar crear la caché de datos (este es un paso necesario)
        try:
            # Parsear el rango
            min_row, min_col, max_row, max_col = ExcelRange.parse_range(source_range)
            
            # Ajustar a base 1 para Reference
            min_row += 1
            min_col += 1
            max_row += 1
            max_col += 1
            
            # Crear referencia de datos para la caché
            data_reference = Reference(source_ws, min_row=min_row, min_col=min_col,
                                     max_row=max_row, max_col=max_col)
            
            # Crear caché de pivot
            pivot_cache = PivotCache(cacheSource=data_reference, cacheDefinition={'refreshOnLoad': True})
            
            # Generar un ID único para la tabla dinámica
            pivot_name = f"PivotTable{len(wb._pivots) + 1 if hasattr(wb, '_pivots') else 1}"
            
            # Crear la tabla dinámica
            pivot_table = PivotTable(name=pivot_name, cache=pivot_cache,
                                    location=target_cell, rowGrandTotals=True, colGrandTotals=True)
            
            # Añadir campos de fila
            for row_field in rows:
                pivot_table.rowFields.append(PivotField(data=row_field))
            
            # Añadir campos de columna
            for col_field in cols:
                pivot_table.colFields.append(PivotField(data=col_field))
            
            # Añadir campos de datos
            for data_field in data_fields:
                pivot_table.dataFields.append(PivotField(data=data_field))
            
            # Añadir la tabla dinámica a la hoja de destino
            target_ws.add_pivot_table(pivot_table)
            
            return pivot_table
            
        except Exception as pivot_error:
            logger.error(f"Error al crear tabla dinámica: {pivot_error}")
            raise PivotTableError(f"Error al crear tabla dinámica: {pivot_error}")
    
    except SheetNotFoundError:
        raise
    except PivotTableError:
        raise
    except Exception as e:
        raise PivotTableError(f"Error al crear tabla dinámica: {e}")

# ----------------------------------------
# NUEVAS FUNCIONES COMBINADAS DE ALTO NIVEL
# ----------------------------------------

def create_sheet_with_data(wb: Any, sheet_name: str, data: List[List[Any]], 
                          index: Optional[int] = None, overwrite: bool = False) -> Any:
    """
    Crea una nueva hoja y escribe datos en un solo paso.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name (str): Nombre para la nueva hoja
        data (List[List]): Datos a escribir
        index (int, opcional): Posición de la hoja en el libro
        overwrite (bool): Si True, sobrescribe una hoja existente con el mismo nombre
        
    Returns:
        Objeto worksheet creado
        
    Raises:
        SheetExistsError: Si la hoja ya existe y overwrite=False
    """
    # Manejar caso de hoja existente
    if sheet_name in list_sheets(wb):
        if overwrite:
            # Eliminar la hoja existente
            delete_sheet(wb, sheet_name)
        else:
            raise SheetExistsError(f"La hoja '{sheet_name}' ya existe. Use overwrite=True para sobrescribirla.")
    
    # Crear nueva hoja
    ws = add_sheet(wb, sheet_name, index)
    
    # Escribir datos
    if data:
        write_sheet_data(ws, "A1", data)
    
    return ws

def create_formatted_table(wb: Any, sheet_name: str, start_cell: str, data: List[List[Any]], 
                          table_name: str, table_style: Optional[str] = None, 
                          formats: Optional[Dict[str, Union[str, Dict]]] = None) -> Tuple[Any, Any]:
    """
    Crea una tabla con formato en un solo paso.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        wb: Objeto workbook de openpyxl
        sheet_name (str): Nombre de la hoja donde crear la tabla
        start_cell (str): Celda inicial para los datos (ej. "A1")
        data (List[List]): Datos para la tabla, incluyendo encabezados
        table_name (str): Nombre único para la tabla
        table_style (str, opcional): Estilo predefinido para la tabla (ej. "TableStyleMedium9")
        formats (dict, opcional): Diccionario de formatos a aplicar:
            - Claves: Rangos relativos (ej. "A2:A10") o celdas
            - Valores: Formato de número o diccionario de estilos 
            
    Returns:
        Tupla (objeto tabla, worksheet)
        
    Ejemplo de formato:
        formats = {
            "B2:B10": "#,##0.00",  # Formato de moneda
            "A1:Z1": {"bold": True, "fill_color": "DDEBF7"}  # Estilo de encabezado
        }
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Obtener dimensiones del rango de datos
    rows = len(data)
    cols = max([len(row) if isinstance(row, list) else 1 for row in data], default=0)
    
    # Escribir los datos
    write_sheet_data(ws, start_cell, data)
    
    # Calcular el rango completo de la tabla
    start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1
    full_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
    
    # Crear la tabla
    table = add_table(ws, table_name, full_range, table_style)
    
    # Aplicar formatos adicionales si se proporcionan
    if formats:
        for range_str, format_value in formats.items():
            # Convertir rango relativo a absoluto si es necesario
            if not any(c in range_str for c in [':', '!']):
                # Es una sola celda, añadir offset
                cell_row, cell_col = ExcelRange.parse_cell_ref(range_str)
                abs_row = start_row + cell_row
                abs_col = start_col + cell_col
                abs_range = ExcelRange.cell_to_a1(abs_row, abs_col)
            elif ':' in range_str and '!' not in range_str:
                # Es un rango sin hoja específica, añadir offset
                range_start, range_end = range_str.split(':')
                start_row_rel, start_col_rel = ExcelRange.parse_cell_ref(range_start)
                end_row_rel, end_col_rel = ExcelRange.parse_cell_ref(range_end)
                
                # Calcular posiciones absolutas
                abs_start_row = start_row + start_row_rel
                abs_start_col = start_col + start_col_rel
                abs_end_row = start_row + end_row_rel
                abs_end_col = start_col + end_col_rel
                
                # Crear rango absoluto
                abs_range = ExcelRange.range_to_a1(abs_start_row, abs_start_col, abs_end_row, abs_end_col)
            else:
                # Ya es un rango absoluto o con hoja específica
                abs_range = range_str
            
            # Aplicar formato según tipo
            if isinstance(format_value, str):
                # Es un formato numérico
                apply_number_format(ws, abs_range, format_value)
            elif isinstance(format_value, dict):
                # Es un diccionario de estilos
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
    """Genera un gráfico a partir de una tabla existente.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    La tabla debe contener encabezados válidos y no incluir filas de totales.
    Se asume que las celdas de datos forman un rango rectangular sin valores en
    blanco. Cuando ``use_headers`` es ``True`` la primera fila de la tabla se
    toma como títulos de las series y como categorías. Todas las columnas de
    datos deben ser numéricas y de igual longitud para evitar errores al crear
    el gráfico.

    Revisa que la tabla no contenga celdas vacías ni columnas de texto donde se
    esperan números. Cualquier discrepancia en la longitud de las series o en la
    cantidad de categorías puede provocar gráficos incompletos o vacíos.

    Args:
        wb: Objeto ``Workbook`` de openpyxl.
        sheet_name: Nombre de la hoja donde está la tabla.
        table_name: Nombre de la tabla a utilizar como origen.
        chart_type: Tipo de gráfico (``'column'``, ``'bar'``, ``'line'``, ``'pie'``,
            etc.).
        title: Título opcional del gráfico.
        position: Celda de anclaje para el gráfico.
        style: Estilo del gráfico (número ``1``–``48`` o nombre descriptivo).
        use_headers: Si ``True`` toma la primera fila como encabezados y
            categorías.

    Returns:
        Tupla ``(ID del gráfico, objeto gráfico)``.
    """
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Obtener información de la tabla
    tables = list_tables(wb, sheet_name)
    table_info = None
    for table in tables:
        if table['name'] == table_name:
            table_info = table
            break
    
    if not table_info:
        raise TableError(f"No se encontró la tabla '{table_name}' en la hoja '{sheet_name}'")
    
    # Usar el rango de la tabla para crear el gráfico
    table_range = table_info['ref']
    
    # Crear el gráfico
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
    """Crea un gráfico a partir de ``data`` escribiendo primero los datos.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    ``data`` debe ser una lista de listas con una estructura rectangular y sin
    celdas vacías. La primera fila o columna se interpreta como encabezados y
    categorías; por ello todas las filas deben tener la misma longitud y las
    columnas numéricas no deben contener texto. Evita incluir filas de totales o
    registros que no deban graficarse.

    Antes de llamar a la función verifica que no existan celdas en blanco,
    encabezados duplicados ni longitudes desiguales entre categorías y series.
    ``add_chart`` utilizará ``titles_from_data=True`` para asignar los nombres de
    serie. Si las series no son coherentes, el gráfico resultante podría quedar
    incompleto o mostrar errores.

    Args:
        wb: Objeto ``Workbook`` de openpyxl.
        sheet_name: Nombre de la hoja donde crear el gráfico.
        data: Matriz de datos que incluye los encabezados.
        chart_type: Tipo de gráfico (``'column'``, ``'bar'``, ``'line'``,
            ``'pie'``, etc.).
        position: Celda donde colocar el gráfico.
        title: Título del gráfico.
        style: Estilo del gráfico (número ``1``–``48`` o nombre descriptivo).
        create_table: Si ``True`` crea una tabla con los datos escritos.
        table_name: Nombre de la tabla (obligatorio si ``create_table`` es
            ``True``).
        table_style: Estilo opcional para la tabla.

    Returns:
        Diccionario con información del gráfico y, en su caso, de la tabla
        creada.
    """
    # Crear hoja si no existe
    if sheet_name not in list_sheets(wb):
        add_sheet(wb, sheet_name)
    
    # Obtener la hoja
    ws = get_sheet(wb, sheet_name)
    
    # Determinar una ubicación adecuada para los datos
    # Por defecto, colocar los datos en A1
    data_start_cell = "A1"
    
    # Escribir los datos
    write_sheet_data(ws, data_start_cell, data)
    
    # Calcular el rango completo de los datos
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
    
    # Crear tabla si se solicita
    if create_table:
        if not table_name:
            # Generar nombre de tabla si no se proporciona
            table_name = f"Tabla_{sheet_name}_{int(time.time())}"
            
        try:
            table = add_table(ws, table_name, data_range, table_style)
            result["table"] = {
                "name": table_name,
                "range": data_range,
                "style": table_style
            }
        except Exception as e:
            logger.warning(f"No se pudo crear la tabla: {e}")
    
    # Crear el gráfico
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
        logger.error(f"Error al crear gráfico: {e}")
        raise ChartError(f"Error al crear gráfico: {e}")

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
    """Genera un gráfico a partir de un ``DataFrame`` de pandas.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    El ``DataFrame`` debe contener columnas numéricas sin valores faltantes en
    las series y no incluir filas de totales. Los encabezados se utilizan como
    títulos de las series, por lo que es importante que no haya duplicados ni
    celdas en blanco. El contenido del ``DataFrame`` se escribe en la hoja y se
    delega a :func:`create_chart_from_data`, por lo que aplican las mismas
    recomendaciones sobre validación previa.

    Args:
        wb: Objeto ``Workbook`` de openpyxl.
        sheet_name: Nombre de la hoja donde crear el gráfico.
        df: Datos en formato ``pandas.DataFrame``.
        chart_type: Tipo de gráfico (``'column'``, ``'bar'``, ``'line'``,
            ``'pie'``, etc.).
        position: Celda de anclaje para el gráfico.
        title: Título del gráfico.
        style: Estilo opcional del gráfico.
        create_table: Si ``True`` crea una tabla con los datos escritos.
        table_name: Nombre de la tabla a crear.
        table_style: Estilo de la tabla.

    Returns:
        Diccionario con información del gráfico y, en su caso, de la tabla
        generada.
    """

    if df is None:
        raise ExcelMCPError("El DataFrame proporcionado es None")

    # Convertir el DataFrame a lista de listas incluyendo encabezados
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
    Crea un informe completo con múltiples hojas, tablas y gráficos en un solo paso.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**


    Esta función sirve como plantilla general para generadores automáticos de
    informes. Todas las hojas creadas deben quedar ordenadas y con estilos
    aplicados. Se recomienda verificar el espacio libre antes de insertar
    gráficos para que no queden encima de ninguna tabla o bloque de texto. Tras
    crear una tabla, comprueba qué columna tiene cadenas más largas y ajusta su
    ancho para que el contenido sea visible sin necesidad de editar manualmente
    el archivo.

    Args:
        wb: Objeto workbook de openpyxl
        data: Diccionario con datos por hoja: {"Hoja1": [[datos]], "Hoja2": [[datos]]}
        tables: Diccionario de configuración de tablas:
            {"TablaVentas": {"sheet": "Ventas", "range": "A1:B10", "style": "TableStyleMedium9"}}
        charts: Diccionario de configuración de gráficos:
            {"GraficoVentas": {"sheet": "Ventas", "type": "column", "data": "TablaVentas", 
                              "title": "Ventas", "position": "D2", "style": "dark-blue"}}
        formats: Diccionario de formatos a aplicar:
            {"Ventas": {"B2:B10": "#,##0.00", "A1:Z1": {"bold": True}}}
        overwrite_sheets: Si True, sobrescribe hojas existentes
        
    Returns:
        Diccionario con información de los elementos creados
    """
    result = {
        "sheets": [],
        "tables": [],
        "charts": []
    }
    
    # Crear/actualizar hojas con datos
    for sheet_name, sheet_data in data.items():
        if sheet_name in list_sheets(wb):
            if overwrite_sheets:
                # Usar la hoja existente
                ws = wb[sheet_name]
                # Escribir los datos
                write_sheet_data(ws, "A1", sheet_data)
            else:
                # Añadir sufijo numérico si la hoja ya existe
                base_name = sheet_name
                counter = 1
                while f"{base_name}_{counter}" in list_sheets(wb):
                    counter += 1
                new_name = f"{base_name}_{counter}"
                ws = create_sheet_with_data(wb, new_name, sheet_data)
                sheet_name = new_name
        else:
            # Crear nueva hoja
            ws = create_sheet_with_data(wb, sheet_name, sheet_data)
        
        result["sheets"].append({"name": sheet_name, "rows": len(sheet_data)})
        
        # Aplicar formatos específicos para esta hoja
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
                logger.warning(f"Configuración incompleta para tabla '{table_name}'. Se requiere sheet y range.")
                continue
            
            try:
                # Verificar que la hoja existe
                if sheet_name not in list_sheets(wb):
                    logger.warning(f"Hoja '{sheet_name}' no encontrada para la tabla '{table_name}'. Omitiendo.")
                    continue
                
                ws = wb[sheet_name]
                table = add_table(ws, table_name, range_str, style)
                
                result["tables"].append({
                    "name": table_name,
                    "sheet": sheet_name,
                    "range": range_str,
                    "style": style
                })
                
                # Aplicar formatos específicos para esta tabla
                if "formats" in table_config:
                    for range_str, format_value in table_config["formats"].items():
                        if isinstance(format_value, str):
                            apply_number_format(ws, range_str, format_value)
                        elif isinstance(format_value, dict):
                            apply_style(ws, range_str, format_value)
            
            except Exception as e:
                logger.warning(f"Error al crear tabla '{table_name}': {e}")
    
    # Crear gráficos
    if charts:
        for chart_name, chart_config in charts.items():
            sheet_name = chart_config.get("sheet")
            chart_type = chart_config.get("type")
            data_source = chart_config.get("data")
            title = chart_config.get("title", chart_name)
            position = chart_config.get("position")
            style = chart_config.get("style")
            
            if not sheet_name or not chart_type or not data_source:
                logger.warning(f"Configuración incompleta para gráfico '{chart_name}'. Se requiere sheet, type y data.")
                continue
            
            try:
                # Verificar que la hoja existe
                if sheet_name not in list_sheets(wb):
                    logger.warning(f"Hoja '{sheet_name}' no encontrada para el gráfico '{chart_name}'. Omitiendo.")
                    continue
                
                # Determinar si data_source es una tabla o un rango
                data_range = data_source
                if data_source in [t["name"] for t in result["tables"]]:
                    # Es una tabla, obtener su rango
                    for table in result["tables"]:
                        if table["name"] == data_source:
                            data_range = table["range"]
                            break
                
                # Crear el gráfico
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
                logger.warning(f"Error al crear gráfico '{chart_name}': {e}")
    
    return result

def create_dashboard(wb: Any, dashboard_config: Dict[str, Any],
                    create_new: bool = True) -> Dict[str, Any]:
    """
    Crea un dashboard completo con tablas, gráficos y filtros interactivos.

     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    Está pensado para que un agente automático construya una hoja atractiva y
    sin solapamientos. Coloca cada gráfico dejando espacio respecto a tablas o
    textos previos. Tras escribir los datos de cada sección revisa el tamaño de
    las columnas y amplíalas cuando alguna celda sea especialmente larga. De esa
    forma se garantiza que la lectura sea cómoda sin modificar manualmente el
    archivo.

    Args:
        wb: Objeto workbook de openpyxl
        dashboard_config: Diccionario con configuración completa del dashboard
            {
                "title": "Dashboard de Ventas",
                "sheet": "Dashboard",
                "data_sheet": "Datos",
                "data": [[datos]],
                "sections": [
                    {
                        "title": "Ventas por Región",
                        "type": "chart",
                        "chart_type": "column",
                        "data_range": "A1:B10",
                        "position": "E1",
                        "style": "dark-blue"
                    },
                    {
                        "title": "Tabla de Productos",
                        "type": "table",
                        "data_range": "D1:F10",
                        "name": "TablaProductos",
                        "style": "TableStyleMedium9"
                    }
                ]
            }
        create_new: Si True, crea una nueva hoja para el dashboard
        
    Returns:
        Diccionario con información de los elementos creados
    """
    # Configuración básica
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
            # Añadir sufijo numérico
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
            # Limpiar datos existentes
            # (Esto podría mejorarse para no borrar todo)
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
    
    # Añadir título al dashboard
    update_cell(ws, "A1", title)
    apply_style(ws, "A1", {
        "font_size": 16,
        "bold": True,
        "alignment": "center"
    })
    

    
    # Espacio después del título
    current_row = 3
    
    # Procesar secciones del dashboard
    sections = dashboard_config.get("sections", [])
    for i, section in enumerate(sections):
        section_type = section.get("type")
        section_title = section.get("title", f"Sección {i+1}")
        
        # Información para el resultado
        section_result = {
            "title": section_title,
            "type": section_type,
            "row": current_row
        }
        
        # Añadir título de sección
        update_cell(ws, f"A{current_row}", section_title)
        apply_style(ws, f"A{current_row}", {
            "font_size": 12,
            "bold": True
        })
        current_row += 1
        
        # Procesar según el tipo de sección
        if section_type == "chart":
            chart_type = section.get("chart_type", "column")
            data_range = section.get("data_range")
            
            # Si el rango no tiene hoja específica, usar la hoja de datos
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
                
                # Avanzar filas según la posición y tamaño estimado del gráfico
                # (esto es una estimación simple)
                current_row += 15
            except Exception as e:
                logger.warning(f"Error al crear gráfico en sección '{section_title}': {e}")
                current_row += 2  # Avanzar unas pocas filas en caso de error
        
        elif section_type == "table":
            table_range = section.get("data_range")
            table_name = section.get("name", f"Tabla_{i}")
            table_style = section.get("style")
            
            # Si el rango no tiene hoja específica, usar la hoja de datos
            if table_range and '!' not in table_range and data_sheet:
                if ' ' in data_sheet or any(c in data_sheet for c in "![]{}?"):
                    full_table_range = f"'{data_sheet}'!{table_range}"
                else:
                    full_table_range = f"{data_sheet}!{table_range}"
            else:
                full_table_range = table_range
                
            try:
                # Extraer datos de la tabla para mostrarlos en el dashboard
                if data_sheet:
                    source_ws = wb[data_sheet]
                    # Extraer rango sin nombre de hoja
                    if '!' in table_range:
                        pure_range = table_range.split('!')[1]
                    else:
                        pure_range = table_range
                    
                    # Leer datos de la fuente
                    table_data = read_sheet_data(wb, data_sheet, pure_range)
                    
                    # Determinar dimensiones
                    table_rows = len(table_data)
                    table_cols = max([len(row) if isinstance(row, list) else 1 for row in table_data], default=0)
                    
                    # Escribir datos en el dashboard
                    write_sheet_data(ws, f"A{current_row}", table_data)
                    
                    # Crear tabla local en el dashboard
                    local_range = f"A{current_row}:{get_column_letter(table_cols)}:{current_row + table_rows - 1}"
                    table = add_table(ws, table_name, local_range, table_style)
                    
                    section_result["table_name"] = table_name
                    section_result["source_range"] = full_table_range
                    section_result["dashboard_range"] = local_range
                    
                    # Avanzar filas según tamaño de tabla
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
                        current_row += 10  # Valor por defecto si falla el cálculo
            except Exception as e:
                logger.warning(f"Error al crear tabla en sección '{section_title}': {e}")
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
        
        # Añadir la sección al resultado
        result["sections"].append(section_result)
        
        # Espacio entre secciones
        current_row += 1
    
    return result

def apply_excel_template(wb: Any, template_name: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Aplica una plantilla predefinida a un libro de Excel.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:

        wb: Objeto workbook de openpyxl
        template_name (str): Nombre de la plantilla a aplicar (ej. "informe_ventas", "dashboard")
        data: Diccionario con datos específicos para la plantilla
        
    Returns:
        Diccionario con información de los elementos creados
    
    Plantillas disponibles:
        - "basic_report": Informe básico con tabla y gráfico
        - "financial_dashboard": Dashboard financiero con múltiples KPIs y gráficos
        - "sales_analysis": Análisis de ventas por región y producto
        - "project_tracker": Seguimiento de proyectos con tablas y gráficos de progreso
    """
    result = {
        "template": template_name,
        "sheets": [],
        "elements": []
    }
    
    # Implementación de plantillas predefinidas
    if template_name == "basic_report":
        # Plantilla de informe básico
        title = data.get("title", "Informe Básico")
        subtitle = data.get("subtitle", "")
        report_date = data.get("date", time.strftime("%d/%m/%Y"))
        sheet_name = data.get("sheet", "Informe")
        report_data = data.get("data", [])
        
        # Crear hoja para el informe si no existe
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Título e información básica
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
            table_name = data.get("table_name", "TablaInforme")
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
            
            # Crear gráfico
            chart_type = data.get("chart_type", "column")
            chart_position = data.get("chart_position", f"G{start_row}")
            chart_title = data.get("chart_title", "Gráfico del Informe")
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
                logger.warning(f"Error al crear gráfico: {e}")
        
        result["sheets"].append({"name": sheet_name, "type": "report"})
    
    elif template_name == "financial_dashboard":
        # Template más avanzado para dashboard financiero
        title = data.get("title", "Dashboard Financiero")
        sheet_name = data.get("sheet", "Dashboard")
        financial_data = data.get("financial_data", {})
        
        # Configuración para crear el dashboard completo
        dashboard_config = {
            "title": title,
            "sheet": sheet_name,
            "sections": []
        }
        
        # 1. Sección de KPIs financieros
        if "kpis" in financial_data:
            kpis = financial_data["kpis"]
            kpi_section = {
                "title": "Indicadores Financieros Clave",
                "type": "text",
                "content": "KPIs Financieros"
            }
            dashboard_config["sections"].append(kpi_section)
            
            # Cada KPI se podría añadir como texto o celda con formato
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
        
        # 2. Sección de gráficos financieros
        if "charts" in financial_data:
            for chart_config in financial_data["charts"]:
                chart_section = {
                    "title": chart_config.get("title", "Gráfico Financiero"),
                    "type": "chart",
                    "chart_type": chart_config.get("type", "column"),
                    "data_range": chart_config.get("data_range", ""),
                    "position": chart_config.get("position", ""),
                    "style": chart_config.get("style", "dark-blue")
                }
                dashboard_config["sections"].append(chart_section)
        
        # 3. Sección de tablas de datos
        if "tables" in financial_data:
            for table_config in financial_data["tables"]:
                table_section = {
                    "title": table_config.get("title", "Tabla Financiera"),
                    "type": "table",
                    "data_range": table_config.get("data_range", ""),
                    "name": table_config.get("name", "TablaFinanzas"),
                    "style": table_config.get("style", "TableStyleMedium9")
                }
                dashboard_config["sections"].append(table_section)
        
        # Crear el dashboard
        dashboard_result = create_dashboard(wb, dashboard_config)
        
        # Añadir resultado
        result["sheets"].append({"name": sheet_name, "type": "dashboard"})
        result["dashboard"] = dashboard_result
    
    elif template_name == "sales_analysis":
        # Template para análisis de ventas
        title = data.get("title", "Análisis de Ventas")
        sheet_data = data.get("sales_data", [])
        sheet_name = data.get("sheet", "Ventas")
        
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
                table = add_table(data_ws, "TablaDatosVentas", data_range, "TableStyleMedium9")
                result["elements"].append({
                    "type": "table",
                    "name": "TablaDatosVentas",
                    "sheet": data_sheet,
                    "range": data_range
                })
            except Exception as e:
                logger.warning(f"Error al crear tabla de datos: {e}")
        
        # Crear hoja de análisis
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Título del análisis
        update_cell(ws, "A1", title)
        apply_style(ws, "A1", {
            "font_size": 16,
            "bold": True,
            "alignment": "center"
        })
        

            
        # Crear secciones de análisis según la estructura de los datos
        current_row = 3
        
        # 1. Ventas por Región (suponiendo que hay una columna de región)
        update_cell(ws, f"A{current_row}", "Ventas por Región")
        apply_style(ws, f"A{current_row}", {"bold": True, "font_size": 12})
        current_row += 1
        
        try:
            # Crear gráfico para ventas por región
            chart_id, chart = add_chart(wb, sheet_name, "column", 
                                       f"{data_sheet}!A1:{get_column_letter(cols)}{rows}", 
                                       "Ventas por Región", f"A{current_row}", "colorful-1")
            
            result["elements"].append({
                "type": "chart",
                "name": "GraficoVentasRegion",
                "sheet": sheet_name,
                "id": chart_id
            })
            
            current_row += 15  # Espacio para el gráfico
        except Exception as e:
            logger.warning(f"Error al crear gráfico de ventas por región: {e}")
            current_row += 2
        
        # 2. Tendencia de Ventas (si hay datos temporales)
        update_cell(ws, f"A{current_row}", "Tendencia de Ventas")
        apply_style(ws, f"A{current_row}", {"bold": True, "font_size": 12})
        current_row += 1
        
        try:
            # Crear gráfico para tendencia de ventas
            chart_id, chart = add_chart(wb, sheet_name, "line", 
                                       f"{data_sheet}!A1:{get_column_letter(cols)}{rows}", 
                                       "Tendencia de Ventas", f"A{current_row}", "line-markers")
            
            result["elements"].append({
                "type": "chart",
                "name": "GraficoTendenciaVentas",
                "sheet": sheet_name,
                "id": chart_id
            })
            
            current_row += 15  # Espacio para el gráfico
        except Exception as e:
            logger.warning(f"Error al crear gráfico de tendencia de ventas: {e}")
            current_row += 2
        
        result["sheets"].append({"name": sheet_name, "type": "analysis"})
        result["sheets"].append({"name": data_sheet, "type": "data"})
        
    elif template_name == "project_tracker":
        # Template para seguimiento de proyectos
        title = data.get("title", "Seguimiento de Proyectos")
        projects = data.get("projects", [])
        sheet_name = data.get("sheet", "Proyectos")
        
        # Preparar datos de proyectos
        if not projects:
            # Crear datos de ejemplo si no se proporcionan
            projects = [
                ["ID", "Proyecto", "Responsable", "Inicio", "Plazo", "Estado", "Avance"],
                ["P001", "Proyecto A", "Juan Pérez", "01/01/2023", "30/06/2023", "En curso", 75],
                ["P002", "Proyecto B", "Ana López", "15/02/2023", "31/07/2023", "En curso", 40],
                ["P003", "Proyecto C", "Carlos Ruiz", "01/03/2023", "31/08/2023", "Retrasado", 20]
            ]
        
        # Crear hoja para proyectos si no existe
        if sheet_name not in list_sheets(wb):
            ws = add_sheet(wb, sheet_name)
        else:
            ws = wb[sheet_name]
        
        # Título del tracker
        update_cell(ws, "A1", title)
        apply_style(ws, "A1", {
            "font_size": 16,
            "bold": True,
            "alignment": "center"
        })
        

        
        # Escribir datos de proyectos
        write_sheet_data(ws, "A3", projects)
        
        # Crear tabla para los datos
        rows = len(projects)
        cols = len(projects[0]) if rows > 0 else 7
        table_range = f"A3:{get_column_letter(cols)}{rows+2}"
        
        try:
            table = add_table(ws, "TablaProyectos", table_range, "TableStyleMedium9")
            result["elements"].append({
                "type": "table",
                "name": "TablaProyectos",
                "sheet": sheet_name,
                "range": table_range
            })
            
            # Aplicar formato porcentual a la columna de avance
            avance_col = get_column_letter(cols)
            apply_number_format(ws, f"{avance_col}4:{avance_col}{rows+2}", "0%")
        except Exception as e:
            logger.warning(f"Error al crear tabla de proyectos: {e}")
        
        # Crear gráfico de avance
        try:
            chart_id, chart = add_chart(wb, sheet_name, "column", 
                                       table_range, 
                                       "Avance de Proyectos", "I3", "colorful-3")
            
            result["elements"].append({
                "type": "chart",
                "name": "GraficoAvance",
                "sheet": sheet_name,
                "id": chart_id
            })
        except Exception as e:
            logger.warning(f"Error al crear gráfico de avance: {e}")
        
        result["sheets"].append({"name": sheet_name, "type": "tracker"})
    
    else:
        logger.warning(f"Plantilla '{template_name}' no reconocida.")
        result["error"] = f"Plantilla '{template_name}' no disponible"
    
    return result

def update_report(wb: Any, report_config: Dict[str, Any], 
                 recalculate: bool = True) -> Dict[str, Any]:
    """
    Actualiza un informe existente con nuevos datos.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        wb: Objeto workbook de openpyxl
        report_config: Configuración del informe a actualizar
            {
                "data_updates": {
                    "Ventas": {"range": "A2:C10", "data": [[nuevos datos]]},
                    "Clientes": {"range": "A2:D20", "data": [[nuevos datos]]}
                },
                "recalculate_formulas": True,
                "refresh_charts": True
            }
        recalculate: Si True, recalcula fórmulas después de actualizar
        
    Returns:
        Diccionario con información de los elementos actualizados
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
            logger.warning(f"Hoja '{sheet_name}' no encontrada. Omitiendo actualización.")
            continue
        
        ws = wb[sheet_name]
        range_str = update_info.get("range")
        data = update_info.get("data")
        
        if not range_str or not data:
            logger.warning(f"Configuración incompleta para actualizar hoja '{sheet_name}'. Se requiere range y data.")
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
            logger.warning("Información de tabla incompleta. Se requiere sheet y name.")
            continue
        
        if sheet_name not in list_sheets(wb):
            logger.warning(f"Hoja '{sheet_name}' no encontrada. Omitiendo actualización de tabla.")
            continue
        
        ws = wb[sheet_name]
        
        try:
            # Verificar si la tabla existe
            if not hasattr(ws, 'tables') or table_name not in ws.tables:
                logger.warning(f"Tabla '{table_name}' no encontrada en hoja '{sheet_name}'.")
                continue
            
            # Obtener referencia actual
            current_range = ws.tables[table_name].ref
            
            # Actualizar rango si se proporciona uno nuevo
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
            logger.warning(f"Error al actualizar tabla '{table_name}': {e}")
    
    # Recalcular fórmulas si se solicita
    if recalculate:
        # En OpenPyXL no hay una funcionalidad directa para recalcular fórmulas
        # Esta es una función de placeholder que podría implementarse en versiones futuras
        # o mediante la API COM de Excel si está disponible
        result["recalculation_note"] = "Recalcular fórmulas en OpenPyXL es limitado"
    
    # Actualizar gráficos
    refresh_charts = report_config.get("refresh_charts", [])
    for chart_info in refresh_charts:
        sheet_name = chart_info.get("sheet")
        chart_id = chart_info.get("id")
        new_data_range = chart_info.get("new_data_range")
        
        if not sheet_name or chart_id is None:
            logger.warning("Información de gráfico incompleta. Se requiere sheet y id.")
            continue
        
        if sheet_name not in list_sheets(wb):
            logger.warning(f"Hoja '{sheet_name}' no encontrada. Omitiendo actualización de gráfico.")
            continue
        
        ws = wb[sheet_name]
        
        try:
            # Verificar si el gráfico existe
            if not hasattr(ws, '_charts') or chart_id >= len(ws._charts) or chart_id < 0:
                logger.warning(f"Gráfico con ID {chart_id} no encontrado en hoja '{sheet_name}'.")
                continue
            
            # En OpenPyXL, actualizar un gráfico no es tan directo
            # Una opción sería eliminar el gráfico y crear uno nuevo
            if new_data_range:
                # Obtener propiedades del gráfico actual
                chart_rel = ws._charts[chart_id]
                chart = chart_rel[0]
                position = chart_rel[1] if len(chart_rel) > 1 else None
                
                # Determinar el tipo de gráfico
                chart_type = "column"  # Valor por defecto
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
                
                # Obtener título si existe
                title = chart.title if hasattr(chart, 'title') and chart.title else None
                
                # Eliminar el gráfico viejo
                del ws._charts[chart_id]
                
                # Crear un nuevo gráfico con los mismos parámetros pero nuevo rango
                new_chart_id, new_chart = add_chart(wb, sheet_name, chart_type, new_data_range,
                                                 title, position)
                
                result["updated_charts"].append({
                    "id": chart_id,
                    "new_id": new_chart_id,
                    "sheet": sheet_name,
                    "old_data_range": "unknown",  # No hay forma fácil de obtener el rango original
                    "new_data_range": new_data_range
                })
            else:
                # Sin nuevo rango, no se puede actualizar fácilmente
                result["updated_charts"].append({
                    "id": chart_id,
                    "sheet": sheet_name,
                    "note": "No se proporcionó nuevo rango. La actualización real de datos requiere Excel COM."
                })
        except Exception as e:
            logger.warning(f"Error al actualizar gráfico {chart_id}: {e}")
    
    return result

def import_data(wb: Any, import_config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Importa datos de distintas fuentes a Excel.
    
    Args:
        wb: Objeto workbook de openpyxl
        import_config: Configuración de la importación
            {
                "source": "csv", // csv, json, pandas, etc.
                "source_path": "datos.csv",
                "sheet": "Datos",
                "start_cell": "A1",
                "options": {
                    "delimiter": ",",
                    "has_header": true
                }
            }
        
    Returns:
        Diccionario con información de los datos importados
    
    Nota: Esta función es un ejemplo simplificado que solo importa datos desde CSV.
    """
    result = {
        "source": import_config.get("source"),
        "imported_rows": 0,
        "imported_columns": 0
    }
    
    source_type = import_config.get("source", "").lower()
    source_path = import_config.get("source_path")
    sheet_name = import_config.get("sheet", "Datos")
    start_cell = import_config.get("start_cell", "A1")
    options = import_config.get("options", {})
    
    if not source_path:
        logger.warning("No se especificó una ruta de origen para importar datos.")
        result["error"] = "No se especificó una ruta de origen"
        return result
        
    # Crear la hoja si no existe
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
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error al importar CSV: {e}")
            result["error"] = f"Error al importar CSV: {e}"
    
    elif source_type == "json":
        try:
            import json
            
            with open(source_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # Convertir JSON a lista de listas
            data = []
            
            if isinstance(json_data, list):
                # Es una lista de objetos
                if json_data and isinstance(json_data[0], dict):
                    # Obtener encabezados (claves del primer objeto)
                    headers = list(json_data[0].keys())
                    data.append(headers)
                    
                    # Añadir filas de datos
                    for item in json_data:
                        row = [item.get(header, "") for header in headers]
                        data.append(row)
                else:
                    # Es una lista simple
                    for item in json_data:
                        data.append([item])
            elif isinstance(json_data, dict):
                # Es un diccionario
                for key, value in json_data.items():
                    data.append([key, value])
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error al importar JSON: {e}")
            result["error"] = f"Error al importar JSON: {e}"
    
    elif source_type == "pandas":
        try:
            import pandas as pd
            
            # Opciones para pandas
            file_ext = os.path.splitext(source_path)[1].lower()
            
            if file_ext == '.csv':
                df = pd.read_csv(source_path)
            elif file_ext in ['.xls', '.xlsx']:
                df = pd.read_excel(source_path)
            elif file_ext == '.json':
                df = pd.read_json(source_path)
            else:
                raise ValueError(f"Formato de archivo no soportado: {file_ext}")
            
            # Convertir DataFrame a lista de listas
            data = [df.columns.tolist()]  # Encabezados
            data.extend(df.values.tolist())  # Datos
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            result["imported_rows"] = len(data)
            result["imported_columns"] = len(data[0]) if data else 0
            result["sheet"] = sheet_name
            result["start_cell"] = start_cell
        except Exception as e:
            logger.error(f"Error al importar con pandas: {e}")
            result["error"] = f"Error al importar con pandas: {e}"
    
    else:
        logger.warning(f"Tipo de fuente no soportado: {source_type}")
        result["error"] = f"Tipo de fuente no soportado: {source_type}"
    
    return result

def export_data(wb: Any, export_config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Exporta datos de Excel a distintos formatos.
    
    Args:
        wb: Objeto workbook de openpyxl
        export_config: Configuración de la exportación
            {
                "format": "csv", // csv, json, pdf, html, etc.
                "sheet": "Datos",
                "range": "A1:D10",
                "output_path": "datos_exportados.csv",
                "options": {
                    "delimiter": ",",
                    "include_header": true
                }
            }
        
    Returns:
        Diccionario con información de los datos exportados
    
    Nota: Esta función es un ejemplo simplificado que solo exporta a CSV y JSON.
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
        logger.warning("No se especificó una hoja para exportar datos.")
        result["error"] = "No se especificó una hoja"
        return result
        
    if sheet_name not in list_sheets(wb):
        logger.warning(f"Hoja '{sheet_name}' no encontrada.")
        result["error"] = f"Hoja '{sheet_name}' no encontrada"
        return result
    
    # Leer los datos del rango especificado
    data = read_sheet_data(wb, sheet_name, range_str)
    
    if not data:
        logger.warning(f"No se encontraron datos en el rango {range_str} de la hoja {sheet_name}")
        return []
    
    # Filtrar los datos según los criterios
    result = []
    headers = data[0] if data else []
    
    # Si tenemos datos con encabezados
    if len(data) > 1:
        # Convertir a formato de registros (lista de diccionarios)
        records = []
        for row in data[1:]:
            record = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    record[header] = row[i]
                else:
                    record[header] = None
            records.append(record)
        
        # Aplicar filtros si se proporcionan
        if filters:
            filtered_records = []
            for record in records:
                include = True
                for field, value in filters.items():
                    if field in record:
                        # Si el valor del filtro es una lista, verificar si el valor está en la lista
                        if isinstance(value, list):
                            if record[field] not in value:
                                include = False
                                break
                        # Si el valor del filtro es un diccionario, aplicar operadores
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
                        # Si el valor del filtro es un valor simple, hacer una comparación de igualdad
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
    Crea un informe basado en una plantilla Excel, sustituyendo datos, actualizando gráficos y aplicando formatos.
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    
    Args:
        template_file (str): Ruta a la plantilla Excel

        output_file (str): Ruta donde guardar el informe generado
        data_mappings (dict): Diccionario con mapeos de datos:
            {
                "sheet_name": {
                    "range1": data_list1,
                    "range2": data_list2,
                    ...
                }
            }
        chart_mappings (dict, opcional): Diccionario con actualizaciones de gráficos:
            {
                "sheet_name": {
                    "chart_id": {
                        "title": "Nuevo título",
                        "data_range": "Nuevo rango",
                        ...
                    }
                }
            }
        format_mappings (dict, opcional): Diccionario con formatos a aplicar:
            {
                "sheet_name": {
                    "range1": {"number_format": "#,##0.00"},
                    "range2": {"style": {"bold": True, "fill_color": "FFFF00"}},
                    ...
                }
            }
    
    Returns:
        dict: Resultado de la operación
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
                    logger.warning(f"La hoja '{sheet_name}' no existe en la plantilla")
                    continue
                
                ws = wb[sheet_name]
                for range_str, data in ranges.items():
                    # Si el rango es una sola celda, extraer la celda de inicio
                    if ':' not in range_str:
                        start_cell = range_str
                    else:
                        start_cell = range_str.split(':')[0]
                    
                    # Escribir los datos
                    write_sheet_data(ws, start_cell, data)
        
        # Aplicar mapeos de gráficos
        if chart_mappings:
            for sheet_name, charts in chart_mappings.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"La hoja '{sheet_name}' no existe en la plantilla")
                    continue
                
                ws = wb[sheet_name]
                existing_charts = list_charts(ws)
                
                for chart_id, chart_updates in charts.items():
                    # Verificar si el chart_id es un índice o un nombre
                    chart_idx = None
                    if isinstance(chart_id, int) or (isinstance(chart_id, str) and chart_id.isdigit()):
                        chart_idx = int(chart_id)
                    else:
                        # Buscar el chart por título
                        for i, chart in enumerate(existing_charts):
                            if chart.get('title') == chart_id:
                                chart_idx = i
                                break
                    
                    if chart_idx is None or chart_idx >= len(existing_charts):
                        logger.warning(f"No se encontró el gráfico '{chart_id}' en la hoja '{sheet_name}'")
                        continue
                    
                    # Actualizar propiedades del gráfico
                    chart = ws._charts[chart_idx][0]
                    
                    if 'title' in chart_updates:
                        chart.title = chart_updates['title']
                    
                    if 'data_range' in chart_updates:
                        # La actualización del rango de datos es compleja y depende del tipo de gráfico
                        # Por ahora, simplemente logueamos que esta funcionalidad no está implementada
                        logger.warning("La actualización del rango de datos de gráficos no está implementada completamente")
        
        # Aplicar mapeos de formato
        if format_mappings:
            for sheet_name, ranges in format_mappings.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"La hoja '{sheet_name}' no existe en la plantilla")
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
    Crea un dashboard dinámico con múltiples visualizaciones en un solo paso.

    Los modelos que utilicen esta función deben procurar que las tablas y los
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    gráficos no se solapen. Es aconsejable dejar filas de separación y comprobar
    el ancho necesario de cada columna tras escribir los datos. Aplicar estilos
    coherentes ayuda a que el resultado sea más limpio y profesional.

    Args:
        file_path (str): Ruta al archivo Excel a crear o modificar
        data (dict): Diccionario con datos por hoja:
            {
                "sheet_name": [
                    ["Header1", "Header2", ...],
                    [value1, value2, ...],
                    ...
                ]
            }
        dashboard_config (dict): Configuración del dashboard:
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
        overwrite (bool): Si es True, sobrescribe el archivo si existe
    
    Returns:
        dict: Resultado de la operación
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
                logger.warning(f"La hoja '{sheet_name}' no existe para crear la tabla '{table_name}'")
                continue
            
            ws = wb[sheet_name]
            
            # Verificar si la tabla ya existe
            table_exists = False
            if hasattr(ws, 'tables') and table_name in ws.tables:
                table_exists = True
                logger.warning(f"La tabla '{table_name}' ya existe, se actualizará")
            
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
                        # Es un formato numérico
                        apply_number_format(ws, cell_range, fmt)
                    elif isinstance(fmt, dict):
                        # Es un estilo
                        apply_style(ws, cell_range, fmt)
        
        # Crear gráficos
        for chart_config in dashboard_config.get("charts", []):
            sheet_name = chart_config["sheet"]
            chart_type = chart_config["type"]
            data_range = chart_config["data_range"]
            title = chart_config.get("title", f"Chart {len(wb[sheet_name]._charts) + 1}")
            position = chart_config.get("position", "E1")
            style = chart_config.get("style")
            
            if sheet_name not in wb.sheetnames:
                logger.warning(f"La hoja '{sheet_name}' no existe para crear el gráfico '{title}'")
                continue
            
            # Crear gráfico
            chart_id, _ = add_chart(wb, sheet_name, chart_type, data_range, title, position, style)
        
        # Configurar tamaños de columna para visualización óptima
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Establecer un ancho mínimo para columnas con fechas
            for i in range(1, ws.max_column + 1):
                column_letter = get_column_letter(i)
                # Verificar si hay celdas con formato de fecha en la columna
                date_format = False
                for cell in ws[column_letter]:
                    if cell.number_format and ('yy' in cell.number_format.lower() or 'mm' in cell.number_format.lower() or 'dd' in cell.number_format.lower()):
                        date_format = True
                        break
                
                if date_format:
                    # Establecer ancho mínimo para columnas con fechas
                    ws.column_dimensions[column_letter].width = max(ws.column_dimensions[column_letter].width or 0, 10)
        
        # Guardar el archivo
        wb.save(file_path)
        
        return {
            "success": True,
            "file_path": file_path,
            "message": f"Dashboard creado correctamente: {file_path}"
        }
    
    except Exception as e:
        logger.error(f"Error al crear dashboard: {e}")
        return {
            "success": False,
            "error": str(e),
            "message": f"Error al crear dashboard: {e}"
        }

def import_multi_source_data(excel_file, import_config, sheet_name=None, start_cell="A1", create_tables=False):
    """
    Importa datos desde múltiples fuentes (CSV, JSON, SQL) a un archivo Excel en un solo paso.
    
     **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

    Args:
        excel_file (str): Ruta al archivo Excel donde importar los datos
        import_config (dict): Configuración de importación:
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
        sheet_name (str, opcional): Nombre de hoja predeterminado si no se especifica en la configuración
        start_cell (str, opcional): Celda inicial predeterminada si no se especifica en la configuración
        create_tables (bool, opcional): Si es True, crea tablas Excel para cada conjunto de datos
    
    Returns:
        dict: Resultado de la operación
    """
    try:
        import csv
        import json
        
        # Intentar importar pandas si está disponible (opcional)
        try:
            import pandas as pd
            HAS_PANDAS = True
        except ImportError:
            HAS_PANDAS = False
            logger.warning("Pandas no está disponible. Algunas funcionalidades estarán limitadas.")
        
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
                # Usar pandas si está disponible
                df = pd.read_csv(csv_file, delimiter=delimiter, encoding=encoding)
                data = [df.columns.tolist()]  # Encabezados
                data.extend(df.values.tolist())  # Datos
            else:
                # Usar csv estándar si pandas no está disponible
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
                
                # Crear un nombre único para la tabla
                table_name = f"Table_{csv_sheet}_{len(imported_data) + 1}"
                table_name = table_name.replace(" ", "_")
                
                try:
                    add_table(ws, table_name, table_range, "TableStyleMedium9")
                except Exception as table_error:
                    logger.warning(f"No se pudo crear la tabla para {csv_file}: {table_error}")
            
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
                        # Si el elemento no es un diccionario, añadirlo como una sola columna
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
                    # Si no es un diccionario ni una lista, usar una representación simple
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
                
                # Crear un nombre único para la tabla
                table_name = f"Table_{json_sheet}_{len(imported_data) + 1}"
                table_name = table_name.replace(" ", "_")
                
                try:
                    add_table(ws, table_name, table_range, "TableStyleMedium9")
                except Exception as table_error:
                    logger.warning(f"No se pudo crear la tabla para {json_file}: {table_error}")
            
            imported_data.append({
                "source": "json",
                "file": json_file,
                "sheet": json_sheet,
                "rows": len(data)
            })
        
        # Procesar consultas SQL (requiere conexión a base de datos)
        if "sql" in import_config and import_config["sql"]:
            try:
                import pyodbc
                HAS_PYODBC = True
            except ImportError:
                HAS_PYODBC = False
                logger.warning("pyodbc no está disponible. No se pueden importar datos SQL.")
            
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
                            # Usar pandas si está disponible
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
                            
                            # Crear un nombre único para la tabla
                            table_name = f"Table_{sql_sheet}_{len(imported_data) + 1}"
                            table_name = table_name.replace(" ", "_")
                            
                            try:
                                add_table(ws, table_name, table_range, "TableStyleMedium9")
                            except Exception as table_error:
                                logger.warning(f"No se pudo crear la tabla para consulta SQL: {table_error}")
                        
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
    Exporta datos de Excel a múltiples formatos (CSV, JSON, PDF) en un solo paso.
    
    Args:
        excel_file (str): Ruta al archivo Excel de origen
        export_config (dict): Configuración de exportación:
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
                    "sheets": ["Sheet1", "Sheet2"]  # o null para todas
                }
            }
    
    Returns:
        dict: Resultado de la operación
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
            
            # Escribir los datos en CSV
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
            
            # Convertir datos a formato JSON según el tipo especificado
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
            
            # Escribir los datos en JSON
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
                # Intentar usar win32com para Excel si está disponible
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
                logger.warning("win32com no está disponible. No se puede exportar a PDF.")
                pass  # Si win32com no está disponible, simplemente omitir la exportación PDF
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
    """Exporta un libro de Excel a PDF solo si contiene una única hoja visible.

    Args:
        excel_file: Ruta al archivo Excel a exportar.
        output_pdf: Ruta del archivo PDF resultante. Si no se indica se usa el
            mismo nombre que ``excel_file`` con extensión ``.pdf``.

    Returns:
        dict: Resultado de la operación.
    """
    try:
        import shutil
        import subprocess

        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"El archivo Excel no existe: {excel_file}")

        wb = openpyxl.load_workbook(excel_file, data_only=True)
        visible_sheets = [ws.title for ws in wb.worksheets if getattr(ws, "sheet_state", "visible") == "visible"]

        if len(visible_sheets) != 1:
            msg = f"El archivo debe tener una única hoja visible. Hojas visibles: {len(visible_sheets)}"
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

            msg = f"Archivo exportado correctamente a PDF: {output_pdf}"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_file": output_pdf,
                "message": msg,
            }
        except ImportError:
            logger.info("win32com no disponible, se intentará usar LibreOffice")
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

        msg = "No se encontró un método disponible para exportar a PDF."
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
    """Exporta una o varias hojas de un libro de Excel a PDF.

    Parameters
    ----------
    excel_file : str
        Ruta al archivo Excel a exportar.
    sheets : Union[str, List[str]], optional
        Nombre de la hoja o lista de hojas a exportar. Si ``None`` se exportan
        todas las hojas del libro (una por una).
    output_dir : str, optional
        Carpeta donde guardar los PDF. Por defecto se usa la carpeta del
        archivo original.
    single_file : bool, optional
        Si es ``True`` y se especifican varias hojas se intentará crear un único
        PDF con todas ellas (si el sistema lo permite). Si es ``False`` se
        generará un PDF por cada hoja.

    Returns
    -------
    dict
        Resultado de la operación con la lista de PDFs generados. Si alguna hoja
        no existe se incluye un aviso en ``warnings``.
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
            msg = "No se encontraron hojas válidas para exportar"
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

        # Intentar usar win32com si está disponible
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

            msg = "Exportación a PDF realizada correctamente"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_files": pdf_files,
                "warnings": warnings,
                "message": msg,
            }
        except ImportError:
            logger.info("win32com no disponible, se intentará usar LibreOffice")
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

            msg = "Exportación a PDF realizada correctamente"
            logger.info(msg)
            return {
                "success": True,
                "file_path": excel_file,
                "pdf_files": pdf_files,
                "warnings": warnings,
                "message": msg,
            }

        msg = "No se encontró un método disponible para exportar a PDF."
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
    
    # Registrar funciones básicas de gestión de workbooks
    @mcp.tool(description="Crea un nuevo fichero Excel vacío")
    def create_workbook_tool(filename, overwrite=False):
        """Crea un nuevo fichero Excel vacío
        
        Esta función permite crear un nuevo archivo Excel (.xlsx) vacío en la ubicación especificada.
        Es el primer paso recomendado cuando se quiere generar un nuevo documento desde cero.
        
        Args:
            filename (str): Ruta completa y nombre del archivo a crear. Debe tener extensión .xlsx
            overwrite (bool, optional): Si es True, sobrescribe el archivo si ya existe. Por defecto es False.
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la ruta del archivo creado.
        
        Ejemplo:
            create_workbook_tool("C:/datos/nuevo_libro.xlsx")
        """
        try:
            wb = create_workbook(filename, overwrite)
            save_workbook(wb, filename)
            
            return {
                "success": True,
                "file_path": filename,
                "message": f"Archivo Excel creado correctamente: {filename}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear archivo Excel: {e}"
            }
    
    @mcp.tool(description="Abre un fichero Excel existente")
    def open_workbook_tool(filename):
        """Abre un fichero Excel existente
        
        Esta función permite abrir un archivo Excel (.xlsx, .xls) existente para su manipulación.
        Es necesario usar esta función antes de realizar cualquier operación sobre un archivo existente.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel a abrir.
            
        Returns:
            dict: Información sobre el archivo abierto, incluyendo número de hojas y otras propiedades.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            
        Ejemplo:
            open_workbook_tool("C:/datos/informe_ventas.xlsx")
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
                "message": f"Archivo Excel abierto correctamente: {filename}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al abrir archivo Excel: {e}"
            }
    
    @mcp.tool(description="Guarda el Workbook en disco")
    def save_workbook_tool(filename, new_filename=None):
        """Guarda el Workbook en disco
        
        Esta función permite guardar un archivo Excel que ha sido modificado.
        Es importante llamar a esta función después de realizar cambios para asegurar que estos se persistan.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel a guardar.
            new_filename (str, optional): Si se proporciona, guarda el archivo con un nuevo nombre
                                          (equivalente a 'Guardar como'). Por defecto es None.
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la ruta donde se guardó el archivo.
        
        Ejemplo:
            save_workbook_tool("C:/datos/informe.xlsx")
            save_workbook_tool("C:/datos/informe.xlsx", "C:/datos/informe_backup.xlsx") # Guardar como
        """
        try:
            wb = open_workbook(filename)
            saved_path = save_workbook(wb, new_filename or filename)
            close_workbook(wb)
            
            return {
                "success": True,
                "original_file": filename,
                "saved_file": saved_path,
                "message": f"Archivo Excel guardado correctamente: {saved_path}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al guardar archivo Excel: {e}"
            }
    
    @mcp.tool(description="Lista las hojas disponibles en un archivo Excel")
    def list_sheets_tool(filename):
        """Lista las hojas disponibles en un archivo Excel
        
        Esta función muestra todas las hojas de cálculo que contiene un archivo Excel.
        Es útil para obtener una visión general del contenido del libro antes de trabajar con él.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel a examinar.
            
        Returns:
            dict: Diccionario con la lista de nombres de hojas y sus posiciones en el libro.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            
        Ejemplo:
            list_sheets_tool("C:/datos/informe_financiero.xlsx")  # Devuelve: {"sheets": ["Ventas", "Gastos", "Resumen"]}
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
    
    # Registrar funciones básicas de manipulación de hojas
    @mcp.tool(description="Añade una nueva hoja vacía")
    def add_sheet_tool(filename, sheet_name, index=None):
        """Añade una nueva hoja vacía
        
        Esta función permite agregar una nueva hoja de cálculo vacía a un libro de Excel existente.
        Puedes especificar la posición donde quieres insertar la nueva hoja.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre para la nueva hoja.
            index (int, optional): Posición donde insertar la hoja (0 es la primera posición).
                                 Si es None, se añade al final. Por defecto es None.
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la lista actualizada de hojas.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetExistsError: Si ya existe una hoja con el mismo nombre.
            
        Ejemplo:
            add_sheet_tool("C:/datos/informe.xlsx", "Nuevo Resumen")  # Añade al final
            add_sheet_tool("C:/datos/informe.xlsx", "Portada", 0)  # Añade como primera hoja
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
                "message": f"Hoja '{sheet_name}' añadida correctamente"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al añadir hoja: {e}"
            }
    
    @mcp.tool(description="Elimina la hoja indicada")
    def delete_sheet_tool(filename, sheet_name):
        """
        Elimina la hoja indicada
        
        Esta función permite eliminar una hoja de cálculo específica de un libro Excel.
        Ten cuidado al usar esta función, ya que la eliminación es permanente una vez guardado el archivo.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre de la hoja que se desea eliminar.
            
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la lista actualizada de hojas.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja especificada.
            ValueError: Si se intenta eliminar la única hoja del libro (Excel requiere al menos una hoja).
            
        Ejemplo:
            delete_sheet_tool("C:/datos/informe.xlsx", "Borrador")
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
                "message": f"Hoja '{sheet_name}' eliminada correctamente"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al eliminar hoja: {e}"
            }
    
    @mcp.tool(description="Renombra una hoja")
    def rename_sheet_tool(filename, old_name, new_name):
        """
        Renombra una hoja
        
        Esta función permite cambiar el nombre de una hoja de cálculo existente en un libro Excel.
        
        Args:
            filename (str): Ruta completa y nombre del archivo Excel.
            old_name (str): Nombre actual de la hoja que se desea renombrar.
            new_name (str): Nuevo nombre para la hoja.
            
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la lista actualizada de hojas.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja con el nombre original.
            SheetExistsError: Si ya existe una hoja con el nuevo nombre.
            
        Ejemplo:
            rename_sheet_tool("C:/datos/informe.xlsx", "Hoja1", "Resumen Ejecutivo")
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
                "message": f"Hoja renombrada de '{old_name}' a '{new_name}'"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al renombrar hoja: {e}"
            }
    
    # Registrar funciones básicas de escritura
    @mcp.tool(description="Escribe un array bidimensional de valores o fórmulas")
    def write_sheet_data_tool(file_path, sheet_name, start_cell, data):
        """
        Escribe un array bidimensional de valores o fórmulas en una hoja de Excel
        
        Esta función permite escribir datos en un rango de celdas de una hoja Excel, comenzando desde
        la celda especificada. Es ideal para insertar tablas de datos o matrices de valores.
        
        Args:
            file_path (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre de la hoja donde se escribirán los datos.
            start_cell (str): Celda inicial desde donde comenzar a escribir (ej: "A1").
            data (list): Array bidimensional (lista de listas) con los datos a escribir.
                        Ejemplo: [["Nombre", "Edad"], ["Juan", 25], ["María", 30]]
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo el rango modificado.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja especificada.
            CellReferenceError: Si la referencia de celda no es válida.
            
        Ejemplo:
            write_sheet_data_tool(
                "C:/datos/informe.xlsx", 
                "Datos", 
                "B2", 
                [["Trimestre", "Ventas", "Gastos"], ["Q1", 5000, 3000], ["Q2", 6200, 3100]]
            )
        """
        try:
            # Validar argumentos
            if not isinstance(data, list):
                raise ValueError("El parámetro 'data' debe ser una lista")
            
            # Abrir el archivo y obtener la hoja
            wb = openpyxl.load_workbook(file_path)
            ws = get_sheet(wb, sheet_name)
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            # Guardar y cerrar
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "start_cell": start_cell,
                "rows_written": len(data),
                "columns_written": max([len(row) if isinstance(row, list) else 1 for row in data], default=0),
                "message": f"Datos escritos correctamente desde {start_cell}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al escribir datos: {e}"
            }
    
    @mcp.tool(description="Actualiza individualmente una celda")
    def update_cell_tool(file_path, sheet_name, cell, value_or_formula):
        """
        Actualiza el valor o fórmula de una celda específica en una hoja de Excel
        
        Esta función permite modificar el contenido de una celda individual en una hoja de Excel.
        Puede usarse tanto para valores normales como para fórmulas.
        
        Args:
            file_path (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre de la hoja que contiene la celda a actualizar.
            cell (str): Referencia de la celda a actualizar (ej: "B5").
            value_or_formula (str/int/float/bool): Valor o fórmula a establecer. Las fórmulas deben comenzar con "=".
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo la celda modificada.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja especificada.
            CellReferenceError: Si la referencia de celda no es válida.
            
        Ejemplo:
            update_cell_tool("C:/datos/informe.xlsx", "Ventas", "C4", 5280.50)  # Valor numérico
            update_cell_tool("C:/datos/informe.xlsx", "Ventas", "D4", "=SUM(A1:A10)")  # Fórmula
        """
        try:
            # Abrir el archivo y obtener la hoja
            wb = openpyxl.load_workbook(file_path)
            ws = get_sheet(wb, sheet_name)
            
            # Actualizar la celda
            update_cell(ws, cell, value_or_formula)
            
            # Guardar y cerrar
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "cell": cell,
                "value": value_or_formula,
                "message": f"Celda {cell} actualizada correctamente en la hoja {sheet_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al actualizar celda: {e}"
            }
    
    # Registrar funciones avanzadas
    @mcp.tool(description="Define un rango como Tabla con estilo en una hoja de Excel")
    def add_table_tool(file_path, sheet_name, table_name, cell_range, style=None):
        """
        Define un rango como Tabla con estilo en Excel
        
        Esta función convierte un rango de celdas en una tabla Excel con formato,
        lo que permite filtrar, ordenar y dar formato automáticamente a los datos.
        
        Args:
            file_path (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre de la hoja donde se creará la tabla.
            table_name (str): Nombre para la tabla (debe ser único en el libro).
            cell_range (str): Rango de celdas para la tabla en formato Excel (ej: "A1:D10").
            style (str, optional): Estilo de tabla a aplicar (ej: "TableStyleMedium9"). Si es None,
                                 se utiliza el estilo predeterminado. Por defecto es None.
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo los detalles de la tabla creada.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja especificada.
            RangeError: Si el rango especificado no es válido.
            TableError: Si ya existe una tabla con el mismo nombre o hay otro problema con la tabla.
            
        Ejemplo:
            add_table_tool(
                "C:/datos/ventas.xlsx", 
                "Datos", 
                "TablaPreciosRegionales", 
                "B3:F15", 
                "TableStyleMedium2"
            )
        """
        try:
            # Abrir el archivo
            wb = openpyxl.load_workbook(file_path)
            
            # Obtener la hoja
            ws = get_sheet(wb, sheet_name)
            
            # Añadir la tabla
            table = add_table(ws, table_name, cell_range, style)
            
            # Guardar cambios
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "table_name": table_name,
                "range": cell_range,
                "style": style,
                "message": f"Tabla '{table_name}' creada correctamente en el rango {cell_range}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear tabla: {e}"
            }
    
    @mcp.tool(description="Inserta un gráfico nativo en una hoja de Excel con múltiples opciones de personalización")
    def add_chart_tool(file_path, sheet_name, chart_type, data_range, title=None, position=None, style=None, theme=None, custom_palette=None):
        """
        Inserta un gráfico profesional nativo en una hoja de Excel
        
        Esta función crea un gráfico basado en datos de la hoja de cálculo, con múltiples opciones
        de personalización para crear visualizaciones profesionales directamente en Excel.
        
        Args:
            file_path (str): Ruta completa y nombre del archivo Excel.
            sheet_name (str): Nombre de la hoja donde se insertará el gráfico.
            chart_type (str): Tipo de gráfico a crear. Opciones: 'line', 'bar', 'column', 'pie', 'scatter', 
                             'area', 'doughnut', 'radar', 'surface', 'stock'.
            data_range (str): Rango de celdas con los datos para el gráfico en formato Excel (ej: "A1:D10").
            title (str, optional): Título para el gráfico. Por defecto es None.
            position (str, optional): Posición donde insertar el gráfico en formato "A1:F15". 
                                     Si es None, se usa una posición por defecto. Por defecto es None.
            style (int, optional): Estilo numérico del gráfico (1-48). Por defecto es None.
            theme (str, optional): Tema de colores para el gráfico. Por defecto es None.
            custom_palette (list, optional): Lista de colores personalizados en formato hex (#RRGGBB). 
                                           Por defecto es None.
        
        Returns:
            dict: Información sobre el resultado de la operación, incluyendo detalles del gráfico creado.
            
        Raises:
            FileNotFoundError: Si el archivo especificado no existe.
            SheetNotFoundError: Si no existe la hoja especificada.
            RangeError: Si el rango de datos especificado no es válido.
            ChartError: Si hay un problema con la creación del gráfico.
            
        Ejemplo:
            add_chart_tool(
                "C:/datos/ventas.xlsx", 
                "Datos", 
                "column", 
                "A1:B10", 
                title="Ventas por Trimestre",
                position="E1:J15",
                style=12,
                custom_palette=["#4472C4", "#ED7D31", "#A5A5A5"]
            )
        """
        try:
            # Abrir el archivo
            wb = openpyxl.load_workbook(file_path)
            
            # Crear gráfico
            chart_id, chart = add_chart(wb, sheet_name, chart_type, data_range, title, position, style, theme, custom_palette)
            
            # Guardar cambios
            wb.save(file_path)
            
            # Extraer tipo de gráfico para mejor mensaje de respuesta
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
                "message": f"Gráfico '{chart_type_display}' creado correctamente con ID {chart_id}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear gráfico: {e}"
            }
    
    # Registrar nuevas funciones combinadas
    @mcp.tool(description="Crea una hoja con datos en un solo paso")
    def create_sheet_with_data_tool(file_path, sheet_name, data, overwrite=False):
        """
        Crea un archivo Excel con una hoja y datos en un solo paso.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            file_path (str): Ruta al archivo Excel a crear
            sheet_name (str): Nombre de la hoja a crear
            data (list): Array bidimensional con los datos
            overwrite (bool): Si es True, sobrescribe el archivo si existe
            
        Returns:
            dict: Resultado de la operación
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
            
            # Verificar si la hoja ya existe
            if sheet_name in wb.sheetnames:
                if overwrite:
                    # Eliminar la hoja existente
                    del wb[sheet_name]
                else:
                    raise SheetExistsError(f"La hoja '{sheet_name}' ya existe. Use overwrite=True para sobrescribir.")
            
            # Crear la hoja
            ws = wb.create_sheet(sheet_name)
            
            # Escribir los datos
            if data:
                write_sheet_data(ws, "A1", data)
            
            # Guardar el archivo
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "rows_written": len(data) if data else 0,
                "columns_written": max([len(row) if isinstance(row, list) else 1 for row in data], default=0) if data else 0,
                "message": f"Archivo creado con hoja '{sheet_name}' y datos"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear hoja con datos: {e}"
            }
    
    @mcp.tool(description="Crea una tabla formateada con datos en un solo paso")
    def create_formatted_table_tool(file_path, sheet_name, start_cell, data, table_name, table_style="TableStyleMedium9", formats=None):
        """
        Crea una tabla formateada con datos en un solo paso.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            file_path (str): Ruta al archivo Excel
            sheet_name (str): Nombre de la hoja
            start_cell (str): Celda inicial (ej. "A1")
            data (list): Array bidimensional con los datos
            table_name (str): Nombre para la tabla
            table_style (str): Estilo de la tabla
            formats (dict): Diccionario con formatos a aplicar:
                {
                    "A2:A10": "#,##0.00",  # Formato numérico
                    "B2:B10": {"bold": True, "fill_color": "FFFF00"}  # Estilo
                }
                
        Returns:
            dict: Resultado de la operación
        """
        try:
            # Verificar si el archivo existe, si no, crearlo
            if not os.path.exists(file_path):
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames and sheet_name != "Sheet":
                    # Renombrar la hoja predeterminada
                    wb["Sheet"].title = sheet_name
            else:
                wb = openpyxl.load_workbook(file_path)
                
                # Crear la hoja si no existe
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)
            
            # Obtener la hoja
            ws = wb[sheet_name]
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            # Determinar el rango de la tabla
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row = start_row + len(data) - 1
            end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
            table_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
            
            # Crear la tabla
            add_table(ws, table_name, table_range, table_style)
            
            # Aplicar formatos si se proporcionan
            if formats:
                for cell_range, fmt in formats.items():
                    if isinstance(fmt, dict):
                        # Es un estilo
                        apply_style(ws, cell_range, fmt)
                    else:
                        # Es un formato numérico
                        apply_number_format(ws, cell_range, fmt)
            
            # Guardar el archivo
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "table_name": table_name,
                "table_range": table_range,
                "table_style": table_style,
                "message": f"Tabla '{table_name}' creada y formateada correctamente"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear tabla formateada: {e}"
            }
    
    @mcp.tool(description="Crea un gráfico a partir de datos nuevos en un solo paso")
    def create_chart_from_data_tool(file_path, sheet_name, data, chart_type, position=None, title=None, style=None):
        """
        Crea un gráfico a partir de datos nuevos en un solo paso.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            file_path (str): Ruta al archivo Excel
            sheet_name (str): Nombre de la hoja
            data (list): Array bidimensional con los datos para el gráfico
            chart_type (str): Tipo de gráfico ('column', 'bar', 'line', 'pie', etc.)
            position (str): Celda donde colocar el gráfico (ej. "E1")
            title (str): Título del gráfico
            style: Estilo del gráfico
                
        Returns:
            dict: Resultado de la operación
        """
        try:
            # Verificar si el archivo existe, si no, crearlo
            if not os.path.exists(file_path):
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames and sheet_name != "Sheet":
                    # Renombrar la hoja predeterminada
                    wb["Sheet"].title = sheet_name
            else:
                wb = openpyxl.load_workbook(file_path)
                
                # Crear la hoja si no existe
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)
            
            # Obtener la hoja
            ws = wb[sheet_name]
            
            # Encontrar una zona libre para los datos
            # Buscar por la parte izquierda para colocar los datos de origen
            # (La convención común es poner datos a la izquierda y gráficos a la derecha)
            start_cell = "A1"
            
            # Comprobar si ya hay datos en esa zona
            if ws["A1"].value is not None:
                # Buscar la primera columna vacía
                col = 1
                while ws.cell(row=1, column=col).value is not None:
                    col += 1
                start_cell = f"{get_column_letter(col)}1"
            
            # Escribir los datos
            write_sheet_data(ws, start_cell, data)
            
            # Determinar el rango de datos para el gráfico
            start_row, start_col = ExcelRange.parse_cell_ref(start_cell)
            end_row = start_row + len(data) - 1
            end_col = start_col + (len(data[0]) if data and len(data) > 0 else 0) - 1
            data_range = ExcelRange.range_to_a1(start_row, start_col, end_row, end_col)
            
            # Determinar posición para el gráfico si no se proporciona
            if not position:
                # Colocar el gráfico a la derecha de los datos con un espacio
                chart_col = end_col + 2  # Dejar una columna de espacio
                position = f"{get_column_letter(chart_col + 1)}1"
            
            # Crear el gráfico
            chart_id, _ = add_chart(wb, sheet_name, chart_type, data_range, title, position, style)
            
            # Guardar el archivo
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "data_range": data_range,
                "chart_id": chart_id,
                "chart_type": chart_type,
                "position": position,
                "message": f"Gráfico '{chart_type}' creado correctamente a partir de nuevos datos"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al crear gráfico con datos: {e}"
            }
    
    
    @mcp.tool(description="Actualiza un informe existente con nuevos datos")
    def update_report_tool(file_path, data_updates, config_updates=None, recalculate=True):
        """
        Actualiza un informe existente con nuevos datos y configuraciones.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            file_path (str): Ruta al archivo Excel a actualizar
            data_updates (dict): Diccionario con actualizaciones de datos:
                {
                    "sheet_name": {
                        "range1": data_list1,
                        "range2": data_list2,
                        ...
                    }
                }
            config_updates (dict, opcional): Actualizaciones de configuración:
                {
                    "charts": [
                        {
                            "sheet": "sheet_name",
                            "id": 0,  # o "title"
                            "title": "New Title",
                            "style": "new_style"
                        }
                    ],
                    "tables": [
                        {
                            "sheet": "sheet_name",
                            "name": "TableName",
                            "range": "A1:D20"  # Nuevo rango
                        }
                    ]
                }
            recalculate (bool): Si es True, recalcula todas las fórmulas
                
        Returns:
            dict: Resultado de la operación
        """
        try:
            # Verificar que el archivo existe
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"El archivo no existe: {file_path}")
            
            # Abrir el archivo
            wb = openpyxl.load_workbook(file_path)
            
            # Actualizar datos
            for sheet_name, ranges in data_updates.items():
                if sheet_name not in wb.sheetnames:
                    logger.warning(f"La hoja '{sheet_name}' no existe, se omitirá")
                    continue
                
                ws = wb[sheet_name]
                
                for range_str, data in ranges.items():
                    # Si el rango es una sola celda, extraer la celda de inicio
                    if ':' not in range_str:
                        start_cell = range_str
                    else:
                        start_cell = range_str.split(':')[0]
                    
                    # Escribir los datos
                    write_sheet_data(ws, start_cell, data)
            
            # Actualizar configuraciones
            if config_updates:
                # Actualizar tablas
                for table_config in config_updates.get("tables", []):
                    sheet_name = table_config["sheet"]
                    table_name = table_config["name"]
                    
                    if sheet_name not in wb.sheetnames:
                        logger.warning(f"La hoja '{sheet_name}' no existe para actualizar la tabla '{table_name}'")
                        continue
                    
                    ws = wb[sheet_name]
                    
                    # Verificar si la tabla existe
                    if not hasattr(ws, 'tables') or table_name not in ws.tables:
                        logger.warning(f"La tabla '{table_name}' no existe en la hoja '{sheet_name}'")
                        continue
                    
                    # Actualizar rango de la tabla si se proporciona
                    if "range" in table_config:
                        refresh_table(ws, table_name, table_config["range"])
                
                # Actualizar gráficos
                for chart_config in config_updates.get("charts", []):
                    sheet_name = chart_config["sheet"]
                    chart_id = chart_config["id"]
                    
                    if sheet_name not in wb.sheetnames:
                        logger.warning(f"La hoja '{sheet_name}' no existe para actualizar el gráfico")
                        continue
                    
                    ws = wb[sheet_name]
                    
                    # Verificar si el chart_id es un índice o un título
                    if isinstance(chart_id, (int, str)) and str(chart_id).isdigit():
                        chart_idx = int(chart_id)
                    else:
                        # Buscar el gráfico por título
                        chart_idx = None
                        for i, chart_rel in enumerate(ws._charts):
                            chart = chart_rel[0]
                            if hasattr(chart, 'title') and chart.title == chart_id:
                                chart_idx = i
                                break
                    
                    if chart_idx is None or chart_idx >= len(ws._charts):
                        logger.warning(f"No se encontró el gráfico con ID/título '{chart_id}' en la hoja '{sheet_name}'")
                        continue
                    
                    # Actualizar propiedades del gráfico
                    chart = ws._charts[chart_idx][0]
                    
                    if "title" in chart_config:
                        chart.title = chart_config["title"]
                    
                    if "style" in chart_config:
                        try:
                            apply_chart_style(chart, chart_config["style"])
                        except Exception as style_error:
                            logger.warning(f"Error al aplicar estilo al gráfico: {style_error}")
            
            # Recalcular fórmulas si se solicita
            if recalculate:
                # openpyxl no tiene un método directo para recalcular
                # En Excel, esto se haría automáticamente al abrir el archivo
                # Aquí simplemente registramos que se solicitó recalcular
                logger.info("Se solicitó recalcular fórmulas (esto ocurrirá al abrir el archivo en Excel)")
            
            # Guardar el archivo
            wb.save(file_path)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheets_updated": list(data_updates.keys()),
                "message": f"Informe actualizado correctamente: {file_path}"
            }
        except Exception as e:
            logger.error(f"Error al actualizar informe: {e}")
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al actualizar informe: {e}"
            }
    
    @mcp.tool(description="Crea un dashboard dinámico con múltiples visualizaciones en un solo paso")
    def create_dashboard_tool(file_path, data, dashboard_config, overwrite=False):
        """
        Crea un dashboard dinámico con múltiples visualizaciones en un solo paso.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            file_path (str): Ruta al archivo Excel a crear
            data (dict): Diccionario con datos por hoja (ver documentación para formato)
            dashboard_config (dict): Configuración del dashboard (ver documentación para formato)
            overwrite (bool): Si es True, sobrescribe el archivo si existe
                
        Returns:
            dict: Resultado de la operación
        """
        return create_dynamic_dashboard(file_path, data, dashboard_config, overwrite)
    
    @mcp.tool(description="Crea un informe basado en una plantilla Excel, sustituyendo datos y actualizando gráficos")
    def create_report_from_template_tool(template_file, output_file, data_mappings, chart_mappings=None, format_mappings=None):
        """
        Crea un informe basado en una plantilla Excel, sustituyendo datos y actualizando gráficos.
        
        Args:
             **Nunca deben incluirse emojis en los textos escritos en celdas, etiquetas, títulos o gráficos de Excel.**

            template_file (str): Ruta a la plantilla Excel
            output_file (str): Ruta donde guardar el informe generado
            data_mappings (dict): Diccionario con mapeos de datos (ver documentación para formato)
            chart_mappings (dict, opcional): Diccionario con actualizaciones de gráficos
            format_mappings (dict, opcional): Diccionario con formatos a aplicar
                
        Returns:
            dict: Resultado de la operación
        """
        return create_report_from_template(template_file, output_file, data_mappings, chart_mappings, format_mappings)
    
    @mcp.tool(description="Importa datos desde múltiples fuentes (CSV, JSON, SQL) a un archivo Excel")
    def import_data_tool(excel_file, import_config, sheet_name=None, start_cell="A1", create_tables=False):
        """
        Importa datos desde múltiples fuentes (CSV, JSON, SQL) a un archivo Excel.
        
        Args:
            excel_file (str): Ruta al archivo Excel donde importar los datos
            import_config (dict): Configuración de importación (ver documentación para formato)
            sheet_name (str, opcional): Nombre de hoja predeterminado
            start_cell (str, opcional): Celda inicial predeterminada
            create_tables (bool, opcional): Si es True, crea tablas Excel
                
        Returns:
            dict: Resultado de la operación
        """
        return import_multi_source_data(excel_file, import_config, sheet_name, start_cell, create_tables)
    
    @mcp.tool(description="Exporta datos de Excel a múltiples formatos (CSV, JSON, PDF)")
    def export_data_tool(excel_file, export_config):
        """
        Exporta datos de Excel a múltiples formatos (CSV, JSON, PDF).
        
        Args:
            excel_file (str): Ruta al archivo Excel de origen
            export_config (dict): Configuración de exportación (ver documentación para formato)
                
        Returns:
            dict: Resultado de la operación
        """
        return export_excel_data(excel_file, export_config)
    
    @mcp.tool(description="Filtra y extrae datos de una tabla o rango en formato de registros")
    def filter_data_tool(file_path, sheet_name, range_str=None, table_name=None, filters=None):
        """
        Filtra y extrae datos de una tabla o rango en formato de registros.
        
        Args:
            file_path (str): Ruta al archivo Excel
            sheet_name (str): Nombre de la hoja
            range_str (str, opcional): Rango en formato A1:B5 (requerido si no se especifica table_name)
            table_name (str, opcional): Nombre de la tabla (requerido si no se especifica range_str)
            filters (dict, opcional): Filtros a aplicar a los datos:
                {
                    "field1": value1,  # Igualdad simple
                    "field2": [value1, value2],  # Lista de valores posibles
                    "field3": {"gt": 100},  # Mayor que
                    "field4": {"lt": 50},  # Menor que
                    "field5": {"contains": "text"}  # Contiene texto
                }
                
        Returns:
            dict: Resultado de la operación con los datos filtrados
        """
        try:
            # Validar argumentos
            if not range_str and not table_name:
                raise ValueError("Debe proporcionar 'range_str' o 'table_name'")
            
            # Abrir el archivo
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            # Verificar que la hoja existe
            if sheet_name not in wb.sheetnames:
                raise SheetNotFoundError(f"La hoja '{sheet_name}' no existe en el archivo")
            
            ws = wb[sheet_name]
            
            # Si se proporciona table_name, obtener su rango
            if table_name:
                if not hasattr(ws, 'tables') or table_name not in ws.tables:
                    raise TableNotFoundError(f"La tabla '{table_name}' no existe en la hoja '{sheet_name}'")
                
                range_str = ws.tables[table_name].ref
            
            # Filtrar los datos
            filtered_data = filter_sheet_data(wb, sheet_name, range_str, filters)
            
            return {
                "success": True,
                "file_path": file_path,
                "sheet_name": sheet_name,
                "source": f"Tabla '{table_name}'" if table_name else f"Rango {range_str}",
                "filtered_data": filtered_data,
                "record_count": len(filtered_data),
                "message": f"Se encontraron {len(filtered_data)} registros que cumplen los criterios"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "message": f"Error al filtrar datos: {e}"
            }

    @mcp.tool(description="Exporta un libro a PDF solo si tiene una única hoja visible")
    def export_single_sheet_pdf_tool(excel_file, output_pdf=None):
        """Exporta un archivo Excel a PDF si solo tiene una hoja visible."""
        return export_single_visible_sheet_pdf(excel_file, output_pdf)

    @mcp.tool(description="Exporta una o varias hojas a PDF")
    def export_sheets_pdf_tool(excel_file, sheets=None, output_dir=None, single_file=False):
        """Exporta las hojas indicadas de un libro Excel a PDF.

        ``sheets`` puede ser un nombre de hoja o una lista. Si es ``None`` se
        exportará cada hoja existente de forma individual. Si ``single_file`` es
        ``True`` y se especifican varias hojas, se intentará generar un único
        PDF con todas ellas.
        """
        return export_sheets_to_pdf(excel_file, sheets, output_dir, single_file)

if __name__ == "__main__":
    logger.info("Master Excel MCP - Ejemplo de uso")
    logger.info("Este módulo unifica todas las funcionalidades Excel en un solo lugar.")
