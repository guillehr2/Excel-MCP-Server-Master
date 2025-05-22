#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Setup script for Excel MCP Server."""

from setuptools import setup, find_packages
import os

# Read the README file
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

# Read requirements
with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="excel-mcp-server",
    version="1.0.3",
    author="Guillem Hermida",
    author_email="qtmsuite@gmail.com",
    description="A comprehensive Model Context Protocol (MCP) server for Excel file manipulation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/guillehr2/Excel-MCP-Server-Master",
    project_urls={
        "Bug Tracker": "https://github.com/guillehr2/Excel-MCP-Server-Master/issues",
        "Documentation": "https://github.com/guillehr2/Excel-MCP-Server-Master#readme",
        "Source Code": "https://github.com/guillehr2/Excel-MCP-Server-Master",
    },
    py_modules=["master_excel_mcp"],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "excel-mcp-server=master_excel_mcp:main",
        ],
    },
    keywords="mcp excel openpyxl automation ai llm spreadsheet",
    include_package_data=True,
    zip_safe=False,
)
