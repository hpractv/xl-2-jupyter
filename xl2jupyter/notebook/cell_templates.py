"""Templates for generating Jupyter notebook cell content."""

from typing import Any

import pandas as pd

from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.workbook import WorkbookModel


def create_markdown_intro(workbook: WorkbookModel) -> str:
    """
    Create markdown introduction cell content.

    Args:
        workbook: WorkbookModel instance

    Returns:
        Markdown string
    """
    lines = [
        "# Excel Workbook Analysis",
        "",
        f"**Source File:** `{workbook.file_path or 'Unknown'}`",
        "",
        f"**Sheets:** {len(workbook.sheets)}",
        "",
        f"**VBA Modules:** {len(workbook.vba_modules)}",
        "",
        "---",
        "",
        "This notebook contains:",
        "- DataFrames for each sheet",
        "- Extracted formulas",
        "- Formula dependency information",
        "- VBA module code",
    ]

    return "\n".join(lines)


def create_imports_cell() -> str:
    """
    Create Python imports cell content.

    Returns:
        Python code string
    """
    return """import pandas as pd
import numpy as np
from datetime import datetime, date"""


def create_dataframe_cell(sheet: SheetModel) -> str:
    """
    Create DataFrame cell content for a sheet.

    Args:
        sheet: SheetModel instance

    Returns:
        Python code string that creates a pandas DataFrame
    """
    data = sheet.get_data_array()

    if not data:
        return f"""# Sheet: {sheet.name}
# No data in this sheet
df_{_sanitize_name(sheet.name)} = pd.DataFrame()"""

    # Convert to DataFrame for proper formatting
    df = pd.DataFrame(data)

    # Generate code to recreate DataFrame
    lines = [f"# Sheet: {sheet.name}", f"df_{_sanitize_name(sheet.name)} = pd.DataFrame("]

    # Format data as list of lists
    data_str = "[\n"
    for row in data:
        row_str = "    [" + ", ".join(_format_value(val) for val in row) + "],\n"
        data_str += row_str
    data_str += "]"

    lines.append(data_str)
    lines.append(")")

    # Add display
    lines.append(f"\ndf_{_sanitize_name(sheet.name)}")

    return "\n".join(lines)


def create_formulas_cell(workbook: WorkbookModel) -> str:
    """
    Create formulas listing cell content.

    Args:
        workbook: WorkbookModel instance

    Returns:
        Python code string with formulas
    """
    lines = ["# Extracted Formulas", ""]

    has_formulas = False
    for sheet_name, sheet in workbook.sheets.items():
        sheet_formulas = []
        for cell in sheet.cells.values():
            if cell.formula:
                formula_str = f"={cell.formula}" if not cell.formula.startswith("=") else cell.formula
                sheet_formulas.append((cell.address, formula_str, cell.dependencies))

        if sheet_formulas:
            has_formulas = True
            lines.append(f"## Sheet: {sheet_name}")
            lines.append("")
            for addr, formula, deps in sorted(sheet_formulas):
                lines.append(f"### {sheet_name}!{addr}")
                lines.append(f"```excel")
                lines.append(formula)
                lines.append("```")
                if deps:
                    lines.append(f"**Dependencies:** {', '.join(deps)}")
                lines.append("")

    if not has_formulas:
        lines.append("No formulas found in this workbook.")

    return "\n".join(lines)


def create_dependency_graph_cell(workbook: WorkbookModel, graph) -> str:
    """
    Create dependency graph visualization cell content.

    Args:
        workbook: WorkbookModel instance
        graph: NetworkX graph (optional)

    Returns:
        Python code string with dependency graph information
    """
    lines = ["# Formula Dependency Graph", ""]

    # Count dependencies
    total_dependencies = 0
    for sheet in workbook.sheets.values():
        for cell in sheet.cells.values():
            if cell.dependencies:
                total_dependencies += len(cell.dependencies)

    lines.append(f"**Total formula dependencies:** {total_dependencies}")
    lines.append("")

    # List dependencies by sheet
    for sheet_name, sheet in workbook.sheets.items():
        sheet_deps = []
        for cell in sheet.cells.values():
            if cell.dependencies:
                for dep in cell.dependencies:
                    sheet_deps.append((f"{sheet_name}!{cell.address}", dep))

        if sheet_deps:
            lines.append(f"## Sheet: {sheet_name}")
            lines.append("")
            for source, target in sorted(sheet_deps):
                lines.append(f"- `{source}` depends on `{target}`")
            lines.append("")

    if total_dependencies == 0:
        lines.append("No dependencies found.")

    return "\n".join(lines)


def create_vba_cell(workbook: WorkbookModel) -> str:
    """
    Create VBA modules cell content.

    Args:
        workbook: WorkbookModel instance

    Returns:
        Python code string with VBA modules
    """
    lines = ["# VBA Modules", ""]

    if not workbook.vba_modules:
        lines.append("No VBA modules found in this workbook.")
        return "\n".join(lines)

    for module in workbook.vba_modules:
        lines.append(f"## {module.name} ({module.module_type})")
        if module.sheet_name:
            lines.append(f"**Sheet:** {module.sheet_name}")
        lines.append("")
        lines.append("```vba")
        lines.append(module.code)
        lines.append("```")
        lines.append("")

    return "\n".join(lines)


def _sanitize_name(name: str) -> str:
    """
    Sanitize sheet name for use in Python variable names.

    Args:
        name: Sheet name

    Returns:
        Sanitized name
    """
    # Replace spaces and special characters with underscores
    sanitized = "".join(c if c.isalnum() or c == "_" else "_" for c in name)
    # Ensure it starts with a letter or underscore
    if sanitized and not (sanitized[0].isalpha() or sanitized[0] == "_"):
        sanitized = "_" + sanitized
    if not sanitized:
        sanitized = "sheet"
    return sanitized


def _format_value(value: Any) -> str:
    """
    Format a value for Python code generation.

    Args:
        value: Value to format

    Returns:
        Formatted string
    """
    if value is None:
        return "None"
    elif isinstance(value, bool):
        return str(value)
    elif isinstance(value, (int, float)):
        return str(value)
    elif isinstance(value, str):
        # Escape quotes and newlines
        escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
        return f'"{escaped}"'
    else:
        return repr(value)
