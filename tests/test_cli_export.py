"""Test the CLI export flags functionality."""

import json
import sys

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule
from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.notebook.notebook_builder import NotebookBuilder
from xl2jupyter.parsing.formula_parser import FormulaParser


def create_test_workbook():
    """Create a test workbook with data, formulas, and VBA."""
    workbook = WorkbookModel(file_path="test_cli.xlsb")

    # Add Sheet1 with data and formulas
    sheet1 = SheetModel(name="Sheet1")
    sheet1.set_cell(CellModel(row=1, column=1, value="Name", data_type="text"))
    sheet1.set_cell(CellModel(row=1, column=2, value="Value", data_type="text"))
    sheet1.set_cell(CellModel(row=2, column=1, value="Item1", data_type="text"))
    sheet1.set_cell(CellModel(row=2, column=2, value=10, data_type="number"))
    sheet1.set_cell(CellModel(row=3, column=1, value="Item2", data_type="text"))
    sheet1.set_cell(CellModel(row=3, column=2, value=20, data_type="number"))
    sheet1.set_cell(CellModel(row=4, column=1, value="Total", data_type="text"))
    sheet1.set_cell(CellModel(row=4, column=2, value=30, formula="B2+B3", data_type="number"))
    workbook.add_sheet(sheet1)

    # Add VBA module
    vba_module = VBAModule(
        name="Module1",
        code='Sub HelloWorld()\n    MsgBox "Hello, World!"\nEnd Sub',
        module_type="Standard",
    )
    workbook.add_vba_module(vba_module)

    return workbook


def test_export_all():
    """Test exporting all content (default)."""
    print("\n=== Testing --export all ===")
    workbook = create_test_workbook()

    # Parse formulas
    parser = FormulaParser()
    for sheet_name, sheet in workbook.sheets.items():
        for cell in sheet.cells.values():
            if cell.formula:
                parser.parse_cell(cell, sheet_name)

    # Build notebook with all exports
    builder = NotebookBuilder(
        workbook,
        export_data=True,
        export_formulas=True,
        export_graphs=True,
        export_vba=True,
    )
    notebook = builder.build()

    print(f"✓ Created notebook with {len(notebook.cells)} cells")

    # Check content
    notebook_text = json.dumps(notebook, indent=2)
    assert "DataFrames" in notebook_text or "df_Sheet1" in notebook_text
    assert "Extracted Formulas" in notebook_text or "Formula" in notebook_text
    assert "VBA" in notebook_text and "Module1" in notebook_text
    print("✓ Notebook contains data, formulas, and VBA")


def test_export_data_only():
    """Test exporting only data."""
    print("\n=== Testing --export data ===")
    workbook = create_test_workbook()

    builder = NotebookBuilder(
        workbook,
        export_data=True,
        export_formulas=False,
        export_graphs=False,
        export_vba=False,
    )
    notebook = builder.build()

    print(f"✓ Created notebook with {len(notebook.cells)} cells")

    # Check content
    notebook_text = json.dumps(notebook, indent=2)
    assert "df_Sheet1" in notebook_text or "DataFrames" in notebook_text
    assert "Extracted Formulas" not in notebook_text
    assert "Module1" not in notebook_text
    print("✓ Notebook contains only data (no formulas or VBA)")


def test_export_formulas_only():
    """Test exporting only formulas."""
    print("\n=== Testing --export formulas ===")
    workbook = create_test_workbook()

    # Parse formulas
    parser = FormulaParser()
    for sheet_name, sheet in workbook.sheets.items():
        for cell in sheet.cells.values():
            if cell.formula:
                parser.parse_cell(cell, sheet_name)

    builder = NotebookBuilder(
        workbook,
        export_data=False,
        export_formulas=True,
        export_graphs=False,
        export_vba=False,
    )
    notebook = builder.build()

    print(f"✓ Created notebook with {len(notebook.cells)} cells")

    # Check content
    notebook_text = json.dumps(notebook, indent=2)
    assert "Extracted Formulas" in notebook_text or "Formula" in notebook_text
    assert "df_Sheet1" not in notebook_text
    assert "Module1" not in notebook_text
    print("✓ Notebook contains only formulas (no data or VBA)")


def test_export_vba_only():
    """Test exporting only VBA."""
    print("\n=== Testing --export vba ===")
    workbook = create_test_workbook()

    builder = NotebookBuilder(
        workbook,
        export_data=False,
        export_formulas=False,
        export_graphs=False,
        export_vba=True,
    )
    notebook = builder.build()

    print(f"✓ Created notebook with {len(notebook.cells)} cells")

    # Check content
    notebook_text = json.dumps(notebook, indent=2)
    assert "VBA" in notebook_text and "Module1" in notebook_text
    assert "df_Sheet1" not in notebook_text
    assert "Extracted Formulas" not in notebook_text
    print("✓ Notebook contains only VBA (no data or formulas)")


def test_export_multiple():
    """Test exporting multiple items."""
    print("\n=== Testing --export data formulas ===")
    workbook = create_test_workbook()

    # Parse formulas
    parser = FormulaParser()
    for sheet_name, sheet in workbook.sheets.items():
        for cell in sheet.cells.values():
            if cell.formula:
                parser.parse_cell(cell, sheet_name)

    builder = NotebookBuilder(
        workbook,
        export_data=True,
        export_formulas=True,
        export_graphs=False,
        export_vba=False,
    )
    notebook = builder.build()

    print(f"✓ Created notebook with {len(notebook.cells)} cells")

    # Check content
    notebook_text = json.dumps(notebook, indent=2)
    assert "df_Sheet1" in notebook_text or "DataFrames" in notebook_text
    assert "Extracted Formulas" in notebook_text or "Formula" in notebook_text
    assert "Module1" not in notebook_text
    print("✓ Notebook contains data and formulas (no VBA)")


if __name__ == "__main__":
    print("Testing export flags CLI functionality...\n")

    try:
        test_export_all()
        test_export_data_only()
        test_export_formulas_only()
        test_export_vba_only()
        test_export_multiple()

        print("\n" + "=" * 50)
        print("✓ All CLI export flag tests passed!")
        print("=" * 50)
    except AssertionError as e:
        print(f"\n✗ Test failed: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
