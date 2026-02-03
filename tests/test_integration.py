"""Integration tests for the full pipeline."""

import pytest
from pathlib import Path

from xl2jupyter.extraction.excel_reader import ExcelReader
from xl2jupyter.notebook.notebook_builder import NotebookBuilder
from xl2jupyter.parsing.formula_parser import FormulaParser


def test_full_pipeline(tmp_path, fixtures_dir):
    """
    Test the full pipeline from Excel file to Jupyter notebook.

    This test requires:
    1. Excel to be installed
    2. Test workbook to be created (run create_test_workbook.py)
    """
    test_workbook = fixtures_dir / "comprehensive_test.xlsb"

    if not test_workbook.exists():
        pytest.skip("Test workbook not found. Run create_test_workbook.py first.")

    try:
        # Step 1: Extract workbook
        reader = ExcelReader(test_workbook)
        workbook = reader.read()

        assert workbook is not None
        assert len(workbook.sheets) > 0

        # Step 2: Parse formulas
        parser = FormulaParser()
        for sheet_name, sheet in workbook.sheets.items():
            for cell in sheet.cells.values():
                if cell.formula:
                    parser.parse_cell(cell, sheet_name)

        # Step 3: Build notebook
        builder = NotebookBuilder(workbook)
        notebook = builder.build()

        assert notebook is not None
        assert len(notebook.cells) > 0

        # Step 4: Save notebook
        output_path = tmp_path / "test_output.ipynb"
        builder.save(output_path)

        assert output_path.exists()

    except Exception as e:
        # If Excel is not available, skip the test
        if "xlwings" in str(e).lower() or "excel" in str(e).lower():
            pytest.skip(f"Excel/xlwings not available: {e}")
        else:
            raise


def test_pipeline_with_sample_workbook(tmp_path):
    """Test pipeline with programmatically created workbook model."""
    from xl2jupyter.model.cell import CellModel
    from xl2jupyter.model.sheet import SheetModel
    from xl2jupyter.model.workbook import WorkbookModel

    # Create workbook model
    workbook = WorkbookModel(file_path="test.xlsb")
    sheet = SheetModel(name="Sheet1")
    sheet.set_cell(CellModel(row=1, column=1, value=10))
    sheet.set_cell(CellModel(row=1, column=2, value=20, formula="A1*2"))
    workbook.add_sheet(sheet)

    # Parse formulas
    parser = FormulaParser()
    for sheet_name, sheet in workbook.sheets.items():
        for cell in sheet.cells.values():
            if cell.formula:
                parser.parse_cell(cell, sheet_name)

    # Build notebook
    builder = NotebookBuilder(workbook)
    notebook = builder.build()

    assert notebook is not None
    assert len(notebook.cells) > 0

    # Save notebook
    output_path = tmp_path / "test_output.ipynb"
    builder.save(output_path)

    assert output_path.exists()
