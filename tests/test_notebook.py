"""Tests for notebook generation."""

import json
from pathlib import Path

import nbformat
import pytest

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule
from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.notebook.notebook_builder import NotebookBuilder


class TestNotebookBuilder:
    """Tests for NotebookBuilder."""

    def test_build_notebook(self):
        """Test building a notebook from workbook."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20))
        workbook.add_sheet(sheet)

        builder = NotebookBuilder(workbook)
        notebook = builder.build()

        assert notebook is not None
        assert len(notebook.cells) > 0

        # Check for markdown intro cell
        assert notebook.cells[0].cell_type == "markdown"

        # Check for imports cell
        assert notebook.cells[1].cell_type == "code"

    def test_notebook_has_dataframe_cells(self):
        """Test that notebook has DataFrame cells for each sheet."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet1 = SheetModel(name="Sheet1")
        sheet1.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet1)

        sheet2 = SheetModel(name="Sheet2")
        sheet2.set_cell(CellModel(row=1, column=1, value=20))
        workbook.add_sheet(sheet2)

        builder = NotebookBuilder(workbook)
        notebook = builder.build()

        # Count code cells (should have imports + 2 DataFrame cells)
        code_cells = [c for c in notebook.cells if c.cell_type == "code"]
        assert len(code_cells) >= 3  # imports + 2 sheets

    def test_notebook_includes_vba(self):
        """Test that notebook includes VBA modules."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        workbook.add_sheet(sheet)

        vba_module = VBAModule(
            name="Module1",
            code='Sub Test()\n    MsgBox "Hello"\nEnd Sub',
            module_type="Standard",
        )
        workbook.add_vba_module(vba_module)

        builder = NotebookBuilder(workbook)
        notebook = builder.build()

        # Check that VBA content is in notebook
        notebook_text = json.dumps(notebook, indent=2)
        assert "Module1" in notebook_text or "VBA" in notebook_text

    def test_save_notebook(self, tmp_path):
        """Test saving notebook to file."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet)

        builder = NotebookBuilder(workbook)
        builder.build()

        output_path = tmp_path / "test.ipynb"
        builder.save(output_path)

        assert output_path.exists()

        # Verify it's valid JSON
        with open(output_path) as f:
            data = json.load(f)
            assert "cells" in data
            assert "metadata" in data
