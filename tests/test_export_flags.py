"""Tests for export flags functionality."""

import json
from pathlib import Path

import nbformat
import pytest

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule
from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.notebook.notebook_builder import NotebookBuilder
from xl2jupyter.notebook.cell_templates import create_markdown_intro


class TestExportFlags:
    """Tests for selective export functionality."""

    def test_export_all_default(self):
        """Test that all content is exported by default."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20, formula="A1*2"))
        workbook.add_sheet(sheet)

        vba_module = VBAModule(
            name="Module1",
            code="Sub Test()\nEnd Sub",
            module_type="Standard",
        )
        workbook.add_vba_module(vba_module)

        builder = NotebookBuilder(workbook)
        notebook = builder.build()

        # Should have intro, imports, data, formulas, dependency graph, and VBA
        assert len(notebook.cells) >= 5

        # Check for code cells (imports + data)
        code_cells = [c for c in notebook.cells if c.cell_type == "code"]
        assert len(code_cells) >= 2

    def test_export_data_only(self):
        """Test exporting only data."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20, formula="A1*2"))
        workbook.add_sheet(sheet)

        vba_module = VBAModule(
            name="Module1",
            code="Sub Test()\nEnd Sub",
            module_type="Standard",
        )
        workbook.add_vba_module(vba_module)

        builder = NotebookBuilder(
            workbook,
            export_data=True,
            export_formulas=False,
            export_graphs=False,
            export_vba=False,
        )
        notebook = builder.build()

        # Should have intro, imports, and data cells only
        assert len(notebook.cells) >= 2

        # Check that notebook has code cells (imports + data)
        code_cells = [c for c in notebook.cells if c.cell_type == "code"]
        assert len(code_cells) >= 2

        # Check that formulas and VBA are not in the notebook
        notebook_text = json.dumps(notebook, indent=2)
        assert "Formula" not in notebook_text or "Extracted Formulas" not in notebook_text
        assert "Module1" not in notebook_text

    def test_export_formulas_only(self):
        """Test exporting only formulas."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20, formula="A1*2"))
        workbook.add_sheet(sheet)

        builder = NotebookBuilder(
            workbook,
            export_data=False,
            export_formulas=True,
            export_graphs=False,
            export_vba=False,
        )
        notebook = builder.build()

        # Should have intro and formulas cells
        assert len(notebook.cells) >= 2

        # Check that we have markdown cells with formulas
        markdown_cells = [c for c in notebook.cells if c.cell_type == "markdown"]
        assert len(markdown_cells) >= 2

        # Check that formulas are in the notebook
        notebook_text = json.dumps(notebook, indent=2)
        assert "Extracted Formulas" in notebook_text or "Formula" in notebook_text

        # Check that data cells (imports) are not included
        code_cells = [c for c in notebook.cells if c.cell_type == "code"]
        assert len(code_cells) == 0

    def test_export_vba_only(self):
        """Test exporting only VBA."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet)

        vba_module = VBAModule(
            name="Module1",
            code="Sub Test()\nEnd Sub",
            module_type="Standard",
        )
        workbook.add_vba_module(vba_module)

        builder = NotebookBuilder(
            workbook,
            export_data=False,
            export_formulas=False,
            export_graphs=False,
            export_vba=True,
        )
        notebook = builder.build()

        # Should have intro and VBA cells
        assert len(notebook.cells) >= 2

        # Check that VBA is in the notebook
        notebook_text = json.dumps(notebook, indent=2)
        assert "Module1" in notebook_text or "VBA" in notebook_text

    def test_export_multiple_items(self):
        """Test exporting multiple items."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20, formula="A1*2"))
        workbook.add_sheet(sheet)

        vba_module = VBAModule(
            name="Module1",
            code="Sub Test()\nEnd Sub",
            module_type="Standard",
        )
        workbook.add_vba_module(vba_module)

        builder = NotebookBuilder(
            workbook,
            export_data=True,
            export_formulas=True,
            export_graphs=False,
            export_vba=False,
        )
        notebook = builder.build()

        # Should have intro, imports, data, formulas, and dependency graph
        assert len(notebook.cells) >= 4

        # Check for code cells (imports + data)
        code_cells = [c for c in notebook.cells if c.cell_type == "code"]
        assert len(code_cells) >= 2

        # Check that formulas are in the notebook
        notebook_text = json.dumps(notebook, indent=2)
        assert "Extracted Formulas" in notebook_text or "Formula" in notebook_text

        # Check that VBA is not in the notebook
        assert "Module1" not in notebook_text

    def test_export_nothing(self):
        """Test exporting nothing (edge case)."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet)

        builder = NotebookBuilder(
            workbook,
            export_data=False,
            export_formulas=False,
            export_graphs=False,
            export_vba=False,
        )
        notebook = builder.build()

        # Should have at least intro cell
        assert len(notebook.cells) >= 1
        assert notebook.cells[0].cell_type == "markdown"

    def test_markdown_intro_reflects_export_settings(self):
        """Test that markdown intro shows what's being exported."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        workbook.add_sheet(sheet)

        # Test data only
        intro = create_markdown_intro(
            workbook,
            export_data=True,
            export_formulas=False,
            export_graphs=False,
            export_vba=False,
        )
        assert "DataFrames" in intro
        assert "Extracted formulas" not in intro
        assert "VBA module code" not in intro

        # Test formulas only
        intro = create_markdown_intro(
            workbook,
            export_data=False,
            export_formulas=True,
            export_graphs=False,
            export_vba=False,
        )
        assert "Extracted formulas" in intro
        assert "DataFrames" not in intro
        assert "VBA module code" not in intro

        # Test VBA only
        intro = create_markdown_intro(
            workbook,
            export_data=False,
            export_formulas=False,
            export_graphs=False,
            export_vba=True,
        )
        assert "VBA module code" in intro
        assert "DataFrames" not in intro

        # Test all
        intro = create_markdown_intro(
            workbook,
            export_data=True,
            export_formulas=True,
            export_graphs=True,
            export_vba=True,
        )
        assert "DataFrames" in intro
        assert "Extracted formulas" in intro
        assert "VBA module code" in intro
        assert "Charts" in intro or "graphs" in intro
