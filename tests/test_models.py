"""Tests for data models."""

import json
from pathlib import Path

import pytest

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule
from xl2jupyter.model.workbook import WorkbookModel


class TestCellModel:
    """Tests for CellModel."""

    def test_cell_creation(self):
        """Test creating a cell."""
        cell = CellModel(row=1, column=1, value=10)
        assert cell.row == 1
        assert cell.column == 1
        assert cell.value == 10

    def test_cell_address(self):
        """Test cell address property."""
        cell = CellModel(row=1, column=1)
        assert cell.address == "A1"

        cell = CellModel(row=1, column=26)
        assert cell.address == "Z1"

        cell = CellModel(row=1, column=27)
        assert cell.address == "AA1"

    def test_cell_serialization(self):
        """Test cell serialization to/from dict."""
        cell = CellModel(row=1, column=1, value=10, formula="SUM(A1:A5)")
        data = cell.to_dict()
        assert data["row"] == 1
        assert data["column"] == 1
        assert data["value"] == 10
        assert data["formula"] == "SUM(A1:A5)"

        cell2 = CellModel.from_dict(data)
        assert cell2.row == cell.row
        assert cell2.column == cell.column
        assert cell2.value == cell.value
        assert cell2.formula == cell.formula


class TestSheetModel:
    """Tests for SheetModel."""

    def test_sheet_creation(self):
        """Test creating a sheet."""
        sheet = SheetModel(name="TestSheet")
        assert sheet.name == "TestSheet"
        assert len(sheet.cells) == 0

    def test_add_cell(self):
        """Test adding cells to sheet."""
        sheet = SheetModel(name="TestSheet")
        cell = CellModel(row=1, column=1, value=10)
        sheet.set_cell(cell)
        assert len(sheet.cells) == 1
        assert sheet.get_cell(1, 1) == cell

    def test_used_range(self):
        """Test used range tracking."""
        sheet = SheetModel(name="TestSheet")
        assert sheet.used_range is None

        cell1 = CellModel(row=1, column=1, value=10)
        sheet.set_cell(cell1)
        assert sheet.used_range == (1, 1, 1, 1)

        cell2 = CellModel(row=5, column=3, value=20)
        sheet.set_cell(cell2)
        assert sheet.used_range == (1, 5, 1, 3)

    def test_get_data_array(self):
        """Test getting sheet data as array."""
        sheet = SheetModel(name="TestSheet")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        sheet.set_cell(CellModel(row=1, column=2, value=20))
        sheet.set_cell(CellModel(row=2, column=1, value=30))

        data = sheet.get_data_array()
        assert len(data) == 2
        assert len(data[0]) == 2
        assert data[0][0] == 10
        assert data[0][1] == 20
        assert data[1][0] == 30

    def test_sheet_serialization(self):
        """Test sheet serialization."""
        sheet = SheetModel(name="TestSheet")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        data = sheet.to_dict()
        assert data["name"] == "TestSheet"
        assert len(data["cells"]) == 1

        sheet2 = SheetModel.from_dict(data)
        assert sheet2.name == sheet.name
        assert len(sheet2.cells) == len(sheet.cells)


class TestWorkbookModel:
    """Tests for WorkbookModel."""

    def test_workbook_creation(self):
        """Test creating a workbook."""
        workbook = WorkbookModel()
        assert len(workbook.sheets) == 0
        assert len(workbook.vba_modules) == 0

    def test_add_sheet(self):
        """Test adding sheets to workbook."""
        workbook = WorkbookModel()
        sheet = SheetModel(name="Sheet1")
        workbook.add_sheet(sheet)
        assert len(workbook.sheets) == 1
        assert workbook.get_sheet("Sheet1") == sheet

    def test_add_vba_module(self):
        """Test adding VBA modules."""
        workbook = WorkbookModel()
        module = VBAModule(name="Module1", code="Sub Test()\nEnd Sub", module_type="Standard")
        workbook.add_vba_module(module)
        assert len(workbook.vba_modules) == 1

    def test_workbook_serialization(self):
        """Test workbook serialization."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet)

        data = workbook.to_dict()
        assert data["file_path"] == "test.xlsb"
        assert len(data["sheets"]) == 1

        workbook2 = WorkbookModel.from_dict(data)
        assert workbook2.file_path == workbook.file_path
        assert len(workbook2.sheets) == len(workbook.sheets)

    def test_workbook_json_serialization(self, tmp_path):
        """Test workbook JSON serialization."""
        workbook = WorkbookModel(file_path="test.xlsb")
        sheet = SheetModel(name="Sheet1")
        sheet.set_cell(CellModel(row=1, column=1, value=10))
        workbook.add_sheet(sheet)

        json_file = tmp_path / "test.json"
        workbook.to_json(json_file)

        assert json_file.exists()
        workbook2 = WorkbookModel.from_json(json_file)
        assert workbook2.file_path == workbook.file_path
        assert len(workbook2.sheets) == len(workbook.sheets)


class TestVBAModule:
    """Tests for VBAModule."""

    def test_vba_module_creation(self):
        """Test creating a VBA module."""
        module = VBAModule(
            name="Module1",
            code='Sub Test()\n    MsgBox "Hello"\nEnd Sub',
            module_type="Standard",
        )
        assert module.name == "Module1"
        assert "Sub Test()" in module.code
        assert module.module_type == "Standard"

    def test_vba_module_serialization(self):
        """Test VBA module serialization."""
        module = VBAModule(
            name="Module1",
            code="Sub Test()\nEnd Sub",
            module_type="Standard",
            sheet_name="Sheet1",
        )
        data = module.to_dict()
        assert data["name"] == "Module1"
        assert data["module_type"] == "Standard"
        assert data["sheet_name"] == "Sheet1"

        module2 = VBAModule.from_dict(data)
        assert module2.name == module.name
        assert module2.code == module.code
