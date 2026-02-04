"""Pytest configuration and fixtures."""

import pytest
from pathlib import Path

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule
from xl2jupyter.model.workbook import WorkbookModel


@pytest.fixture
def sample_cell():
    """Create a sample cell for testing."""
    return CellModel(row=1, column=1, value=10, formula="SUM(A1:A5)")


@pytest.fixture
def sample_sheet():
    """Create a sample sheet for testing."""
    sheet = SheetModel(name="TestSheet")
    for row in range(1, 4):
        for col in range(1, 4):
            cell = CellModel(row=row, column=col, value=row * col)
            sheet.set_cell(cell)
    return sheet


@pytest.fixture
def sample_workbook():
    """Create a sample workbook for testing."""
    workbook = WorkbookModel(file_path="test.xlsb")
    sheet = SheetModel(name="Sheet1")
    cell1 = CellModel(row=1, column=1, value=10)
    cell2 = CellModel(row=1, column=2, value=20, formula="A1*2")
    sheet.set_cell(cell1)
    sheet.set_cell(cell2)
    workbook.add_sheet(sheet)
    return workbook


@pytest.fixture
def sample_vba_module():
    """Create a sample VBA module for testing."""
    return VBAModule(
        name="Module1",
        code='Sub Test()\n    MsgBox "Hello"\nEnd Sub',
        module_type="Standard",
    )


@pytest.fixture
def fixtures_dir():
    """Get path to test fixtures directory."""
    return Path(__file__).parent / "fixtures"
