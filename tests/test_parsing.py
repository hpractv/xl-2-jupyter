"""Tests for formula parsing and dependency analysis."""

import pytest

from xl2jupyter.model.cell import CellModel
from xl2jupyter.parsing.formula_parser import FormulaParser


class TestFormulaParser:
    """Tests for FormulaParser."""

    def test_parse_simple_formula(self):
        """Test parsing a simple formula."""
        parser = FormulaParser()
        cell = CellModel(row=1, column=1, formula="SUM(A1:A5)")
        cell = parser.parse_cell(cell, "Sheet1")

        assert cell.formula == "SUM(A1:A5)"
        assert len(cell.dependencies) > 0
        assert "A1:A5" in cell.dependencies or any("A1" in dep and "A5" in dep for dep in cell.dependencies)

    def test_parse_cross_sheet_reference(self):
        """Test parsing cross-sheet references."""
        parser = FormulaParser()
        cell = CellModel(row=1, column=1, formula="Sheet2!A1")
        cell = parser.parse_cell(cell, "Sheet1")

        assert "Sheet2!A1" in cell.dependencies

    def test_parse_cross_sheet_range(self):
        """Test parsing cross-sheet ranges."""
        parser = FormulaParser()
        cell = CellModel(row=1, column=1, formula="SUM(Sheet2!A1:A10)")
        cell = parser.parse_cell(cell, "Sheet1")

        assert any("Sheet2" in dep for dep in cell.dependencies)

    def test_parse_formula_with_equals(self):
        """Test parsing formula that already has equals sign."""
        parser = FormulaParser()
        cell = CellModel(row=1, column=1, formula="=A1+B1")
        cell = parser.parse_cell(cell, "Sheet1")

        assert cell.formula == "=A1+B1" or cell.formula == "A1+B1"
        assert len(cell.dependencies) >= 2

    def test_parse_cell_without_formula(self):
        """Test parsing cell without formula."""
        parser = FormulaParser()
        cell = CellModel(row=1, column=1, value=10)
        original_cell = CellModel(row=1, column=1, value=10)
        cell = parser.parse_cell(cell, "Sheet1")

        assert cell.formula is None
        assert cell.dependencies == []
        assert cell.value == original_cell.value
