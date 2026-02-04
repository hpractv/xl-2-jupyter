"""Sheet model representing an Excel worksheet."""

from dataclasses import dataclass, field
from typing import Any, Optional

from xl2jupyter.model.cell import CellModel


@dataclass
class SheetModel:
    """Model representing an Excel worksheet."""

    name: str
    cells: dict[str, CellModel] = field(default_factory=dict)
    used_range: Optional[tuple[int, int, int, int]] = None  # (min_row, max_row, min_col, max_col)

    def get_cell(self, row: int, column: int) -> Optional[CellModel]:
        """
        Get cell by row and column.

        Args:
            row: 1-based row number
            column: 1-based column number

        Returns:
            CellModel or None if not found
        """
        key = f"{row},{column}"
        return self.cells.get(key)

    def set_cell(self, cell: CellModel) -> None:
        """
        Set or update a cell.

        Args:
            cell: CellModel instance
        """
        key = f"{cell.row},{cell.column}"
        self.cells[key] = cell
        self._update_used_range(cell.row, cell.column)

    def _update_used_range(self, row: int, column: int) -> None:
        """Update the used range based on cell position."""
        if self.used_range is None:
            self.used_range = (row, row, column, column)
        else:
            min_row, max_row, min_col, max_col = self.used_range
            self.used_range = (
                min(min_row, row),
                max(max_row, row),
                min(min_col, column),
                max(max_col, column),
            )

    def get_data_array(self) -> list[list[Any]]:
        """
        Get sheet data as a 2D array.

        Returns:
            2D list of cell values
        """
        if self.used_range is None:
            return []

        min_row, max_row, min_col, max_col = self.used_range
        data = []

        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                cell = self.get_cell(row, col)
                row_data.append(cell.value if cell else None)
            data.append(row_data)

        return data

    def to_dict(self) -> dict:
        """Convert sheet to dictionary for JSON serialization."""
        return {
            "name": self.name,
            "cells": {key: cell.to_dict() for key, cell in self.cells.items()},
            "used_range": self.used_range,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SheetModel":
        """Create sheet from dictionary."""
        sheet = cls(name=data["name"], used_range=data.get("used_range"))
        for key, cell_data in data.get("cells", {}).items():
            cell = CellModel.from_dict(cell_data)
            sheet.set_cell(cell)
        return sheet
