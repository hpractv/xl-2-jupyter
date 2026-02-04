"""Cell model representing a single Excel cell."""

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class CellModel:
    """Model representing a single Excel cell."""

    row: int
    column: int
    value: Any = None
    formula: Optional[str] = None
    formula_ast: Optional[dict] = None
    dependencies: list[str] = field(default_factory=list)
    data_type: Optional[str] = None

    def to_dict(self) -> dict:
        """Convert cell to dictionary for JSON serialization."""
        return {
            "row": self.row,
            "column": self.column,
            "value": self.value,
            "formula": self.formula,
            "formula_ast": self.formula_ast,
            "dependencies": self.dependencies,
            "data_type": self.data_type,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "CellModel":
        """Create cell from dictionary."""
        return cls(
            row=data["row"],
            column=data["column"],
            value=data.get("value"),
            formula=data.get("formula"),
            formula_ast=data.get("formula_ast"),
            dependencies=data.get("dependencies", []),
            data_type=data.get("data_type"),
        )

    @property
    def address(self) -> str:
        """Get Excel cell address (e.g., 'A1')."""
        col_letter = self._column_to_letter(self.column)
        return f"{col_letter}{self.row}"

    @staticmethod
    def _column_to_letter(column: int) -> str:
        """Convert 1-based column number to Excel column letter."""
        result = ""
        while column > 0:
            column -= 1
            result = chr(65 + (column % 26)) + result
            column //= 26
        return result
