"""VBA module model."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class VBAModule:
    """Model representing a VBA module."""

    name: str
    code: str
    module_type: str  # 'Standard', 'Sheet', 'Workbook', 'Class', etc.
    sheet_name: Optional[str] = None  # For sheet-specific modules

    def to_dict(self) -> dict:
        """Convert VBA module to dictionary for JSON serialization."""
        return {
            "name": self.name,
            "code": self.code,
            "module_type": self.module_type,
            "sheet_name": self.sheet_name,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "VBAModule":
        """Create VBA module from dictionary."""
        return cls(
            name=data["name"],
            code=data["code"],
            module_type=data["module_type"],
            sheet_name=data.get("sheet_name"),
        )

