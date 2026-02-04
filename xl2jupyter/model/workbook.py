"""Workbook model representing an Excel workbook."""

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.vba import VBAModule


@dataclass
class WorkbookModel:
    """Model representing an Excel workbook."""

    file_path: Optional[str] = None
    sheets: dict[str, SheetModel] = field(default_factory=dict)
    vba_modules: list[VBAModule] = field(default_factory=list)

    def get_sheet(self, name: str) -> Optional[SheetModel]:
        """
        Get sheet by name.

        Args:
            name: Sheet name

        Returns:
            SheetModel or None if not found
        """
        return self.sheets.get(name)

    def add_sheet(self, sheet: SheetModel) -> None:
        """
        Add or update a sheet.

        Args:
            sheet: SheetModel instance
        """
        self.sheets[sheet.name] = sheet

    def add_vba_module(self, module: VBAModule) -> None:
        """
        Add a VBA module.

        Args:
            module: VBAModule instance
        """
        self.vba_modules.append(module)

    def to_dict(self) -> dict:
        """Convert workbook to dictionary for JSON serialization."""
        return {
            "file_path": self.file_path,
            "sheets": {name: sheet.to_dict() for name, sheet in self.sheets.items()},
            "vba_modules": [module.to_dict() for module in self.vba_modules],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "WorkbookModel":
        """Create workbook from dictionary."""
        workbook = cls(file_path=data.get("file_path"))
        for name, sheet_data in data.get("sheets", {}).items():
            sheet = SheetModel.from_dict(sheet_data)
            workbook.add_sheet(sheet)
        for module_data in data.get("vba_modules", []):
            module = VBAModule.from_dict(module_data)
            workbook.add_vba_module(module)
        return workbook

    def to_json(self, path: Optional[Path] = None) -> str:
        """
        Serialize workbook to JSON.

        Args:
            path: Optional path to save JSON file

        Returns:
            JSON string
        """
        json_str = json.dumps(self.to_dict(), indent=2, default=str)
        if path:
            path.write_text(json_str, encoding="utf-8")
        return json_str

    @classmethod
    def from_json(cls, json_str: str | Path) -> "WorkbookModel":
        """
        Deserialize workbook from JSON.

        Args:
            json_str: JSON string or path to JSON file

        Returns:
            WorkbookModel instance
        """
        if isinstance(json_str, Path):
            json_str = json_str.read_text(encoding="utf-8")

        data = json.loads(json_str)
        return cls.from_dict(data)
