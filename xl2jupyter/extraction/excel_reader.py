"""Excel workbook extraction using xlwings."""

from pathlib import Path
from typing import Optional

import xlwings as xw

from xl2jupyter.model.cell import CellModel
from xl2jupyter.model.sheet import SheetModel
from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.extraction.vba_extractor import VBAExtractor
from xl2jupyter.utils.logging import get_logger

logger = get_logger(__name__)


class ExcelReader:
    """Extract data and formulas from Excel workbooks using xlwings."""

    def __init__(self, file_path: Path | str):
        """
        Initialize Excel reader.

        Args:
            file_path: Path to .xlsb file
        """
        self.file_path = Path(file_path)
        self.workbook: Optional[xw.Book] = None

    def read(self) -> WorkbookModel:
        """
        Read workbook and extract all data, formulas, and VBA.

        Returns:
            WorkbookModel instance with all extracted data

        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file is not a .xlsb file
            Exception: If Excel automation fails
        """
        if not self.file_path.exists():
            raise FileNotFoundError(f"File not found: {self.file_path}")

        if self.file_path.suffix.lower() != ".xlsb":
            raise ValueError(f"File must be .xlsb, got: {self.file_path.suffix}")

        logger.info(f"Opening workbook: {self.file_path}")

        try:
            # Open workbook
            self.workbook = xw.Book(str(self.file_path))
            logger.info(f"Workbook opened successfully: {self.workbook.name}")

            # Create workbook model
            workbook_model = WorkbookModel(file_path=str(self.file_path))

            # Extract sheets
            for sheet in self.workbook.sheets:
                try:
                    sheet_model = self._extract_sheet(sheet)
                    workbook_model.add_sheet(sheet_model)
                    logger.info(f"Extracted sheet: {sheet.name}")
                except Exception as e:
                    logger.error(f"Error extracting sheet {sheet.name}: {e}")

            # Extract VBA modules
            try:
                vba_extractor = VBAExtractor(self.workbook)
                vba_modules = vba_extractor.extract_all_modules()
                for module in vba_modules:
                    workbook_model.add_vba_module(module)
                logger.info(f"Extracted {len(vba_modules)} VBA modules")
            except Exception as e:
                logger.warning(f"Error extracting VBA modules: {e}")

            return workbook_model

        except Exception as e:
            logger.error(f"Error reading workbook: {e}")
            raise
        finally:
            # Close workbook if opened
            if self.workbook is not None:
                try:
                    self.workbook.close()
                except Exception:
                    pass

    def _extract_sheet(self, sheet: xw.Sheet) -> SheetModel:
        """
        Extract data from a single sheet.

        Args:
            sheet: xlwings Sheet instance

        Returns:
            SheetModel instance
        """
        sheet_model = SheetModel(name=sheet.name)

        try:
            # Get used range
            used_range = sheet.used_range
            if used_range is None:
                logger.debug(f"Sheet {sheet.name} has no used range")
                return sheet_model

            # Get range coordinates
            first_row = used_range.row
            last_row = used_range.last_cell.row
            first_col = used_range.column
            last_col = used_range.last_cell.column

            sheet_model.used_range = (first_row, last_row, first_col, last_col)

            logger.debug(
                f"Sheet {sheet.name} used range: "
                f"R{first_row}:R{last_row}, C{first_col}:C{last_col}"
            )

            # Extract cells
            for row in range(first_row, last_row + 1):
                for col in range(first_col, last_col + 1):
                    try:
                        cell = sheet.cells(row, col)
                        cell_model = self._extract_cell(cell, row, col)
                        if cell_model:
                            sheet_model.set_cell(cell_model)
                    except Exception as e:
                        logger.debug(f"Error extracting cell ({row}, {col}): {e}")

        except Exception as e:
            logger.warning(f"Error extracting sheet {sheet.name}: {e}")

        return sheet_model

    def _extract_cell(self, cell: xw.Range, row: int, col: int) -> Optional[CellModel]:
        """
        Extract data from a single cell.

        Args:
            cell: xlwings Range instance (single cell)
            row: Row number (1-based)
            col: Column number (1-based)

        Returns:
            CellModel instance or None if cell is empty
        """
        try:
            # Get cell value
            value = cell.value

            # Get formula if present
            formula = None
            try:
                formula = cell.formula
                # Remove leading '=' if present (we'll add it back when needed)
                if formula and formula.startswith("="):
                    formula = formula[1:]
            except Exception:
                # Cell doesn't have a formula
                pass

            # Determine data type
            data_type = None
            if value is not None:
                if isinstance(value, (int, float)):
                    data_type = "number"
                elif isinstance(value, bool):
                    data_type = "boolean"
                elif isinstance(value, str):
                    data_type = "text"
                else:
                    data_type = str(type(value).__name__)

            # Only create cell model if there's data or a formula
            if value is not None or formula:
                return CellModel(
                    row=row,
                    column=col,
                    value=value,
                    formula=formula,
                    data_type=data_type,
                )

        except Exception as e:
            logger.debug(f"Error extracting cell ({row}, {col}): {e}")

        return None
