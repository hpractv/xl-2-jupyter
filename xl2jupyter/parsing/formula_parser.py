"""Excel formula parsing into AST and dependency extraction."""

import re
from typing import Optional

try:
    from formulas import Parser
except ImportError:
    Parser = None

from xl2jupyter.model.cell import CellModel
from xl2jupyter.utils.logging import get_logger

logger = get_logger(__name__)


class FormulaParser:
    """Parse Excel formulas into AST and extract dependencies."""

    def __init__(self):
        """Initialize formula parser."""
        if Parser is None:
            logger.warning(
                "formulas library not available. Formula parsing will be limited."
            )
            self.parser = None
        else:
            try:
                self.parser = Parser()
            except Exception as e:
                logger.warning(f"Could not initialize formula parser: {e}")
                self.parser = None

    def parse_cell(self, cell: CellModel, sheet_name: str) -> CellModel:
        """
        Parse formula in a cell and extract dependencies.

        Args:
            cell: CellModel instance with formula
            sheet_name: Name of the sheet containing the cell

        Returns:
            Updated CellModel with formula_ast and dependencies
        """
        if not cell.formula:
            return cell

        try:
            # Add '=' prefix if not present
            formula = cell.formula
            if not formula.startswith("="):
                formula = "=" + formula

            # Parse formula to AST
            if self.parser:
                try:
                    ast = self.parser.ast(formula)
                    cell.formula_ast = self._ast_to_dict(ast)
                except Exception as e:
                    logger.debug(f"Could not parse formula {formula}: {e}")
                    cell.formula_ast = None

            # Extract dependencies
            dependencies = self._extract_dependencies(formula, sheet_name)
            cell.dependencies = dependencies

        except Exception as e:
            logger.warning(f"Error parsing formula in cell {cell.address}: {e}")

        return cell

    def _extract_dependencies(self, formula: str, current_sheet: str) -> list[str]:
        """
        Extract cell and range references from formula.

        Args:
            formula: Excel formula string
            current_sheet: Name of the sheet containing the formula

        Returns:
            List of dependency references (e.g., ['A1', 'Sheet2!B5', 'Sheet3!A1:A10'])
        """
        dependencies = []

        # Pattern for cell references: Sheet!A1, Sheet!A1:B10, A1, A1:B10
        # Also handles named ranges and functions
        patterns = [
            # Cross-sheet range: Sheet!A1:B10
            r"([A-Za-z0-9_]+)!([A-Z]+[0-9]+):([A-Z]+[0-9]+)",
            # Cross-sheet cell: Sheet!A1
            r"([A-Za-z0-9_]+)!([A-Z]+[0-9]+)",
            # Range in current sheet: A1:B10
            r"([A-Z]+[0-9]+):([A-Z]+[0-9]+)",
            # Cell in current sheet: A1
            r"\b([A-Z]+[0-9]+)\b",
        ]

        for pattern in patterns:
            matches = re.finditer(pattern, formula)
            for match in matches:
                if len(match.groups()) == 3:
                    # Cross-sheet range
                    sheet = match.group(1)
                    start = match.group(2)
                    end = match.group(3)
                    dep = f"{sheet}!{start}:{end}"
                    if dep not in dependencies:
                        dependencies.append(dep)
                elif len(match.groups()) == 2:
                    if "!" in match.group(0):
                        # Cross-sheet cell
                        sheet = match.group(1)
                        cell_ref = match.group(2)
                        dep = f"{sheet}!{cell_ref}"
                        if dep not in dependencies:
                            dependencies.append(dep)
                    else:
                        # Range in current sheet
                        start = match.group(1)
                        end = match.group(2)
                        dep = f"{start}:{end}"
                        if dep not in dependencies:
                            dependencies.append(dep)
                elif len(match.groups()) == 1:
                    # Single cell - check if it's part of a function name
                    cell_ref = match.group(1)
                    # Avoid matching function names like SUM, AVERAGE, etc.
                    if not self._is_function_name(cell_ref, formula, match.start()):
                        if cell_ref not in dependencies:
                            dependencies.append(cell_ref)

        return dependencies

    def _is_function_name(self, text: str, formula: str, position: int) -> bool:
        """
        Check if text is part of a function name.

        Args:
            text: Text to check
            formula: Full formula
            position: Position of text in formula

        Returns:
            True if text is part of a function name
        """
        # Common Excel function names
        excel_functions = {
            "SUM",
            "AVERAGE",
            "COUNT",
            "MAX",
            "MIN",
            "IF",
            "VLOOKUP",
            "HLOOKUP",
            "INDEX",
            "MATCH",
            "OFFSET",
            "INDIRECT",
            "CONCATENATE",
            "TEXT",
            "DATE",
            "NOW",
            "TODAY",
        }

        if text.upper() in excel_functions:
            return True

        # Check if it's followed by '(' (function call)
        if position + len(text) < len(formula):
            next_char = formula[position + len(text)]
            if next_char == "(":
                return True

        return False

    def _ast_to_dict(self, ast) -> Optional[dict]:
        """
        Convert AST object to dictionary for serialization.

        Args:
            ast: AST object from formulas parser

        Returns:
            Dictionary representation of AST
        """
        if ast is None:
            return None

        try:
            # Try to convert AST to dict
            if hasattr(ast, "__dict__"):
                return self._object_to_dict(ast)
            elif isinstance(ast, (dict, list, str, int, float, bool, type(None))):
                return ast
            else:
                return str(ast)
        except Exception as e:
            logger.debug(f"Error converting AST to dict: {e}")
            return None

    def _object_to_dict(self, obj) -> dict:
        """Recursively convert object to dictionary."""
        if isinstance(obj, (str, int, float, bool, type(None))):
            return obj
        elif isinstance(obj, (list, tuple)):
            return [self._object_to_dict(item) for item in obj]
        elif isinstance(obj, dict):
            return {k: self._object_to_dict(v) for k, v in obj.items()}
        else:
            result = {}
            for key, value in obj.__dict__.items():
                if not key.startswith("_"):
                    result[key] = self._object_to_dict(value)
            return result
