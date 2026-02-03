"""Dependency graph construction for Excel formulas."""

import re
from typing import Optional

import networkx as nx

from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.utils.logging import get_logger

logger = get_logger(__name__)


class DependencyGraph:
    """Build and analyze dependency graphs for Excel formulas."""

    def __init__(self, workbook: WorkbookModel):
        """
        Initialize dependency graph builder.

        Args:
            workbook: WorkbookModel instance
        """
        self.workbook = workbook
        self.graph: Optional[nx.DiGraph] = None

    def build(self) -> nx.DiGraph:
        """
        Build dependency graph from workbook formulas.

        Returns:
            NetworkX DiGraph with formula dependencies
        """
        self.graph = nx.DiGraph()

        # Add nodes for all cells with formulas
        for sheet_name, sheet in self.workbook.sheets.items():
            for cell_key, cell in sheet.cells.items():
                if cell.formula:
                    node_id = self._cell_to_node(sheet_name, cell)
                    self.graph.add_node(node_id, sheet=sheet_name, cell=cell)

        # Add edges for dependencies
        for sheet_name, sheet in self.workbook.sheets.items():
            for cell_key, cell in sheet.cells.items():
                if cell.formula and cell.dependencies:
                    source_node = self._cell_to_node(sheet_name, cell)
                    for dep in cell.dependencies:
                        target_nodes = self._dependency_to_nodes(dep, sheet_name)
                        for target_node in target_nodes:
                            if target_node in self.graph:
                                self.graph.add_edge(source_node, target_node)

        logger.info(f"Built dependency graph with {self.graph.number_of_nodes()} nodes and {self.graph.number_of_edges()} edges")
        return self.graph

    def _cell_to_node(self, sheet_name: str, cell) -> str:
        """
        Convert cell to node identifier.

        Args:
            sheet_name: Name of sheet
            cell: CellModel instance

        Returns:
            Node identifier string (e.g., "Sheet1!A1")
        """
        return f"{sheet_name}!{cell.address}"

    def _dependency_to_nodes(self, dependency: str, current_sheet: str) -> list[str]:
        """
        Convert dependency string to list of node identifiers.

        Args:
            dependency: Dependency string (e.g., "A1", "Sheet2!B5", "A1:B10")
            current_sheet: Name of current sheet

        Returns:
            List of node identifiers
        """
        nodes = []

        # Handle range references (e.g., "A1:B10" or "Sheet2!A1:B10")
        if ":" in dependency:
            if "!" in dependency:
                # Cross-sheet range
                sheet, range_part = dependency.split("!")
                start, end = range_part.split(":")
            else:
                # Same-sheet range
                sheet = current_sheet
                start, end = dependency.split(":")

            # Expand range to individual cells
            start_col, start_row = self._parse_cell_address(start)
            end_col, end_row = self._parse_cell_address(end)

            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell_addr = self._column_to_letter(col) + str(row)
                    nodes.append(f"{sheet}!{cell_addr}")

        else:
            # Single cell reference
            if "!" in dependency:
                # Cross-sheet reference
                nodes.append(dependency)
            else:
                # Same-sheet reference
                nodes.append(f"{current_sheet}!{dependency}")

        return nodes

    def _parse_cell_address(self, address: str) -> tuple[int, int]:
        """
        Parse Excel cell address to column and row numbers.

        Args:
            address: Cell address (e.g., "A1", "B10")

        Returns:
            Tuple of (column, row) where both are 1-based
        """
        # Extract column letters and row number
        match = re.match(r"([A-Z]+)([0-9]+)", address.upper())
        if not match:
            raise ValueError(f"Invalid cell address: {address}")

        col_letters = match.group(1)
        row = int(match.group(2))

        # Convert column letters to number
        col = 0
        for char in col_letters:
            col = col * 26 + (ord(char) - ord("A") + 1)

        return (col, row)

    def _column_to_letter(self, column: int) -> str:
        """
        Convert 1-based column number to Excel column letter.

        Args:
            column: Column number (1-based)

        Returns:
            Column letter(s) (e.g., "A", "B", "AA")
        """
        result = ""
        while column > 0:
            column -= 1
            result = chr(65 + (column % 26)) + result
            column //= 26
        return result

    def get_topological_order(self) -> list[str]:
        """
        Get cells in topological order (dependencies first).

        Returns:
            List of node identifiers in topological order
        """
        if self.graph is None:
            self.build()

        try:
            return list(nx.topological_sort(self.graph))
        except nx.NetworkXError:
            # Graph has cycles
            logger.warning("Dependency graph contains cycles")
            return list(self.graph.nodes())

    def get_circular_dependencies(self) -> list[list[str]]:
        """
        Find circular dependencies in the graph.

        Returns:
            List of cycles, where each cycle is a list of node identifiers
        """
        if self.graph is None:
            self.build()

        try:
            cycles = list(nx.simple_cycles(self.graph))
            return cycles
        except Exception as e:
            logger.warning(f"Error finding cycles: {e}")
            return []

    def get_dependents(self, sheet_name: str, cell_address: str) -> list[str]:
        """
        Get all cells that depend on the specified cell.

        Args:
            sheet_name: Name of sheet
            cell_address: Cell address (e.g., "A1")

        Returns:
            List of node identifiers that depend on this cell
        """
        if self.graph is None:
            self.build()

        node_id = f"{sheet_name}!{cell_address}"
        if node_id not in self.graph:
            return []

        # Get all nodes that have a path to this node (reverse graph)
        reverse_graph = self.graph.reverse()
        if node_id not in reverse_graph:
            return []

        # Get all descendants (cells that depend on this one)
        try:
            dependents = list(nx.descendants(reverse_graph, node_id))
            return dependents
        except Exception:
            return []
