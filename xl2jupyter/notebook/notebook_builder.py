"""Jupyter notebook generation from workbook model."""

from pathlib import Path

import nbformat

from xl2jupyter.model.workbook import WorkbookModel
from xl2jupyter.notebook.cell_templates import (
    create_dataframe_cell,
    create_dependency_graph_cell,
    create_formulas_cell,
    create_imports_cell,
    create_markdown_intro,
    create_vba_cell,
)
from xl2jupyter.parsing.dependency_graph import DependencyGraph
from xl2jupyter.utils.logging import get_logger

logger = get_logger(__name__)


class NotebookBuilder:
    """Build Jupyter notebooks from workbook models."""

    def __init__(
        self,
        workbook: WorkbookModel,
        export_data: bool = True,
        export_formulas: bool = True,
        export_graphs: bool = True,
        export_vba: bool = True,
    ):
        """
        Initialize notebook builder.

        Args:
            workbook: WorkbookModel instance
            export_data: Whether to export data (DataFrames)
            export_formulas: Whether to export formulas
            export_graphs: Whether to export graphs/charts
            export_vba: Whether to export VBA modules
        """
        self.workbook = workbook
        self.notebook = None
        self.export_data = export_data
        self.export_formulas = export_formulas
        self.export_graphs = export_graphs
        self.export_vba = export_vba

    def build(self) -> nbformat.NotebookNode:
        """
        Build Jupyter notebook from workbook model.

        Returns:
            nbformat.NotebookNode instance
        """
        logger.info("Building Jupyter notebook...")

        # Create new notebook
        self.notebook = nbformat.v4.new_notebook()

        # Add markdown introduction
        intro_cell = nbformat.v4.new_markdown_cell(
            create_markdown_intro(
                self.workbook,
                export_data=self.export_data,
                export_formulas=self.export_formulas,
                export_graphs=self.export_graphs,
                export_vba=self.export_vba,
            )
        )
        self.notebook.cells.append(intro_cell)

        # Add imports cell (if exporting data)
        if self.export_data:
            imports_cell = nbformat.v4.new_code_cell(create_imports_cell())
            self.notebook.cells.append(imports_cell)

        # Add DataFrame cells for each sheet (if exporting data)
        if self.export_data:
            for sheet_name, sheet in self.workbook.sheets.items():
                df_cell = nbformat.v4.new_code_cell(create_dataframe_cell(sheet))
                self.notebook.cells.append(df_cell)
                logger.debug(f"Added DataFrame cell for sheet: {sheet_name}")

        # Add formulas cell (if exporting formulas)
        if self.export_formulas:
            formulas_cell = nbformat.v4.new_markdown_cell(create_formulas_cell(self.workbook))
            self.notebook.cells.append(formulas_cell)

            # Build dependency graph and add cell (if exporting formulas)
            try:
                dep_graph = DependencyGraph(self.workbook)
                graph = dep_graph.build()
                dep_cell = nbformat.v4.new_markdown_cell(
                    create_dependency_graph_cell(self.workbook, graph)
                )
                self.notebook.cells.append(dep_cell)
            except Exception as e:
                logger.warning(f"Error building dependency graph: {e}")
                dep_cell = nbformat.v4.new_markdown_cell(
                    create_dependency_graph_cell(self.workbook, None)
                )
                self.notebook.cells.append(dep_cell)

        # Add VBA modules cell (if exporting VBA)
        if self.export_vba and self.workbook.vba_modules:
            vba_cell = nbformat.v4.new_markdown_cell(create_vba_cell(self.workbook))
            self.notebook.cells.append(vba_cell)

        logger.info(f"Notebook built with {len(self.notebook.cells)} cells")
        return self.notebook

    def save(self, output_path: Path | str) -> None:
        """
        Save notebook to file.

        Args:
            output_path: Path to save .ipynb file
        """
        if self.notebook is None:
            self.build()

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"Saving notebook to: {output_path}")

        with open(output_path, "w", encoding="utf-8") as f:
            nbformat.write(self.notebook, f)

        logger.info("Notebook saved successfully")
