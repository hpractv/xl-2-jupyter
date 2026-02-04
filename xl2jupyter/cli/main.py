"""Command-line interface for xl2jupyter."""

import sys
from pathlib import Path

from xl2jupyter.extraction.excel_reader import ExcelReader
from xl2jupyter.notebook.notebook_builder import NotebookBuilder
from xl2jupyter.parsing.formula_parser import FormulaParser
from xl2jupyter.utils.logging import setup_logging
from xl2jupyter.utils.paths import validate_input_file, validate_output_path


def main():
    """Main CLI entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert XLSB Excel workbooks to Jupyter Notebooks",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "input",
        type=str,
        help="Path to input .xlsb file",
    )

    parser.add_argument(
        "output",
        type=str,
        nargs="?",
        default=None,
        help="Path to output .ipynb file (default: same as input with .ipynb extension)",
    )

    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    parser.add_argument(
        "--no-vba",
        action="store_true",
        help="Skip VBA extraction (deprecated, use --export instead)",
    )

    parser.add_argument(
        "--export",
        type=str,
        nargs="+",
        choices=["data", "formulas", "graphs", "vba", "all"],
        default=["all"],
        help=(
            "Specify what to export: data, formulas, graphs, vba, or all. "
            "Can specify multiple values (e.g., --export data formulas). "
            "Default: all"
        ),
    )

    args = parser.parse_args()

    # Setup logging
    logger = setup_logging(verbose=args.verbose)

    try:
        # Validate input file
        input_path = validate_input_file(args.input)
        logger.info(f"Input file: {input_path}")

        # Validate/generate output path
        output_path = validate_output_path(args.output, input_path)
        logger.info(f"Output file: {output_path}")

        # Determine what to export
        export_items = set(args.export)
        if "all" in export_items:
            export_items = {"data", "formulas", "graphs", "vba"}
        
        # Handle deprecated --no-vba flag
        if args.no_vba:
            logger.warning("--no-vba flag is deprecated. Use '--export data formulas graphs' instead")
            export_items.discard("vba")
        
        export_data = "data" in export_items
        export_formulas = "formulas" in export_items
        export_graphs = "graphs" in export_items
        export_vba = "vba" in export_items
        
        logger.info(f"Export settings: data={export_data}, formulas={export_formulas}, graphs={export_graphs}, vba={export_vba}")

        # Extract workbook data
        logger.info("Extracting workbook data...")
        reader = ExcelReader(input_path)
        workbook = reader.read()

        # Skip VBA extraction if not requested
        if not export_vba:
            workbook.vba_modules.clear()
            logger.info("VBA extraction skipped")

        # Parse formulas (only if formulas are being exported)
        if export_formulas:
            logger.info("Parsing formulas...")
            parser = FormulaParser()
            for sheet_name, sheet in workbook.sheets.items():
                for cell in sheet.cells.values():
                    if cell.formula:
                        parser.parse_cell(cell, sheet_name)

        # Build notebook
        logger.info("Building Jupyter notebook...")
        builder = NotebookBuilder(
            workbook,
            export_data=export_data,
            export_formulas=export_formulas,
            export_graphs=export_graphs,
            export_vba=export_vba,
        )
        builder.build()
        builder.save(output_path)

        logger.info(f"Successfully converted {input_path} to {output_path}")
        print(f"✓ Conversion complete: {output_path}")

    except FileNotFoundError as e:
        logger.error(f"File not found: {e}")
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        logger.error(f"Invalid input: {e}")
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        logger.info("Conversion cancelled by user")
        print("\nConversion cancelled.", file=sys.stderr)
        sys.exit(130)
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
