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
        help="Skip VBA extraction",
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

        # Extract workbook data
        logger.info("Extracting workbook data...")
        reader = ExcelReader(input_path)
        workbook = reader.read()

        # Skip VBA extraction if requested
        if args.no_vba:
            workbook.vba_modules.clear()
            logger.info("VBA extraction skipped (--no-vba flag)")

        # Parse formulas
        logger.info("Parsing formulas...")
        parser = FormulaParser()
        for sheet_name, sheet in workbook.sheets.items():
            for cell in sheet.cells.values():
                if cell.formula:
                    parser.parse_cell(cell, sheet_name)

        # Build notebook
        logger.info("Building Jupyter notebook...")
        builder = NotebookBuilder(workbook)
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
