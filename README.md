# xl-2-jupyter

Convert XLSB Excel workbooks to Jupyter Notebooks with data, formulas, cross-sheet references, and VBA.

## Features

- Extract data from all sheets into pandas DataFrames
- Parse Excel formulas into ASTs and extract dependencies
- Build dependency graphs for formula relationships
- Extract VBA modules (standard, sheet-specific, workbook events)
- Generate comprehensive Jupyter Notebooks with all extracted content

## Requirements

- Python 3.11+
- Microsoft Excel installed (required for xlwings)
- macOS or Windows

## Installation

```bash
pip install -e .
```

## Usage

```bash
xl2jupyter input.xlsb output.ipynb
```

Or as a Python module:

```bash
python -m xl2jupyter input.xlsb output.ipynb
```

## Project Structure

```
xl2jupyter/
├── extraction/     # Excel data extraction via xlwings
├── parsing/        # Formula parsing and dependency analysis
├── model/          # Data models (Workbook, Sheet, Cell, VBA)
├── notebook/       # Jupyter notebook generation
├── cli/            # Command-line interface
└── utils/          # Logging and utilities
```

## Development

```bash
# Install with dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Format code
black xl2jupyter tests

# Lint code
ruff check xl2jupyter tests
```

## License

MIT

