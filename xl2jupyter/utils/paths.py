"""Path utilities."""

from pathlib import Path
from typing import Optional


def normalize_path(path: str | Path) -> Path:
    """
    Normalize a path string or Path object.

    Args:
        path: Path string or Path object

    Returns:
        Normalized Path object
    """
    return Path(path).expanduser().resolve()


def validate_input_file(path: Path) -> Path:
    """
    Validate that an input file exists and has the correct extension.

    Args:
        path: Path to input file

    Returns:
        Validated Path object

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file doesn't have .xlsb extension
    """
    path = normalize_path(path)

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    if path.suffix.lower() != ".xlsb":
        raise ValueError(f"Input file must be a .xlsb file, got: {path.suffix}")

    return path


def validate_output_path(path: Optional[str | Path], input_path: Path) -> Path:
    """
    Generate or validate output path for notebook.

    Args:
        path: Optional output path
        input_path: Input file path

    Returns:
        Output Path object
    """
    if path:
        output_path = normalize_path(path)
    else:
        # Default: same name as input but with .ipynb extension
        output_path = input_path.with_suffix(".ipynb")

    # Ensure parent directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    return output_path
