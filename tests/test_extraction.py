"""Tests for Excel extraction layer."""

import pytest
from pathlib import Path

from xl2jupyter.extraction.excel_reader import ExcelReader
from xl2jupyter.extraction.vba_extractor import VBAExtractor


def test_excel_reader_validation(tmp_path):
    """Test ExcelReader file validation."""
    # Test non-existent file
    with pytest.raises(FileNotFoundError):
        reader = ExcelReader(tmp_path / "nonexistent.xlsb")
        reader.read()

    # Test wrong file extension
    test_file = tmp_path / "test.txt"
    test_file.write_text("test")
    with pytest.raises(ValueError, match="must be .xlsb"):
        reader = ExcelReader(test_file)
        reader.read()


@pytest.mark.skipif(
    True, reason="Requires Excel and xlwings - run manually with actual Excel file"
)
def test_excel_reader_extraction(fixtures_dir):
    """
    Test extracting data from Excel file.

    This test requires:
    1. Excel to be installed
    2. Test workbook to exist
    """
    test_workbook = fixtures_dir / "comprehensive_test.xlsb"

    if not test_workbook.exists():
        pytest.skip("Test workbook not found")

    reader = ExcelReader(test_workbook)
    workbook = reader.read()

    assert workbook is not None
    assert workbook.file_path == str(test_workbook)
    assert len(workbook.sheets) > 0

    # Check that sheets have data
    for sheet_name, sheet in workbook.sheets.items():
        assert sheet.name == sheet_name
        assert len(sheet.cells) > 0 or sheet.used_range is None


def test_vba_extractor_initialization():
    """Test VBAExtractor can be initialized (without actual Excel)."""
    # This test just checks the class can be imported and has the right structure
    assert VBAExtractor is not None
    assert hasattr(VBAExtractor, "extract_all_modules")
