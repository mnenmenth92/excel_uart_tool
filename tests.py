import pytest
from xlsx_driver import ExcelColumnReader
import os

CONFIG_DIR = "test_configs"

def test_wrong_file_config():
    """
    Test case: config file points to a non-existing Excel file.
    Expected error: -1000
    """
    config_path = os.path.join(CONFIG_DIR, "wrong_file.ini")
    reader = ExcelColumnReader(config_path)
    result = reader.read_column("A")
    assert result == -1000, f"Expected -1000 for missing file, got {result}"

def test_wrong_sheet_config():
    """
    Test case: config file points to a non-existing sheet.
    Expected error: -1001
    """
    config_path = os.path.join(CONFIG_DIR, "wrong_sheet.ini")
    reader = ExcelColumnReader(config_path)
    result = reader.read_column("A")
    assert result == -1001, f"Expected -1001 for missing sheet, got {result}"


def test_no_param_config():
    """
    Test case: check no parameters error.
    Expected error: -1002
    """
    reader = ExcelColumnReader("config.ini")
    result = reader.read_column("B")
    assert result == -1002, f"Expected -1002 for empty param column, got {result}"

def test_correct_correct_config_no_err():
    """
    Test case: reading the correct Excel file and sheet returns correct number of parameters.
    """
    reader = ExcelColumnReader("config.ini")

    assert reader.error == 0, f"Expected 0  - no error {reader.error}"

def test_no_config_err():
    """
    Test case: reading the correct Excel file and sheet returns correct number of parameters.
    """
    reader = ExcelColumnReader("conf4ig.ini")

    assert reader.error == -1005, f"Expected 0  - no error {reader.error}"


def test_read_correct_data_length():
    """
    Test case: reading the correct Excel file and sheet returns correct number of parameters.
    """
    number_of_parameters = 23
    reader = ExcelColumnReader("config.ini")
    result = reader.read_column("A")

    # Make sure we got a list, not an error code
    assert isinstance(result, list), f"Expected list, got {result}"
    assert len(result) == number_of_parameters, f"Expected {number_of_parameters} parameters, got {len(result)}"