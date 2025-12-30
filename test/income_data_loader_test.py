import pytest
import pandas as pd
from unittest.mock import Mock, patch
from income_data_loader import IncomeDataLoader

@pytest.fixture
def mock_excel_file():
    # Create a mock Excel file with sample data
    data = {
        "Header1": ["Value1", "Value2"],
        "Header2": ["Value3", "Value4"],
    }
    return pd.DataFrame(data)

@patch("income_data_loader.pd.read_excel")
def test_find_title_row(mock_read_excel, mock_excel_file):
    mock_read_excel.return_value = mock_excel_file
    loader = IncomeDataLoader("mock_file.xlsx", ["Header1", "Header2"])
    title_row = loader._find_title_row()
    assert title_row == 0  # Adjust based on expected behavior

@patch("income_data_loader.pd.read_excel")
def test_load_income_data(mock_read_excel, mock_excel_file):
    mock_read_excel.return_value = mock_excel_file
    loader = IncomeDataLoader("mock_file.xlsx",["Header1", "Header2"])
    income_data = loader.load_income_data()
    assert not income_data.empty

@patch("income_data_loader.pd.read_excel")
def test_get_dataframes(mock_read_excel, mock_excel_file):
    mock_read_excel.return_value = mock_excel_file
    loader = IncomeDataLoader("mock_file.xlsx")
    dataframes = loader.get_dataframes()
    assert len(dataframes) > 0

# Add edge cases
@patch("income_data_loader.pd.read_excel")
def test_empty_excel_file(mock_read_excel):
    mock_read_excel.return_value = pd.DataFrame()
    loader = IncomeDataLoader("mock_file.xlsx")
    with pytest.raises(ValueError):
        loader.load_income_data()

@patch("income_data_loader.pd.read_excel")
def test_invalid_file_path(mock_read_excel):
    mock_read_excel.side_effect = FileNotFoundError
    loader = IncomeDataLoader("invalid_file.xlsx")
    with pytest.raises(FileNotFoundError):
        loader.load_income_data()