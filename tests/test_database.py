# tests/test_database.py

import os
import pytest
from openpyxl import Workbook
from excel_database.database import ExcelDatabase

@pytest.fixture
def tmp_excel_file(tmp_path):
    """
    Create a temporary Excel file with:
    - Sheet1 headers: id, name, age
    - Two data rows
    """
    file = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name", "age"])
    ws.append([1, "Alice", 30])
    ws.append([2, "Bob", 25])
    wb.save(file)
    return str(file)

def test_select_existing(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    result = db.select({"name": "Alice"})
    assert result == [{"id": 1, "name": "Alice", "age": 30}]

def test_select_non_existing(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    assert db.select({"name": "Charlie"}) is None

def test_get_column_value(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    assert db.get_column_value("id", 2, "name") == "Bob"
    assert db.get_column_value("id", 3, "name") is None

def test_insert(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    db.insert({"id": 3, "name": "Charlie", "age": 20})
    # Reload as a new instance to verify the file was updated
    db2 = ExcelDatabase(tmp_excel_file)
    names = [r["name"] for r in db2.data]
    assert "Charlie" in names

def test_update(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    db.update({"id": 1}, {"age": 31})
    assert db.get_column_value("id", 1, "age") == 31

def test_delete(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    db.delete({"id": 2})
    assert db.select({"id": 2}) is None

def test_add_sheet_and_exists(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    assert db.is_sheet_exists("NewSheet") is None
    db.add_sheet("NewSheet", [{"id": 10, "value": "test"}])
    assert db.is_sheet_exists("NewSheet") == 1
    assert "NewSheet" in db.get_all_sheet_names()

def test_get_all_sheet_names(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    names = db.get_all_sheet_names()
    # The default sheet should be included
    assert "Sheet1" in names

def test_get_column_datas_number(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    # There should be two non-empty values in the 'age' column
    assert db.get_column_datas_number("age") == 2

def test_add_and_remove_column(tmp_excel_file):
    db = ExcelDatabase(tmp_excel_file)
    db.add_column("email", "unknown@example.com")
    assert all(row.get("email") == "unknown@example.com" for row in db.data)
    db.remove_column("email")
    assert all("email" not in row for row in db.data)
