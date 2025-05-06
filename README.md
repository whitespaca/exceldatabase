# excel-database
==========

.. image:: https://img.shields.io/pypi/v/excel-database.svg
   :target: https://pypi.python.org/pypi/excel-database
   :alt: PyPI version info
.. image:: https://img.shields.io/pypi/pyversions/excel-database.svg
   :target: https://pypi.python.org/pypi/excel-database
   :alt: PyPI supported Python versions

Lightweight Excel-based CRUD “database” backed by an Excel file, powered by `openpyxl`.

## 📦 Installation

```bash
pip install excel-database
```

## 🚀 Quick Start

```python
from excel_database.database import ExcelDatabase

# Initialize with an Excel file path and sheet name
db = ExcelDatabase("./example.xlsx", sheet_name="Sheet1")

# --- Reading ---
rows = db.select({"id": 1})                   # Select all rows matching the query
single = db.get_column_value("id", 1, "name") # Get a single column value from the first matching row

# --- Writing ---
db.insert({"id": 3, "name": "Charlie", "age": 20})
db.update({"id": 2}, {"age": 26})
db.delete({"id": 3})

# --- Sheet Management ---
db.add_sheet("Archive", [{"date": "2025-04-27", "note": "backup"}])
exists = db.is_sheet_exists("Archive")        # Returns 1 if the sheet exists, otherwise None
names = db.get_all_sheet_names()

# --- Column Management ---
count = db.get_column_datas_number("age")     # Count of non-empty 'age' entries
db.add_column("email", "none@example.com")
db.remove_column("email")
```

## 📖 API Reference

### `ExcelDatabase(file_path: str, sheet_name: str = "Sheet1")`

- **file_path**: Path to the `.xlsx` file.
- **sheet_name**: Name of the worksheet to operate on.

| Method                                                      | Description                                                                     |
|-------------------------------------------------------------|---------------------------------------------------------------------------------|
| `select(query: dict) -> list[dict] \| None`                 | Return all rows matching every key/value in `query`. Returns `None` if no rows. |
| `get_column_value(search_col, search_val, target_col)`      | Return the `target_col` value from the first row where `search_col == search_val`, or `None`. |
| `insert(new_row: dict) -> None`                             | Append `new_row` to the sheet and save the file.                                |
| `update(query: dict, update_data: dict) -> None`            | Merge `update_data` into every row matching `query` and save.                   |
| `delete(query: dict) -> None`                               | Delete all rows matching `query` and save.                                      |
| `add_sheet(name: str, initial_data: list[dict]) -> None`    | Create a new sheet named `name`. If `initial_data` is provided, use its keys as headers and append rows. |
| `is_sheet_exists(name: str) -> int \| None`                 | Return `1` if the sheet exists, otherwise `None`.                               |
| `get_all_sheet_names() -> list[str]`                        | List all sheet names in the workbook.                                           |
| `get_column_datas_number(col: str) -> int`                  | Count non-empty cells in column `col`.                                          |
| `add_column(name: str, default: Any) -> None`               | Add a new column with header `name`, filling existing rows with `default`.      |
| `remove_column(name: str) -> None`                          | Remove column `name` from all rows and save.                                    |

## 🤝 Contributing

1. Fork the repository  
2. Create a branch: `git checkout -b feature/your-feature`  
3. Implement your changes and add tests  
4. Open a Pull Request

## 📄 License

MIT License
