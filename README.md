# README.md
# Excel Database

A simple Excel-based database system using openpyxl.

## Installation
```sh
pip install excel-database
```

## Usage
```python
from excel_database.database import ExcelDatabase

db = ExcelDatabase("data.xlsx")
db.insert({"Name": "Alice", "Age": 25})
print(db.select({"Name": "Alice"}))
```

## License
MIT License