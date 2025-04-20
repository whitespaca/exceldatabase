# It's not working now

# README.md
# Excel Database

A simple Excel-based database system using openpyxl.

## Installation
```sh
pip install exceldatabase
```

## Usage
```python
from exceldatabase.database import ExcelDatabase

db = ExcelDatabase("data.xlsx")
db.insert({"Name": "Alice", "Age": 25})
print(db.select({"Name": "Alice"}))
```

## License
MIT License