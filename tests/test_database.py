import unittest
import os
from exceldatabase.database import ExcelDatabase

class TestExcelDatabase(unittest.TestCase):
    def setUp(self):
        self.test_file = "test.xlsx"
        self.db = ExcelDatabase(self.test_file)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_insert_and_select(self):
        self.db.insert({"Name": "Alice", "Age": 25})
        result = self.db.select({"Name": "Alice"})
        self.assertIsNotNone(result)
        self.assertEqual(result[0]["Age"], 25)

if __name__ == "__main__":
    unittest.main()