import unittest
import converter

class Test_Converter(unittest.TestCase):
    test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')

    def test_query_column(self):
        actual_col = test_converter.query_column(0)
        expected_col = [1, 6, 7, 8, 9]
        self.assertEqual(expected_col, actual_col)



if __name__ == '__main__':
    unittest.main()
