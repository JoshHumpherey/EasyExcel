import unittest
import converter
import datetime

class Test_Converter(unittest.TestCase):

    def test_query_column(self):
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        expected = ['John Smith', 'Sarah Kay', 'Mr. Developer', 'Patrick Stewart']
        actual = list(test_converter.query_column(0))
        self.assertEqual(expected, actual)

    def test_query_row(self):
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        expected = ['John Smith', datetime.datetime(1994, 9, 1, 0, 0), 3, 7.4]
        actual = list(test_converter.query_row(0))
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()
