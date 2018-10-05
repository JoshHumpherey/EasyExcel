import unittest
import converter
import datetime
import xlsxwriter

class Test_Converter(unittest.TestCase):

    def CreateSpreasheet(self, filename):
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        # Some data we want to write to the worksheet.
        expenses = (
            ['Rent', 1000],
            ['Gas',   100],
            ['Food',  300],
            ['Gym',    50],
        )

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for item, cost in (expenses):
            worksheet.write(row, col,     item)
            worksheet.write(row, col + 1, cost)
            row += 1

        # Write a total using a formula.
        worksheet.write(row, 0, 'Total')
        worksheet.write(row, 1, '=SUM(B1:B4)')

        workbook.close()

    def test_query_column(self):
        self.CreateSpreasheet('test_sheet.xlsx')
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        expected = ['Gas', 'Food', 'Gym', 'Total']
        actual = list(test_converter.query_column(0))
        self.assertEqual(expected, actual)

    def test_query_row(self):
        self.CreateSpreasheet('test_sheet.xlsx')
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        expected = ['Food', 300]
        actual = list(test_converter.query_row(1))
        self.assertEqual(expected, actual)

    def test_query_cell(self):
        self.CreateSpreasheet('test_sheet.xlsx')
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        actual = test_converter.query_cell('A',0)
        self.assertEqual('Rent', actual)

    def test_write_to_cell(self):
        self.CreateSpreasheet('test_sheet.xlsx')
        test_converter = converter.PythonExcelConverter('test_sheet', '.xlsx', 'Sheet1','test_sheet')
        test_converter.write_to_cell('A', 0, 'Eggs')
        actual = test_converter.query_cell('A',0)
        self.assertEqual('Eggs', actual)
        test_converter.write_to_cell('A', 0, 50)
        actual = test_converter.query_cell('A',0)
        self.assertEqual(50, actual)

if __name__ == '__main__':
    unittest.main()
