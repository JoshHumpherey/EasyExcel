import pandas as pd
from openpyxl import load_workbook

class PythonExcelConverter():
    """ Class for interfacing between the python client and our excel spreadsheet """

    def __init__(self, filename, extension, sheet_name, output_filename):
        """
        Initializes the Python Excel CONVERTER class, loads the workbook for editing, and also
        adds all of the data into a pandas dataframe
        """
        self.work_book = load_workbook(filename + extension)
        self.work_sheet = self.work_book.active
        self.filename = filename
        self.extension = extension
        self.output_filename = output_filename + self.extension
        x_1 = pd.ExcelFile(self.filename + self.extension)
        data = x_1.parse(sheet_name)
        self.data = data

    def query_column(self, column_number):
        """ Gets a column of data from our pandas dataframe """
        return self.data.iloc[:, column_number]

    def query_row(self, row_number):
        """ Gets a column of data from our pandas dataframe """
        return self.data.iloc[row_number]

    def query_cell(self, col_letter, col_number):
        """ Gets a value from a specific cell """
        index = str(col_letter + col_number)
        return self.work_sheet[index]

    def write_to_cell(self, col_letter, col_number, val):
        """ Writes a value to a specific cell """
        index = str(col_letter + col_number)
        self.work_sheet[index] = val

    def write_list_to_column(self, the_list, col_letter, title, start):
        """ Takes a list object and writes it as a column in our excel sheet """
        self.work_sheet[str(col_letter + str(start))] = title
        start += 1
        for i in range(start, len(the_list)+start):
            str_i = str(i)
            self.work_sheet[str(col_letter + str_i)] = the_list[i-start]

    def save_file(self):
        """ Saves our excel file """
        self.work_book.save(self.output_filename)
        print("Finished analyzing " + str(self.filename + self.extension))
