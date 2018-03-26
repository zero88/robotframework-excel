from datetime import datetime, timedelta

import xlsxwriter
from ExcelRobot.reader import ExcelReader
from ExcelRobot.utils import BoolFormat, DataType, DateFormat, NumberFormat
from xlrd import (XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
                  XL_CELL_ERROR, XL_CELL_NUMBER, XL_CELL_TEXT, cellname,
                  xldate_as_tuple)
from xlutils.copy import copy as copy
from xlwt import Workbook, easyxf


class ExcelWriter(ExcelReader):

    def __init__(self, file_path, new_path=None, override=False, date_format=DateFormat(), number_format=NumberFormat(), bool_format=BoolFormat()):
        super().__init__(file_path, date_format, number_format, bool_format)
        self.new_path = new_path
        self.override = override
        self.copied_workbook = None

    def write_number_to_cell(self, sheet_name, column, row, value, decimal_sep=None, thousand_sep=None):
        """
        Using the sheet name the value of the indicated cell is set to be the number given in the parameter.
        """
        if self.workbook:
            my_sheet_index = self.workbook.sheet_names().index(sheet_name)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_NUMBER:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            plain = easyxf('')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), float(value), plain)

    def write_string_to_cell(self, sheet_name, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the string given in the parameter.
        """
        if self.workbook:
            my_sheet_index = self.workbook.sheet_names().index(sheet_name)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_TEXT:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            plain = easyxf('')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), value, plain)

    def write_date_to_cell(self, sheet_name, column, row, value, date_format=None):
        """
        Using the sheet name the value of the indicated cell is set to be the date given in the parameter.
        """
        if self.workbook:
            my_sheet_index = self.workbook.sheet_names().index(sheet_name)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_DATE:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            print(value)
            date_str = value.split('.')
            date_arr = [int(date_str[2]), int(date_str[1]), int(date_str[0])]
            print(date_str, date_arr)
            ymd = datetime(*date_arr)
            plain = easyxf('', num_format_str=date_format or self.date_format)
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), ymd, plain)

    def modify_cell_with(self, sheet_name, column, row, op, val):
        """
        Using the sheet name a cell is modified with the given operation and value.
        """
        my_sheet_index = self.workbook.sheet_names().index(sheet_name)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        curval = cell.value
        if cell.ctype is XL_CELL_NUMBER:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            plain = easyxf('')
            modexpr = str(curval) + op + val
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), eval(modexpr), plain)

    def add_to_date(self, sheet_name, column, row, numdays):
        """
        Using the sheet name the number of days are added to the date in the indicated cell.
        """
        my_sheet_index = self.workbook.sheet_names().index(sheet_name)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        if cell.ctype is XL_CELL_DATE:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            curval = datetime(*xldate_as_tuple(cell.value, self.workbook.datemode))
            newval = curval + timedelta(int(numdays))
            plain = easyxf('', num_format_str=self.date_format)
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), newval, plain)

    def subtract_from_date(self, sheet_name, column, row, numdays):
        """
        Using the sheet name the number of days are subtracted from the date in the indicated cell.
        """
        my_sheet_index = self.workbook.sheet_names().index(sheet_name)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        if cell.ctype is XL_CELL_DATE:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            curval = datetime(*xldate_as_tuple(cell.value, self.workbook.datemode))
            newval = curval - timedelta(int(numdays))
            plain = easyxf('', num_format_str=self.date_format)
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), newval, plain)

    def save_excel(self):
        """
        Saves the Excel file
        """
        self.copied_workbook.save()

    def create_sheet(self, sheet_name):
        """
        Creates and appends new Excel worksheet using the new sheet name to the current workbook.
        """
        self.copied_workbook = copy(self.workbook)
        self.copied_workbook.add_sheet(sheet_name)

    def create_workbook(self, sheet_name):
        """
        Creates a new Excel workbook
        """
        self.copied_workbook = Workbook()
        self.copied_workbook.add_sheet(sheet_name)
