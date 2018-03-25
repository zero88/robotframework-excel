import functools
import logging
import re
import os
import os.path as path
from enum import Enum
from operator import itemgetter

import natsort
from xlrd import (XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
                  XL_CELL_ERROR, XL_CELL_NUMBER, XL_CELL_TEXT, cellname,
                  open_workbook, xldate)

LOGGER = logging.getLogger(__name__)


class DataType(Enum):

    DATE = XL_CELL_DATE
    TEXT = XL_CELL_TEXT
    NUMBER = XL_CELL_NUMBER
    BLANK = XL_CELL_BLANK
    EMPTY = XL_CELL_EMPTY
    ERROR = XL_CELL_ERROR
    BOOL = XL_CELL_BOOLEAN


class ExcelReader:

    @staticmethod
    def _get_file_path(file_name):
        if not file_name.endswith('.xlsx') and not file_name.endswith('.xls'):
            raise ValueError('Only support file with extenstion: xls and xlsx')
        file_path = path.normpath(file_name)
        if len(path.splitdrive(file_path)) > 1:
            return path.join(os.getcwd(), file_name)
        return file_name

    @staticmethod
    def _excel2num(cell_name):
        matrix = list(filter(lambda x: x.strip(), re.split(r'(\d+)', cell_name.upper())))
        LOGGER.debug('Matrix: %s', matrix)
        if len(matrix) != 2 or not re.match(r'[A-Z]+', matrix[0]) or not re.match(r'\d+', matrix[1]):
            raise ValueError('Cell name is invalid')
        col = int(functools.reduce(lambda s, a: s * 26 + ord(a) - ord('A') + 1, matrix[0], 0)) - 1
        row = int(matrix[1]) - 1
        LOGGER.debug('Col, Row: %d, %d', col, row)
        return col, row

    @staticmethod
    def _excel2format(dt_format):
        # TODO: Enhance by regex
        if dt_format == 'yyyy-mm-dd':
            return '%Y-%m-%d'
        if dt_format == 'hh:mm:ss AM/PM':
            return '%I:%M:%S %p'
        if dt_format == 'yyyy-mm-dd hh:mm':
            return '%Y-%m-%d %H:%M'
        return ''

    def __init__(self, file_path,
                 date_format='yyyy-mm-dd', time_format='hh:mm:ss AM/PM',
                 datetime_format='yyyy-mm-dd hh:mm', decimal_sep='.', thousand_sep=','):
        file_path = self._get_file_path(file_path)
        LOGGER.info('Opening file at %s', file_path)
        self.is_xls = not file_path.endswith('.xlsx')
        self.workbook = open_workbook(file_path, formatting_info=self.is_xls, on_demand=True)
        self.date_format = self._excel2format(date_format)
        self.time_format = self._excel2format(time_format)
        self.datetime_format = self._excel2format(datetime_format)
        self.decimal_sep = decimal_sep
        self.thousand_sep = thousand_sep

    def _get_sheet(self, sheet_name):
        return self.workbook.sheet_by_name(sheet_name)

    def _get_cell_type(self, sheet_name, column, row):
        return self._get_sheet(sheet_name).cell_type(int(row), int(column))

    def get_sheet_names(self):
        """
        Returns the names of all the worksheets in the current workbook.
        """
        return self.workbook.sheet_names()

    def get_number_of_sheets(self):
        """
        Returns the number of worksheets in the current workbook.
        """
        return self.workbook.nsheets

    def get_column_count(self, sheet_name):
        """
        Returns the specific number of columns of the sheet name specified.
        """
        return self._get_sheet(sheet_name).ncols

    def get_row_count(self, sheet_name):
        """
        Returns the specific number of rows of the sheet name specified.
        """
        return self._get_sheet(sheet_name).nrows

    def get_column_values(self, sheet_name, column, include_empty_cells=True):
        """
        Returns the specific column values of the sheet name specified.
        """
        sheet = self._get_sheet(sheet_name)
        data = {}
        for row_index in range(sheet.nrows):
            cell = cellname(row_index, int(column))
            value = sheet.cell(row_index, int(column)).value
            data[cell] = value
        if not include_empty_cells:
            data = dict([(k, v) for (k, v) in data.items() if v])
        return natsort.natsorted(data.items(), key=itemgetter(0))

    def get_row_values(self, sheet_name, row, include_empty_cells=True):
        """
        Returns the specific row values of the sheet name specified.
        """
        sheet = self._get_sheet(sheet_name)
        data = {}
        for col_index in range(sheet.ncols):
            cell = cellname(int(row), col_index)
            value = sheet.cell(int(row), col_index).value
            data[cell] = value
        if not include_empty_cells:
            data = dict([(k, v) for (k, v) in data.items() if v])
        return natsort.natsorted(data.items(), key=itemgetter(0))

    def get_sheet_values(self, sheet_name, include_empty_cells=True):
        """
        Returns the values from the sheet name specified.
        """
        sheet = self._get_sheet(sheet_name)
        data = {}
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = cellname(row_index, col_index)
                value = sheet.cell(row_index, col_index).value
                data[cell] = value
        if not include_empty_cells:
            data = dict([(k, v) for (k, v) in data.items() if v])
        return natsort.natsorted(data.items(), key=itemgetter(0))

    def get_workbook_values(self, include_empty_cells=True):
        """
        Returns the values from each sheet of the current workbook.
        """
        sheet_data = []
        workbook_data = []
        for sheet_name in self.workbook.sheet_names():
            sheet_data = self.get_sheet_values(sheet_name, include_empty_cells)
            sheet_data.insert(0, sheet_name)
            workbook_data.append(sheet_data)
        return workbook_data

    def read_cell_data_by_name(self, sheet_name, cell_name, data_format=None):
        """
        Uses the cell name to return the data from that cell.
        """
        col, row = self._excel2num(cell_name)
        return self.read_cell_data_by_coordinates(sheet_name, col, row, data_format)

    def read_cell_data_by_coordinates(self, sheet_name, column, row, data_format=None):
        """
        Uses the column and row to return the data from that cell.
        """
        sheet = self._get_sheet(sheet_name)
        cell = sheet.cell(int(row), int(column))
        ctype = cell.ctype
        LOGGER.info('Type: %s', cell.ctype)
        LOGGER.info('Value: %s', cell.value)
        if ctype == DataType.DATE.value:
            return xldate.xldate_as_datetime(cell.value, 0).strftime(self._excel2format(data_format) or self.date_format)
        return cell.value

    def check_cell_type(self, sheet_name, column, row, data_type):
        """
        Checks the type of value that is within the cell of the sheet name selected.
        """
        excel_type = DataType[data_type]
        if excel_type is None:
            raise ValueError('Not support %s' % data_type)
        return excel_type.value == self._get_cell_type(sheet_name, column, row)
