import logging
from operator import itemgetter

import natsort
from ExcelRobot.utils import (BoolFormat, DataType, DateFormat, NumberFormat,
                              excel_name2coord, get_file_path, is_file)
from six import PY2
from xlrd import cellname, open_workbook, xldate

LOGGER = logging.getLogger(__name__)


class ExcelReader(object):

    def __init__(self, file_path, date_format=DateFormat(), number_format=NumberFormat(), bool_format=BoolFormat()):
        self.file_path = get_file_path(file_path)
        LOGGER.info('Opening file at %s', self.file_path)
        if not self.file_path or not is_file(self.file_path):
            self._workbook = None
            raise IOError('Excel file is not found') if PY2 else FileNotFoundError('Excel file is not found')
        self._workbook = open_workbook(self.file_path, formatting_info=self.is_xls, on_demand=True)
        self.date_format = date_format
        self.number_format = number_format
        self.bool_format = bool_format

    @property
    def is_xls(self):
        return not self.file_path.endswith('.xlsx')

    @property
    def extension(self):
        return 'xls' if self.is_xls else 'xlsx'

    @property
    def workbook(self):
        if not self._workbook:
            self._workbook = open_workbook(self.file_path, formatting_info=self.is_xls, on_demand=True)
        return self._workbook

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

    def read_cell_data_by_name(self, sheet_name, cell_name, data_type=None, use_format=True):
        """
        Uses the cell name to return the data from that cell.
        """
        col, row = excel_name2coord(cell_name)
        return self.read_cell_data(sheet_name, col, row, data_type, use_format)

    def read_cell_data(self, sheet_name, column, row, data_type=None, use_format=True):
        """
        Uses the column and row to return the data from that cell.

        :Args:
        data_type: Indicate explicit data type to convert
        use_format: Use format to convert data to string
        """
        sheet = self._get_sheet(sheet_name)
        cell = sheet.cell(int(row), int(column))
        ctype = DataType.parse_type_by_value(cell.ctype)
        value = cell.value
        gtype = DataType.parse_type(data_type)
        LOGGER.debug('Given Type: %s', gtype)
        LOGGER.debug('Cell Type: %s', ctype)
        LOGGER.debug('Cell Value: %s', value)
        if DataType.is_date(ctype):
            if gtype and not DataType.is_date(gtype):
                raise ValueError('Cell type does not match with given data type')
            date_value = xldate.xldate_as_datetime(value, self.date_format.datemode)
            if use_format:
                return self.date_format.format(gtype, date_value)
            elif DataType.DATE == gtype:
                return date_value.date()
            elif DataType.TIME == gtype:
                return date_value.time()
            return date_value
        if DataType.is_number(ctype):
            if gtype and not DataType.is_number(gtype):
                raise ValueError('Cell type does not match with given data type')
            return self.number_format.format(gtype, value) if use_format else value
        if DataType.is_bool(ctype):
            if gtype and not DataType.is_bool(gtype):
                raise ValueError('Cell type does not match with given data type')
            return self.bool_format.format(value) if use_format else value
        return value

    def check_cell_type(self, sheet_name, column, row, data_type):
        """
        Checks the type of value that is within the cell of the sheet name selected.
        """
        ctype = DataType.parse_type_by_value(self._get_cell_type(sheet_name, column, row))
        gtype = DataType.parse_type(data_type)
        LOGGER.debug('Given Type: %s', gtype)
        LOGGER.debug('Cell Type: %s', ctype)
        return ctype == gtype
