import functools
import logging
import numbers
import os
import os.path as path
import re
import shutil
import string
import tempfile
from datetime import date, datetime, time
from enum import Enum
from random import choice

from xlrd import (XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
                  XL_CELL_ERROR, XL_CELL_NUMBER, XL_CELL_TEXT)

LOGGER = logging.getLogger(__name__)


def random_temp_file(ext='txt'):
    return path.join(tempfile.gettempdir(), random_name() + '.' + ext)


def random_name():
    return ''.join(choice(string.ascii_letters) for x in range(8))


def is_file(file_path):
    return path.isfile(file_path)


def del_file(file_path):
    os.remove(file_path)


def copy_file(src, dest, force=False):
    if src == dest:
        raise ValueError('Same file error')
    if force and is_file(dest):
        del_file(dest)
    shutil.copy(src, dest)


def get_file_path(file_name):
    """
    Return None if file_name is None
    """
    if not file_name:
        return None
    if not file_name.endswith('.xlsx') and not file_name.endswith('.xls'):
        raise ValueError('Only support file with extenstion: xls and xlsx')
    file_path = path.normpath(file_name)
    if len(path.splitdrive(file_path)) > 1:
        return path.join(os.getcwd(), file_name)
    return file_name


def excel_name2coord(cell_name):
    matrix = list(filter(lambda x: x.strip(), re.split(r'(\d+)', cell_name.upper())))
    LOGGER.debug('Matrix: %s', matrix)
    if len(matrix) != 2 or not re.match(r'[A-Z]+', matrix[0]) or not re.match(r'\d+', matrix[1]):
        raise ValueError('Cell name is invalid')
    col = int(functools.reduce(lambda s, a: s * 26 + ord(a) - ord('A') + 1, matrix[0], 0)) - 1
    row = int(matrix[1]) - 1
    LOGGER.debug('Col, Row: %d, %d', col, row)
    return col, row



class DataType(Enum):

    DATE = XL_CELL_DATE
    TIME = XL_CELL_DATE * 11
    DATE_TIME = XL_CELL_DATE * 13
    TEXT = XL_CELL_TEXT
    NUMBER = XL_CELL_NUMBER
    CURRENCY = XL_CELL_NUMBER * 11
    PERCENTAGE = XL_CELL_NUMBER * 13
    BLANK = XL_CELL_BLANK
    EMPTY = XL_CELL_EMPTY
    ERROR = XL_CELL_ERROR
    BOOL = XL_CELL_BOOLEAN

    @staticmethod
    def is_date(dtype, value=None):
        if value and not dtype:
            return isinstance(value, (date, time))
        return DataType.DATE == dtype or DataType.TIME == dtype or DataType.DATE_TIME == dtype

    @staticmethod
    def is_number(ntype, value=None):
        if value and not ntype:
            return isinstance(value, numbers.Number)
        return DataType.NUMBER == ntype or DataType.CURRENCY == ntype or DataType.PERCENTAGE == ntype

    @staticmethod
    def is_bool(btype, value=None):
        if value and not btype:
            return isinstance(value, bool)
        return DataType.BOOL == btype

    @staticmethod
    def parse_type_by_value(type_value):
        return None if type_value is None else DataType(type_value)

    @staticmethod
    def parse_type(type_name):
        return None if type_name is None else DataType[type_name]


class DateFormat:

    @staticmethod
    def excel2python_format(dt_format):
        """
        https://docs.python.org/3/library/datetime.html#strftime-strptime-behavior
        """
        if not dt_format:
            return ''
        _format = dt_format
        _format = re.sub(r'(?:(?<!%))y{4}', '%Y', _format)
        _format = re.sub(r'(?:(?<!%))y{2}', '%y', _format)
        _format = re.sub(r'(?:(?<!%))m{4}', '%B', _format)
        _format = re.sub(r'(?:(?<!%))m{3}', '%b', _format)
        _format = re.sub(r'(?:(?<!%))m{2}', '%m', _format)
        _format = re.sub(r'(?:(?<!%))m{1}', '%-m', _format)
        _format = re.sub(r'(?:(?<!%))d{2}', '%d', _format)
        _format = re.sub(r'(?:(?<!%))d{1}', '%-d', _format)
        hour_code = 'I' if re.search(r'\b(AM/PM)|(A/P)\b', _format, re.RegexFlag.IGNORECASE) else 'H'
        _format = re.sub(r'\b(AM/PM)|(A/P)\b', '%p', _format)
        _format = re.sub(r'(?:(?<!%))H{2}', '%' + hour_code, _format)
        _format = re.sub(r'(?:(?<!%))H{1}', '%-' + hour_code, _format)
        _format = re.sub(r'(?:(?<!%))M{2}', '%M', _format)
        _format = re.sub(r'(?:(?<!%))M{1}', '%-M', _format)
        _format = re.sub(r'(?:(?<!%))S{2}', '%S', _format)
        _format = re.sub(r'(?:(?<!%))S{1}', '%-S', _format)
        return _format

    def __init__(self, date_format='yyyy-mm-dd', time_format='HH:MM:SS AM/PM', datetime_format='yyyy-mm-dd HH:MM'):
        self.date_format = date_format
        self.time_format = time_format
        self.datetime_format = datetime_format
        self.py_date_format = self.excel2python_format(self.date_format)
        self.py_time_format = self.excel2python_format(self.time_format)
        self.py_datetime_format = self.excel2python_format(self.datetime_format)
        self.datemode = 0  # 0: 1900-based, 1: 1904-based

    def get_excel_format(self, data_type, value):
        if value and isinstance(value, time) or DataType.TIME == data_type:
            return self.time_format.lower()
        if value and isinstance(value, datetime) or DataType.DATE_TIME == data_type:
            return self.datetime_format.lower()
        return self.date_format

    def get_py_format(self, data_type, value=None):
        if value and isinstance(value, time) or DataType.TIME == data_type:
            return self.py_time_format
        if value and isinstance(value, datetime) or DataType.DATE_TIME == data_type:
            return self.py_datetime_format
        return self.py_date_format

    def format(self, data_type, value):
        return value.strftime(self.get_py_format(data_type))

    def parse(self, data_type, value):
        return value if isinstance(value, (datetime, time, date)) else datetime.strptime(str(value), self.get_py_format(data_type, value))


class NumberFormat:

    def __init__(self, decimal_sep='.', thousand_sep=',', precision='2'):
        self.number_format = '{0:' + thousand_sep + decimal_sep + precision + 'f}'

    def get_excel_format(self, data_type):
        # TODO: Missing percentage and currency
        return self.number_format

    def format(self, data_type, value):
        # TODO: Missing percentage and currency
        return self.number_format.format(value)

    def parse(self, data_type, value):
        return value if isinstance(value, numbers.Number) else float(value)


class BoolFormat:

    def __init__(self, bool_format='Yes/No'):
        self.true = bool_format.split('/')[0]
        self.false = bool_format.split('/')[1]

    def format(self, value):
        return self.true if value else self.false

    def parse(self, value):
        return value if isinstance(value, bool) else value and str(value).lower() == self.true.lower()
