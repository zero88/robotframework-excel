import re
from enum import Enum
from xlrd import (XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
                  XL_CELL_ERROR, XL_CELL_NUMBER, XL_CELL_TEXT)


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
    def is_date(dtype):
        return DataType.DATE == dtype or DataType.TIME == dtype or DataType.DATE_TIME == dtype

    @staticmethod
    def is_number(ntype):
        return DataType.NUMBER == ntype or DataType.CURRENCY == ntype or DataType.PERCENTAGE == ntype

    @staticmethod
    def is_bool(btype):
        return DataType.BOOL == btype

    @staticmethod
    def parse_type(type_name):
        if not type_name:
            return DataType.TEXT
        excel_type = DataType[type_name]
        if excel_type is None:
            raise AttributeError('Not support %s' % type_name)
        return excel_type


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
        self.date_format = self.excel2python_format(date_format)
        self.time_format = self.excel2python_format(time_format)
        self.datetime_format = self.excel2python_format(datetime_format)
        self.datemode = 0  # 0: 1900-based, 1: 1904-based

    def parse(self, data_type):
        if DataType.TIME == data_type:
            return self.time_format
        if DataType.DATE_TIME == data_type:
            return self.datetime_format
        return self.date_format


class NumberFormat:

    def __init__(self, decimal_sep='.', thousand_sep=',', precision='2'):
        self.number_format = '{0:' + thousand_sep + decimal_sep + precision + 'f}'

    def format(self, value):
        return self.number_format.format(value)


class BoolFormat:

    def __init__(self, bool_format='Yes/No'):
        self.true = bool_format.split('/')[0]
        self.false = bool_format.split('/')[1]

    def format(self, value):
        return self.true if value else self.false
