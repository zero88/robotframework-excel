#!/usr/bin/python
import os.path as path
from datetime import datetime

import pytest
from six import PY2

from ExcelRobot.reader import ExcelReader
from ExcelRobot.utils import DataType

CURRENT_DIR = path.dirname(path.abspath(__file__))
DATA_DIR = path.join(CURRENT_DIR, '../data')


def test_open_not_valid():
    with pytest.raises(ValueError):
        ExcelReader(path.join(DATA_DIR, 'a.txt'))


def test_open_not_found():
    with pytest.raises(IOError if PY2 else FileNotFoundError):
        ExcelReader(path.join(DATA_DIR, 'a.xls'))


@pytest.mark.parametrize('input_file,expected', [('ExcelRobotTest.xls', 5), ('ExcelRobotTest.xlsx', 5)])
def test_open_success(input_file, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.get_number_of_sheets() == expected


@pytest.mark.parametrize(
    'input_file,expected',
    [('ExcelRobotTest.xls', 'TestSheet1'), ('ExcelRobotTest.xlsx', 'TestSheet1')],
)
def test_sheet_name(input_file, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert expected in reader.get_sheet_names()


@pytest.mark.parametrize(
    'input_file, sheet_name, col_count, row_count',
    [
        ('ExcelRobotTest.xls', 'TestSheet1', 2, 3),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 2, 3),
    ],
)
def test_sheet_size(input_file, sheet_name, col_count, row_count):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.get_column_count(sheet_name) == col_count
    assert reader.get_row_count(sheet_name) == row_count


@pytest.mark.parametrize(
    'input_file, sheet_name, column, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet1', 0, [('A1', 'This is a test sheet'), ('A2', 'User1'), ('A3', 'User2')]),
        ('ExcelRobotTest.xls', 'TestSheet1', 1, [('B1', 'Points'), ('B2', 57), ('B3', 5178)]),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 0, [('A1', 'This is a test sheet'), ('A2', 'User1'), ('A3', 'User2')]),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 1, [('B1', 'Points'), ('B2', 57), ('B3', 5178)]),
    ],
)
def test_get_col_values(input_file, sheet_name, column, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.get_column_values(sheet_name, column) == expected


@pytest.mark.parametrize(
    'input_file, sheet_name, column, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet2', 0, [('A1', 'This is a test sheet'), ('B1', 'Date of Birth')]),
        ('ExcelRobotTest.xls', 'TestSheet2', 1, [('A2', 'User3'), ('B2', '23.8.1982')]),
        ('ExcelRobotTest.xlsx', 'TestSheet2', 0, [('A1', 'This is a test sheet'), ('B1', 'Date of Birth')]),
        ('ExcelRobotTest.xlsx', 'TestSheet2', 1, [('A2', 'User3'), ('B2', '23.8.1982')]),
    ],
)
def test_get_row_values(input_file, sheet_name, column, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.get_row_values(sheet_name, column) == expected


@pytest.mark.parametrize(
    'input_file, sheet_name, cell_name, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet1', 'a2', 'User1'),
        ('ExcelRobotTest.xls', 'TestSheet1', 'B2', '57.00'),
        ('ExcelRobotTest.xls', 'TestSheet2', 'B2', '23.8.1982'),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 'A2', 'User1'),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 'B2', '57.00'),
        ('ExcelRobotTest.xlsx', 'TestSheet2', 'B2', '23.8.1982'),
    ],
)
def test_get_cell_value_by_name(input_file, sheet_name, cell_name, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.read_cell_data_by_name(sheet_name, cell_name) == expected


@pytest.mark.parametrize(
    'input_file, sheet_name, col, row, data_type, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet1', 0, 1, None, 'User1'),
        ('ExcelRobotTest.xls', 'TestSheet1', 1, 1, DataType.NUMBER.name, '57.00'),
        ('ExcelRobotTest.xls', 'TestSheet2', 1, 1, DataType.TEXT.name, '23.8.1982'),
        ('ExcelRobotTest.xls', 'TestSheet3', 2, 1, DataType.DATE.name, '1982-05-14'),
        ('ExcelRobotTest.xls', 'TestSheet3', 3, 1, DataType.BOOL.name, 'Yes'),
        ('ExcelRobotTest.xls', 'TestSheet3', 6, 1, DataType.NUMBER.name, '7.50'),
        ('ExcelRobotTest.xls', 'TestSheet3', 7, 1, DataType.TIME.name, '08:00:00 AM'),
        ('ExcelRobotTest.xls', 'TestSheet3', 8, 1, DataType.DATE_TIME.name, '2018-01-02 22:00'),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 0, 1, None, 'User1'),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 1, 1, DataType.NUMBER.name, '57.00'),
        ('ExcelRobotTest.xlsx', 'TestSheet2', 1, 1, DataType.TEXT.name, '23.8.1982'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 2, 1, DataType.DATE.name, '1982-05-14'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 3, 1, DataType.BOOL.name, 'Yes'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 6, 1, DataType.NUMBER.name, '7.50'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 7, 1, DataType.TIME.name, '08:00:00 AM'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 8, 1, DataType.DATE_TIME.name, '2018-01-02 22:00'),
    ],
)
def test_get_cell_value_by_coord(input_file, sheet_name, col, row, data_type, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.read_cell_data(sheet_name, col, row, data_type=data_type) == expected


@pytest.mark.parametrize(
    'input_file, sheet_name, col, row, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet1', 0, 1, 'User1'),
        ('ExcelRobotTest.xls', 'TestSheet1', 1, 1, 57),
        ('ExcelRobotTest.xls', 'TestSheet2', 1, 1, '23.8.1982'),
        ('ExcelRobotTest.xls', 'TestSheet3', 2, 1, datetime(1982, 5, 14)),
        ('ExcelRobotTest.xls', 'TestSheet3', 3, 1, True),
        ('ExcelRobotTest.xls', 'TestSheet3', 6, 1, 7.5),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 0, 1, 'User1'),
        ('ExcelRobotTest.xlsx', 'TestSheet1', 1, 1, 57),
        ('ExcelRobotTest.xlsx', 'TestSheet2', 1, 1, '23.8.1982'),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 2, 1, datetime(1982, 5, 14)),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 3, 1, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 6, 1, 7.5),
    ],
)
def test_get_cell_raw_value_by_coord(input_file, sheet_name, col, row, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.read_cell_data(sheet_name, col, row, use_format=False) == expected


@pytest.mark.parametrize(
    'input_file, sheet_name, col, row, data_type, expected',
    [
        ('ExcelRobotTest.xls', 'TestSheet3', 0, 1, DataType.NUMBER.name, True),
        ('ExcelRobotTest.xls', 'TestSheet3', 1, 1, DataType.TEXT.name, True),
        ('ExcelRobotTest.xls', 'TestSheet3', 2, 1, DataType.DATE.name, True),
        ('ExcelRobotTest.xls', 'TestSheet3', 3, 1, DataType.BOOL.name, True),
        ('ExcelRobotTest.xls', 'TestSheet3', 5, 1, DataType.EMPTY.name, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 0, 1, DataType.NUMBER.name, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 1, 1, DataType.TEXT.name, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 2, 1, DataType.DATE.name, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 3, 1, DataType.BOOL.name, True),
        ('ExcelRobotTest.xlsx', 'TestSheet3', 5, 1, DataType.EMPTY.name, True),
    ],
)
def test_check_cell_type(input_file, sheet_name, col, row, data_type, expected):
    reader = ExcelReader(path.join(DATA_DIR, input_file))
    assert reader.check_cell_type(sheet_name, col, row, data_type) == expected
