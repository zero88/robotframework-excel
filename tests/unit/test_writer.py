#!/usr/bin/python

import os
import os.path as path
import shutil
import tempfile
from datetime import date, datetime, time

import pytest
from six import PY2

from ExcelRobot.reader import ExcelReader
from ExcelRobot.utils import DataType, copy_file, random_name
from ExcelRobot.writer import ExcelWriter

CURRENT_DIR = path.dirname(path.abspath(__file__))
DATA_DIR = path.join(CURRENT_DIR, '../data')
TEMP_DIR = path.join(tempfile.gettempdir(), 'ExcelRobot')


@pytest.fixture(autouse=True, scope='session')
def setup_module():
    try:
        os.makedirs(TEMP_DIR)
    except OSError as _:
        shutil.rmtree(TEMP_DIR)
        os.makedirs(TEMP_DIR)

    yield
    shutil.rmtree(TEMP_DIR)


@pytest.fixture
def create_tmp_file():
    open(path.join(TEMP_DIR, 'temp.xls'), mode='w+').close()
    yield
    os.remove(path.join(TEMP_DIR, 'temp.xls'))


def test_open_without_override(create_tmp_file):
    with pytest.raises(IOError if PY2 else FileExistsError):
        ExcelWriter(path.join(DATA_DIR, 'ExcelRobotTest.xls'), new_path=path.join(TEMP_DIR, 'temp.xls'), override=False)


@pytest.mark.parametrize('input_file, nb_of_sheets, sheet_name', [('1.xls', 1, 'Sheet'), ('2.xlsx', 1, 'Sheet')])
def test_open_new_file(input_file, nb_of_sheets, sheet_name):
    writer = ExcelWriter(path.join(TEMP_DIR, input_file))
    writer.save_excel()
    assert writer.get_number_of_sheets() == nb_of_sheets
    assert sheet_name in writer.get_sheet_names()


@pytest.mark.parametrize(
    'input_file, nb_of_sheets, sheet_name',
    [('ExcelRobotTest.xls', 5, 'TestSheet3'), ('ExcelRobotTest.xlsx', 5, 'TestSheet3')],
)
def test_open_existed_file(input_file, nb_of_sheets, sheet_name):
    new_file = path.join(TEMP_DIR, random_name() + '_' + input_file)
    writer = ExcelWriter(path.join(DATA_DIR, input_file), new_file)
    writer.save_excel()
    assert writer.get_number_of_sheets() == nb_of_sheets
    assert sheet_name in writer.get_sheet_names()


@pytest.mark.parametrize(
    'input_file, sheet_name', [('ExcelRobotTest.xls', 'TestSheet10'), ('ExcelRobotTest.xlsx', 'TestSheet10')]
)
def test_create_sheet_in_new_file(input_file, sheet_name):
    new_file = path.join(TEMP_DIR, random_name() + '_' + input_file)
    writer = ExcelWriter(path.join(DATA_DIR, input_file), new_file)
    writer.create_sheet(sheet_name)
    writer.save_excel()
    assert sheet_name in ExcelReader(new_file).get_sheet_names()


@pytest.mark.parametrize(
    'input_file, sheet_name', [('ExcelRobotTest.xls', 'TestSheet11'), ('ExcelRobotTest.xlsx', 'TestSheet11')]
)
def test_create_sheet_in_same_file(input_file, sheet_name):
    # TODO: Prepare data in tmp
    test_file = path.join(TEMP_DIR, random_name() + '_' + input_file)
    copy_file(path.join(DATA_DIR, input_file), test_file, True)
    writer = ExcelWriter(path.join(DATA_DIR, test_file))
    writer.create_sheet(sheet_name)
    writer.save_excel()
    assert sheet_name in ExcelReader(test_file).get_sheet_names()


def list_data():
    d1 = {'row': 0, 'column': 0, 'raw': 'Name', 'value': 'Name', 'type': DataType.TEXT.name}
    d2 = {'row': 0, 'column': 1, 'raw': 25.5, 'value': '25.50', 'type': DataType.NUMBER.name}
    d3 = {'row': 0, 'column': 2, 'raw': date(2018, 1, 1), 'value': '2018-01-01', 'type': DataType.DATE.name}
    d4 = {'row': 0, 'column': 3, 'raw': time(8, 0, 0), 'value': '08:00:00 AM', 'type': DataType.TIME.name}
    d5 = {
        'row': 0,
        'column': 4,
        'raw': datetime(2018, 1, 2, 22, 00),
        'value': '2018-01-02 22:00',
        'type': DataType.DATE_TIME.name,
    }
    d6 = {'row': 0, 'column': 5, 'raw': True, 'value': 'Yes', 'type': DataType.BOOL.name}
    return [d1, d2, d3, d4, d5, d6]


@pytest.mark.parametrize(
    'input_file, sheet_name, data',
    list(map(lambda x: ('ExcelRobotTest.xls', 'TestSheet20', x), list_data()))
    + list(map(lambda x: ('ExcelRobotTest.xlsx', 'TestSheet20', x), list_data())),
)
def test_write_raw_value_in_new_sheet(input_file, sheet_name, data):
    new_file = path.join(TEMP_DIR, random_name() + '_' + input_file)
    writer = ExcelWriter(path.join(DATA_DIR, input_file), new_file)
    writer.create_sheet(sheet_name)
    writer.write_to_cell(sheet_name, data['column'], data['row'], data['raw'])
    writer.save_excel()
    reader = ExcelReader(new_file)
    assert (
        reader.read_cell_data(sheet_name, data['column'], data['row'], data_type=data['type'], use_format=False)
        == data['raw']
    )
    assert (
        reader.read_cell_data(sheet_name, data['column'], data['row'], data_type=data['type'], use_format=True)
        == data['value']
    )


@pytest.mark.parametrize(
    'input_file, sheet_name, data',
    list(map(lambda x: ('ExcelRobotTest.xls', 'TestSheet30', x), list_data()))
    + list(map(lambda x: ('ExcelRobotTest.xlsx', 'TestSheet30', x), list_data())),
)
def test_write_format_value_in_new_sheet(input_file, sheet_name, data):
    new_file = path.join(TEMP_DIR, random_name() + '_' + input_file)
    writer = ExcelWriter(path.join(DATA_DIR, input_file), new_file)
    writer.create_sheet(sheet_name)
    writer.write_to_cell(sheet_name, data['column'], data['row'], data['value'], data_type=data['type'])
    writer.save_excel()
    reader = ExcelReader(new_file)
    assert (
        reader.read_cell_data(sheet_name, data['column'], data['row'], data_type=data['type'], use_format=False)
        == data['raw']
    )
    assert (
        reader.read_cell_data(sheet_name, data['column'], data['row'], data_type=data['type'], use_format=True)
        == data['value']
    )
