#!/usr/bin/python

import pytest

from ExcelRobot.utils import DateFormat


@pytest.mark.parametrize(
    'pattern,expected',
    [
        ('yyyy-mm-dd', '%Y-%m-%d'),
        ('yyyymmdd', '%Y%m%d'),
        ('dd/mm/yyyy', '%d/%m/%Y'),
        ('mmm, dd yyyy', '%b, %d %Y'),
        ('m/d/yy', '%-m/%-d/%y'),
        ('HH:MM:SS', '%H:%M:%S'),
        ('HH.MM.SS', '%H.%M.%S'),
        ('HH:MM:SS AM/PM', '%I:%M:%S %p'),
        ('HHMMSS-A/P', '%I%M%S-%p'),
        ('yyyy mm dd HH:MM:SS', '%Y %m %d %H:%M:%S'),
    ],
)
def test_date_format(pattern, expected):
    assert DateFormat.excel2python_format(pattern) == expected
