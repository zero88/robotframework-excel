#!/usr/bin/python

from ExcelRobot.utils import DateFormat
from nose.tools import eq_
from parameterized import parameterized


@parameterized([
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
])
def test_date_format(pattern, expected):
    eq_(expected, DateFormat.excel2python_format(pattern))
