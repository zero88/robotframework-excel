#!/usr/bin/python

import logging
import six
from ExcelRobot.base import ExcelLibrary
from ExcelRobot.version import VERSION

_version_ = VERSION


class ExcelRobot(ExcelLibrary):
    """
    This test library provides some keywords to allow opening, reading, writing, and saving Excel files from Robot Framework.

    *Before running tests*

    Prior to running tests, ExcelRobot must first be imported into your Robot test suite.

    Example:
        | Library | ExcelRobot |
    """
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    ROBOT_LIBRARY_VERSION = VERSION

    def __init__(self, date_format='yyyy-mm-dd', time_format='hh:mm:ss AM/PM', datetime_format='yyyy-mm-dd hh:mm', decimal_sep='.', thousand_sep=','):
        logging.basicConfig()
        logging.getLogger().setLevel(logging.INFO)
        logger = logging.getLogger(__name__)
        logger.info('ExcelRobot::Robotframework Excel Library')
        super().__init__(date_format, time_format, datetime_format, decimal_sep, thousand_sep)
