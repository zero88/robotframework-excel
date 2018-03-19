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
