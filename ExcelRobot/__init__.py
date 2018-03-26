#!/usr/bin/python

import logging
import six
from ExcelRobot.base import ExcelLibrary
from ExcelRobot.version import VERSION
from ExcelRobot.utils import DateFormat, NumberFormat, BoolFormat

_version_ = VERSION


class ExcelRobot(ExcelLibrary):
    """
    This test library provides some keywords to allow opening, reading, writing, and saving Excel files from Robot Framework.

    *Before running tests*

    Prior to running tests, ExcelRobot must first be imported into your Robot test suite.

    Example:
        | Library | ExcelRobot |

    To setup some Excel configurations related to date format and number format, use these arguments
        *Excel Date Time format*
        | Date Format       | Default: `yyyy-mm-dd`         |
        | Time Format       | Default: `HH:MM:SS AM/PM`     |
        | Date Time Format  | Default: `yyyy-mm-dd HH:MM`   |
        For more information, check this article
        https://support.office.com/en-us/article/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309

        *Excel Number format*
        | Decimal Separator     | Default: `.`  |
        | Thousand Separator    | Default: `,`  |
        | Precision             | Default: `2`  |

        *Excel Boolean format*
        | Boolean Format        | Default: `Yes/No`  |

    Example:
        | Library | ExcelRobot | date_format='dd/mm/yyyy'
    """
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    ROBOT_LIBRARY_VERSION = VERSION

    def __init__(self,
                 date_format='yyyy-mm-dd', time_format='HH:MM:SS AM/PM', datetime_format='yyyy-mm-dd HH:MM',
                 decimal_sep='.', thousand_sep=',', precision='2', bool_format='Yes/No'):
        logging.basicConfig()
        logging.getLogger().setLevel(logging.INFO)
        logger = logging.getLogger(__name__)
        logger.info('ExcelRobot::Robotframework Excel Library')
        # super().__init__(date_format, time_format, datetime_format, decimal_sep, thousand_sep, precision)
        super().__init__(
            DateFormat(date_format, time_format, datetime_format),
            NumberFormat(decimal_sep, thousand_sep, precision),
            BoolFormat(bool_format))
