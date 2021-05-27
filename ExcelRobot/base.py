from ExcelRobot.reader import ExcelReader
from ExcelRobot.utils import BoolFormat, DateFormat, NumberFormat
from ExcelRobot.writer import ExcelWriter


class ExcelLibrary(object):

    def __init__(self, date_format=DateFormat(), number_format=NumberFormat(), bool_format=BoolFormat()):
        """
        Init Excel Keyword with some default configuration.

        Excel Date Time format
        https://support.office.com/en-us/article/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309
        """
        self.date_format = date_format
        self.number_format = number_format
        self.bool_format = bool_format
        self.active = None

    def open_excel(self, file_path):
        """
        Opens the Excel file to read from the path provided in the file path parameter.

        Arguments:
                |  File Path (string) | The Excel file name or path will be opened. If file name then openning file in current directory.   |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """
        self.active = ExcelReader(file_path, self.date_format, self.number_format, self.bool_format)

    def open_excel_to_write(self, file_path, new_path=None, override=False):
        """
        Opens the Excel file to write from the path provided in the file name parameter.
        In case `New Path` is given, new file will be created based on content of current file.

        Arguments:
                |  File Path (string)           | The Excel file name or path will be opened. If file name then openning file in current directory. |
                |  New Path                     | New path will be saved.                                                                           |
                |  Override (Default: `False`)  | If `True`, new file will be overriden if it exists.                                               |
        Example:

        | *Keywords*                |  *Parameters*                                      |
        | Open Excel To Write       |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """
        self.active = ExcelWriter(file_path, new_path, override, self.date_format, self.number_format, self.bool_format)

    def get_sheet_names(self):
        """
        Returns the names of all the worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheets Names        |                                                    |

        """
        return self.active.get_sheet_names()

    def get_number_of_sheets(self):
        """
        Returns the number of worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Number of Sheets    |                                                    |

        """
        return self.active.get_number_of_sheets()

    def get_column_count(self, sheet_name):
        """
        Returns the specific number of columns of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the column count will be returned from. |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Column Count    |  TestSheet1                                        |

        """
        return self.active.get_column_count(sheet_name)

    def get_row_count(self, sheet_name):
        """
        Returns the specific number of rows of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the row count will be returned from. |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Row Count       |  TestSheet1                                        |

        """
        return self.active.get_row_count(sheet_name)

    def get_column_values(self, sheet_name, column, include_empty_cells=True):
        """
        Returns the specific column values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the column values will be returned from.   |
                |  Column (int)                             | The column integer value is indicated to get values.               |
                |  Include Empty Cells (Default: `True`)    | If `False` then only return cells with values.                     |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Column Values    |  TestSheet1                                        | 0 |

        """
        return self.active.get_column_values(sheet_name, column, include_empty_cells)

    def get_row_values(self, sheet_name, row, include_empty_cells=True):
        """
        Returns the specific row values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the row values will be returned from.         |
                |  Row (int)                                | The row integer value value is indicated to get values.               |
                |  Include Empty Cells (Default: `True`)    | If `False` then only return cells with values.                        |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Row Values       |  TestSheet1                                        | 0 |

        """
        return self.active.get_row_values(sheet_name, row, include_empty_cells)

    def get_sheet_values(self, sheet_name, include_empty_cells=True):
        """
        Returns the values from the sheet name specified.

        Arguments:
                |  Sheet Name (string                       | The selected sheet that the cell values will be returned from.    |
                |  Include Empty Cells (Default: `True`)    | If `False` then only return cells with values.                    |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheet Values     |  TestSheet1                                        |

        """
        return self.active.get_sheet_values(sheet_name, include_empty_cells)

    def get_workbook_values(self, include_empty_cells=True):
        """
        Returns the values from each sheet of the current workbook.

        Arguments:
                |  Include Empty Cells (Default: `True`)    | If `False` then only return cells with values.                    |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Workbook Values  |                                                    |

        """
        return self.active.get_workbook_values(include_empty_cells)

    def read_cell_data_by_name(self, sheet_name, cell_name, data_type=None, use_format=True):
        """
        Uses the cell name to return the data from that cell.

        - `Data Type` indicates explicit data type to convert cell value to correct data type.
        - `Use Format` is False, then cell value will be raw data with correct data type.

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the cell value will be returned from.             |
                |  Cell Name (string)                       | The selected cell name that the value will be returned from.              |
                |  Data Type (string)                       | Available options: `TEXT`, DATE`, `TIME`, `DATETIME`, `NUMBER`, `BOOL`    |
                |  Use Format (boolean) (Default: `True`)   | Use format to convert data to string.                                     |
        Example:

        | *Keywords*                |  *Parameters*                                             |
        | Open Excel                |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |      |
        | Read Cell Data By Name    |  TestSheet1                                        |  A2  |

        """
        return self.active.read_cell_data_by_name(sheet_name, cell_name, data_type, use_format)

    def read_cell_data(self, sheet_name, column, row, data_type=None, use_format=True):
        """
        Uses the column and row to return the data from that cell.

        - `Data Type` indicates explicit data type to convert cell value to correct data type.
        - `Use Format` is False, then cell value will be raw data with correct data type.

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the cell value will be returned from.             |
                |  Column (int)                             | The column integer value that the cell value will be returned from.       |
                |  Row (int)                                | The row integer value that the cell value will be returned from.          |
                |  Data Type (string)                       | Available options: `TEXT`, DATE`, `TIME`, `DATETIME`, `NUMBER`, `BOOL`    |
                |  Use Format (boolean) (Default: `True`)   | Use format to convert data to string.                                     |
        Example:

        | *Keywords*        |  *Parameters*                                              |
        | Open Excel        |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Read Cell Data    |  TestSheet1                                        | 0 | 0 |

        """
        return self.active.read_cell_data(sheet_name, column, row, data_type, use_format)

    def check_cell_type(self, sheet_name, column, row, data_type):
        """
        Checks the type of value that is within the cell of the sheet name selected.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell type will be checked from.                                   |
                |  Column (int)         | The column integer value that will be used to check the cell type.                            |
                |  Row (int)            | The row integer value that will be used to check the cell type.                               |
                |  Data Type (string)   | Available options: `DATE`, `TIME`, `DATE_TIME`, `TEXT`, `NUMBER`, `BOOL`, `EMPTY`, `ERROR`    |
        Example:

        | *Keywords*           |  *Parameters*                                              |       |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |       |
        | Check Cell Type      |  TestSheet1                                        | 0 | 0 | DATE  |

        """
        return self.active.check_cell_type(sheet_name, column, row, data_type)

    def write_to_cell_by_name(self, sheet_name, cell_name, value, data_type=None):
        """
        Write data to cell by using the given sheet name and the given cell that defines by name.

        If `Data Type` is not provided, `ExcelRobot` will introspect data type from given `value` to define cell type

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the cell will be modified from.                       |
                |  Cell Name (string)                       | The selected cell name that the value will be returned from.                  |
                |  Value (string|number|datetime|boolean)   | Raw value or string value then using DataType to decide data type to write    |
                |  Data Type (string)                       | Available options: `DATE`, `TIME`, `DATE_TIME`, `TEXT`, `NUMBER`, `BOOL`      |
        Example:

        | *Keywords*            |  *Parameters*                                                                     |
        | Open Excel To Write   |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |                      |       |
        | Write To Cell By Name |  TestSheet1                                        |  A1  |  34           |       |
        | Write To Cell By Name |  TestSheet1                                        |  A2  |  2018-03-29   | DATE  |
        | Write To Cell By Name |  TestSheet1                                        |  A3  |  YES          | BOOL  |

        """
        self.active.write_to_cell_by_name(sheet_name, cell_name, value, data_type)

    def write_to_cell(self, sheet_name, column, row, value, data_type=None):
        """
        Write data to cell by using the given sheet name and the given cell that defines by column and row.

        If `Data Type` is not provided, `ExcelRobot` will introspect data type from given `value` to define cell type

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the cell will be modified from.                       |
                |  Column (int)                             | The column integer value that will be used to modify the cell.                |
                |  Row (int)                                | The row integer value that will be used to modify the cell.                   |
                |  Value (string|number|datetime|boolean)   | Raw value or string value then using DataType to decide data type to write    |
                |  Data Type (string)                       | Available options: `DATE`, `TIME`, `DATE_TIME`, `TEXT`, `NUMBER`, `BOOL`      |
        Example:

        | *Keywords*            |  *Parameters*                                                                 |
        | Open Excel To Write   |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |              |       |
        | Write To Cell         |  TestSheet1                                        |  0  |  0  |  34          |       |
        | Write To Cell         |  TestSheet1                                        |  1  |  1  |  2018-03-29  | DATE  |
        | Write To Cell         |  TestSheet1                                        |  2  |  2  |  YES         | BOOL  |

        """
        self.active.write_to_cell(sheet_name, column, row, value, data_type)

    # def modify_cell_with(self, sheet_name, column, row, op, val):
    #     """
    #     Using the sheet name a cell is modified with the given operation and value.

    #     Arguments:
    #             |  Sheet Name (string)  | The selected sheet that the cell will be modified from.                                                  |
    #             |  Column (int)         | The column integer value that will be used to modify the cell.                                           |
    #             |  Row (int)            | The row integer value that will be used to modify the cell.                                              |
    #             |  Operation (operator) | The operation that will be performed on the value within the cell located by the column and row values.  |
    #             |  Value (int)          | The integer value that will be used in conjuction with the operation parameter.                          |
    #     Example:

    #     | *Keywords*           |  *Parameters*                                                               |
    #     | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |      |
    #     | Modify Cell With     |  TestSheet1                                        |  0  |  0  |  *  |  56  |

    #     """
    #     self.active.modify_cell_with(sheet_name, column, row, op, val)

    def save_excel(self):
        """
        Saves the Excel file that was opened to write before.

        Example:

        | *Keywords*            |  *Parameters*                                      |
        | Open Excel To Write   |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |                  |
        | Write To Cell         |  TestSheet1                                        |  0  |  0  |  34  |
        | Save Excel            |                                                    |                  |

        """
        self.active.save_excel()

    def create_sheet(self, sheet_name):
        """
        Creates and appends new Excel worksheet using the new sheet name to the current workbook.

        Arguments:
                |  New Sheet name (string)  | The name of the new sheet added to the workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel To Write  |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Create Sheet         |  NewSheet                                          |

        """
        self.active.create_sheet(sheet_name)

    def remove_sheet(self, sheet_name):
        """
        Removes Excel worksheet by name.
        Arguments:
                |  Sheet name (string)  | The name of the Sheet    |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xlsx |
        | Remove Sheet         |  TestSheet1                                          |

        """
        self.active.remove_sheet(sheet_name)