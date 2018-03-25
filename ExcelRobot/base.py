from ExcelRobot.reader import ExcelReader
from ExcelRobot.writer import ExcelWriter


class ExcelLibrary:

    def __init__(self, date_format='yyyy-mm-dd', time_format='hh:mm:ss AM/PM', datetime_format='yyyy-mm-dd hh:mm', decimal_sep='.', thousand_sep=','):
        self.date_format = date_format
        self.time_format = time_format
        self.datetime_format = datetime_format
        self.decimal_sep = decimal_sep
        self.thousand_sep = thousand_sep
        self.reader = None
        self.writer = None

    def open_excel(self, file_path):
        """
        Opens the Excel file to read from the path provided in the file path parameter.

        Arguments:
                |  File Path (string) | The Excel file name or path will be opened. If file name then openning file in current directory.   |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """
        self.reader = ExcelReader(file_path,
                                  self.date_format, self.time_format,
                                  self.datetime_format, self.decimal_sep, self.thousand_sep)

    def open_excel_to_write(self, file_path, new_path=None, override=False):
        """
        Opens the Excel file to write from the path provided in the file name parameter.
        In case `New Path` is given, new file will be created based on content of current file.

        Arguments:
                |  File Path (string)           | The Excel file name or path will be opened. If file name then openning file in current directory.   |
                |  New Path                     | New path will be saved.                                                                             |
                |  Override (Default: `False`)  | If `True`, new file will be overriden if it exists.                                                 |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """
        self.writer = ExcelWriter(file_path, new_path, override,
                                  self.date_format, self.time_format,
                                  self.datetime_format, self.decimal_sep, self.thousand_sep)

    def get_sheet_names(self):
        """
        Returns the names of all the worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheets Names        |                                                    |

        """
        return self.reader.get_sheet_names()

    def get_number_of_sheets(self):
        """
        Returns the number of worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Number of Sheets    |                                                    |

        """
        return self.reader.get_number_of_sheets()

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
        return self.reader.get_column_count(sheet_name)

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
        return self.reader.get_row_count(sheet_name)

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
        return self.reader.get_column_values(sheet_name, column, include_empty_cells)

    def get_row_values(self, sheet_name, row, include_empty_cells=True):
        """
        Returns the specific row values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                      | The selected sheet that the row values will be returned from.         |
                |  Row (int)                                | The row integer value value is indicated to get values.               |
                |  Include Empty Cells (Default: `True`)    |  If `False` then only return cells with values.                       |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Row Values       |  TestSheet1                                        | 0 |

        """
        return self.reader.get_row_values(sheet_name, row, include_empty_cells)

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
        return self.reader.get_sheet_values(sheet_name, include_empty_cells)

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
        return self.reader.get_workbook_values(include_empty_cells)

    def read_cell_data_by_name(self, sheet_name, cell_name, data_format=None):
        """
        Uses the cell name to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.  |
                |  Cell Name (string)   | The selected cell name that the value will be returned from.   |
        Example:

        | *Keywords*                |  *Parameters*                                             |
        | Open Excel                |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |      |
        | Read Cell Data By Name    |  TestSheet1                                        |  A2  |

        """
        return self.reader.read_cell_data_by_name(sheet_name, cell_name, data_format)

    def read_cell_data_by_coordinates(self, sheet_name, column, row, data_format=None):
        """
        Uses the column and row to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.         |
                |  Column (int)         | The column integer value that the cell value will be returned from.   |
                |  Row (int)            | The row integer value that the cell value will be returned from.      |
        Example:

        | *Keywords*                        |  *Parameters*                                              |
        | Open Excel                        |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Read Cell Data By Coordinates     |  TestSheet1                                        | 0 | 0 |

        """
        return self.reader.read_cell_data_by_coordinates(sheet_name, column, row, data_format)

    def check_cell_type(self, sheet_name, column, row, data_type):
        """
        Checks the type of value that is within the cell of the sheet name selected.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell type will be checked from.                                   |
                |  Column (int)         | The column integer value that will be used to check the cell type.                            |
                |  Row (int)            | The row integer value that will be used to check the cell type.                               |
                |  Data Type (string)   | Data type to check. Available options: `DATE`, `TEXT`, `NUMBER`, `BOOL`, `EMPTY`, `ERROR`     |
        Example:

        | *Keywords*           |  *Parameters*                                              |       |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |       |
        | Check Cell Type      |  TestSheet1                                        | 0 | 0 | DATE  |

        """
        return self.reader.check_cell_type(sheet_name, column, row, data_type)

    def write_number_to_cell(self, sheet_name, column, row, value, decimal_sep=None, thousand_sep=None):
        """
        Using the sheet name the value of the indicated cell is set to be the number given in the parameter.

        Arguments:
                |  Sheet Name (string)          | The selected sheet that the cell will be modified from.        |
                |  Column (int)                 | The column integer value that will be used to modify the cell. |
                |  Row (int)                    | The row integer value that will be used to modify the cell.    |
                |  Value (int)                  | The integer value that will be added.                          |
                |  Decimal Separator (string)   | Overide decimal separtor in global scope.                      |
                |  Thousand Separator (string)  | Overide thousand separtor in global scope.                     |
        Example:

        | *Keywords*           |  *Parameters*                                                         |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |      |
        | Write Number To Cell |  TestSheet1                                        |  0  |  0  |  34  |

        """
        self.writer.write_number_to_cell(sheet_name, column, row, value, decimal_sep, thousand_sep)

    def write_string_to_cell(self, sheet_name, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the string given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.        |
                |  Column (int)        | The column integer value that will be used to modify the cell. |
                |  Row (int)           | The row integer value that will be used to modify the cell.    |
                |  Value (string)      | The string value that will be added.                           |
        Example:

        | *Keywords*            |  *Parameters*                                                           |
        | Open Excel            |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |        |
        | Write String To Cell  |  TestSheet1                                        |  0  |  0  |  Hello |

        """
        self.writer.write_string_to_cell(sheet_name, column, row, value)

    def write_date_to_cell(self, sheet_name, column, row, value, date_format=None):
        """
        Using the sheet name the value of the indicated cell is set to be the date given in the parameter.

        Arguments:
                |  Sheet Name (string)              | The selected sheet that the cell will be modified from.           |
                |  Column (int)                     | The column integer value that will be used to modify the cell.    |
                |  Row (int)                        | The row integer value that will be used to modify the cell.       |
                |  Value (int)                      | The integer value containing a date that will be added.           |
                |  Date Format (string)             | Overide decimal separtor in global scope.                         |
        Example:

        | *Keywords*            |  *Parameters*                                                               |
        | Open Excel            |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |            |
        | Write Date To Cell    |  TestSheet1                                        |  0  |  0  |  12.3.1999 |

        """
        self.writer.write_date_to_cell(sheet_name, column, row, value, date_format)

    def modify_cell_with(self, sheet_name, column, row, op, val):
        """
        Using the sheet name a cell is modified with the given operation and value.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell will be modified from.                                                  |
                |  Column (int)         | The column integer value that will be used to modify the cell.                                           |
                |  Row (int)            | The row integer value that will be used to modify the cell.                                              |
                |  Operation (operator) | The operation that will be performed on the value within the cell located by the column and row values.  |
                |  Value (int)          | The integer value that will be used in conjuction with the operation parameter.                          |
        Example:

        | *Keywords*           |  *Parameters*                                                               |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |      |
        | Modify Cell With     |  TestSheet1                                        |  0  |  0  |  *  |  56  |

        """
        self.writer.modify_cell_with(sheet_name, column, row, op, val)

    def add_to_date(self, sheet_name, column, row, numdays):
        """
        Using the sheet name the number of days are added to the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.            |
                |  Column (int)                    | The column integer value that will be used to modify the cell.     |
                |  Row (int)                       | The row integer value that will be used to modify the cell.        |
                |  Number of Days (int)            | The integer value is the number of days that will be added.        |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Add To Date          |  TestSheet1                                        |  0  |  0  |  4  |

        """
        self.writer.add_to_date(sheet_name, column, row, numdays)

    def subtract_from_date(self, sheet_name, column, row, numdays):
        """
        Using the sheet name the number of days are subtracted from the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.            |
                |  Column (int)                    | The column integer value that will be used to modify the cell.     |
                |  Row (int)                       | The row integer value that will be used to modify the cell.        |
                |  Number of Days (int)            | The integer value is the number of days that will be added.        |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Subtract From Date   |  TestSheet1                                        |  0  |  0  |  7  |

        """
        self.writer.subtract_from_date(sheet_name, column, row, numdays)

    def save_excel(self):
        """
        Saves the Excel file that was opened to write before.

        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel To Write  |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Save Excel           |                                                    |

        """
        self.writer.save_excel()

    def create_sheet(self, sheet_name):
        """
        Creates and appends new Excel worksheet using the new sheet name to the current workbook.

        Arguments:
                |  New Sheet name (string)  | The name of the new sheet added to the workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Create Sheet         |  NewSheet                                          |

        """
        self.writer.create_sheet(sheet_name)

    def create_workbook(self, sheet_name):
        """
        Creates a new Excel workbook

        Arguments:
                |  New Sheet Name (string)  | The name of the new sheet added to the new workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Create Excel         |  NewExcelSheet                                     |

        """
        self.writer.create_workbook(sheet_name)
