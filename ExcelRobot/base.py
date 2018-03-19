import os
import tempfile
from datetime import datetime, timedelta
from operator import itemgetter
from os import path

import natsort
from xlrd import (XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_EMPTY,
                  XL_CELL_ERROR, XL_CELL_NUMBER, XL_CELL_TEXT, cellname,
                  error_text_from_code, open_workbook, xldate_as_tuple)
from xlutils.copy import copy as copy
from xlwt import Workbook, easyxf


class ExcelLibrary:

    def __init__(self):
        self.workbook = None
        self.copied_workbook = None
        self.sheet_num = None
        self.sheet_names = None
        self.tmp_dir = tempfile.gettempdir()

    def __get_file_path(self, file_name, use_temp_dir):
        if not file_name.endswith('.xlsx') and not file_name.endswith('.xls'):
            raise ValueError('Only support file with extenstion: xls and xlsx')
        file_path = path.normpath(file_name)
        is_file = not len(path.splitdrive(file_path)) > 1
        return path.join(self.tmp_dir, file_name) if use_temp_dir else file_name if is_file else path.join(os.getcwd(), file_name)

    def open_excel(self, filename, use_temp_dir=False):
        """
        Opens the Excel file from the path provided in the file name parameter.
        If the boolean `use_temp_dir` is set to `true`, depending on the operating system of the computer running the test the file will be opened in the Temp directory if the operating system is Windows or tmp directory if it is not.

        Arguments:
                |  File Name (string)                      | The file name or path value that will be used to open the excel file to perform tests upon. If file name then application will open file in current directory. |
                |  Use Temporary Directory (default=False) | The file will not open in a temporary directory by default. To activate and open the file in a temporary directory, pass 'True' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """
        file_path = self.__get_file_path(filename, use_temp_dir)
        print('Opening file at %s' % file_path)
        use_format = not filename.endswith('.xlsx')
        self.workbook = open_workbook(file_path, formatting_info=use_format, on_demand=True)
        self.sheet_names = self.workbook.sheet_names()

    def get_sheet_names(self):
        """
        Returns the names of all the worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheets Names        |                                                    |

        """
        sheet_names = self.workbook.sheet_names()
        return sheet_names

    def get_number_of_sheets(self):
        """
        Returns the number of worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Number of Sheets    |                                                    |

        """
        sheet_num = self.workbook.nsheets
        return sheet_num

    def get_column_count(self, sheetname):
        """
        Returns the specific number of columns of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the column count will be returned from. |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Column Count    |  TestSheet1                                        |

        """
        sheet = self.workbook.sheet_by_name(sheetname)
        return sheet.ncols

    def get_row_count(self, sheetname):
        """
        Returns the specific number of rows of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the row count will be returned from. |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Row Count       |  TestSheet1                                        |

        """
        sheet = self.workbook.sheet_by_name(sheetname)
        return sheet.nrows

    def get_column_values(self, sheetname, column, include_empty_cells=True):
        """
        Returns the specific column values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the column values will be returned from.                                                            |
                |  Column (int)                        | The column integer value that will be used to select the column from which the values will be returned.                     |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Column Values    |  TestSheet1                                        | 0 |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        data = {}
        for row_index in range(sheet.nrows):
            cell = cellname(row_index, int(column))
            value = sheet.cell(row_index, int(column)).value
            data[cell] = value
        if include_empty_cells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_row_values(self, sheetname, row, include_empty_cells=True):
        """
        Returns the specific row values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the row values will be returned from.                                                               |
                |  Row (int)                           | The row integer value that will be used to select the row from which the values will be returned.                           |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Row Values       |  TestSheet1                                        | 0 |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        data = {}
        for col_index in range(sheet.ncols):
            cell = cellname(int(row), col_index)
            value = sheet.cell(int(row), col_index).value
            data[cell] = value
        if include_empty_cells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_sheet_values(self, sheetname, include_empty_cells=True):
        """
        Returns the values from the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the cell values will be returned from.                                                              |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheet Values     |  TestSheet1                                        |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        data = {}
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = cellname(row_index, col_index)
                value = sheet.cell(row_index, col_index).value
                data[cell] = value
        if include_empty_cells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_workbook_values(self, include_empty_cells=True):
        """
        Returns the values from each sheet of the current workbook.

        Arguments:
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Workbook Values  |                                                    |

        """
        sheet_data = []
        workbook_data = []
        for sheet_name in self.sheet_names:
            if include_empty_cells is True:
                sheet_data = self.get_sheet_values(sheet_name)
            else:
                sheet_data = self.get_sheet_values(sheet_name, False)
            sheet_data.insert(0, sheet_name)
            workbook_data.append(sheet_data)
        return workbook_data

    def read_cell_data_by_name(self, sheetname, cell_name):
        """
        Uses the cell name to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.  |
                |  Cell Name (string)   | The selected cell name that the value will be returned from.   |
        Example:

        | *Keywords*           |  *Parameters*                                             |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |      |
        | Get Cell Data        |  TestSheet1                                        |  A2  |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = cellname(row_index, col_index)
                if cell_name == cell:
                    return sheet.cell(row_index, col_index).value
        return ""

    def read_cell_data_by_coordinates(self, sheetname, column, row):
        """
        Uses the column and row to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.         |
                |  Column (int)         | The column integer value that the cell value will be returned from.   |
                |  Row (int)            | The row integer value that the cell value will be returned from.      |
        Example:

        | *Keywords*     |  *Parameters*                                              |
        | Open Excel     |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Read Cell      |  TestSheet1                                        | 0 | 0 |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        return sheet.cell(int(row), int(column)).value

    def check_cell_type(self, sheetname, column, row):
        """
        Checks the type of value that is within the cell of the sheet name selected.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell type will be checked from.          |
                |  Column (int)         | The column integer value that will be used to check the cell type.   |
                |  Row (int)            | The row integer value that will be used to check the cell type.      |
        Example:

        | *Keywords*           |  *Parameters*                                              |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Check Cell Type      |  TestSheet1                                        | 0 | 0 |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        sheet = self.workbook.sheet_by_index(my_sheet_index)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        if cell.ctype is XL_CELL_NUMBER:
            print("The cell value is a number")
        elif cell.ctype is XL_CELL_TEXT:
            print("The cell value is a string")
        elif cell.ctype is XL_CELL_DATE:
            print("The cell value is a date")
        elif cell.ctype is XL_CELL_BOOLEAN:
            print("The cell value is a boolean operator")
        elif cell.ctype is XL_CELL_ERROR:
            print("The cell value has an error")
        elif cell.ctype is XL_CELL_BLANK:
            print("The cell value is blank")
        elif cell.ctype is XL_CELL_EMPTY:
            print("The cell value is empty")
        else:
            print(error_text_from_code[sheet.cell(row, column).value])

    def put_number_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the number given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.                                           |
                |  Column (int)        | The column integer value that will be used to modify the cell.                                    |
                |  Row (int)           | The row integer value that will be used to modify the cell.                                       |
                |  Value (int)         | The integer value that will be added to the specified sheetname at the specified column and row.  |
        Example:

        | *Keywords*           |  *Parameters*                                                         |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |      |
        | Put Number To Cell   |  TestSheet1                                        |  0  |  0  |  34  |

        """
        if self.workbook:
            my_sheet_index = self.sheet_names.index(sheetname)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_NUMBER:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            plain = easyxf('')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), float(value), plain)

    def put_string_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the string given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.                                           |
                |  Column (int)        | The column integer value that will be used to modify the cell.                                    |
                |  Row (int)           | The row integer value that will be used to modify the cell.                                       |
                |  Value (string)      | The string value that will be added to the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                           |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |        |
        | Put String To Cell   |  TestSheet1                                        |  0  |  0  |  Hello |

        """
        if self.workbook:
            my_sheet_index = self.sheet_names.index(sheetname)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_TEXT:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            plain = easyxf('')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), value, plain)

    def put_date_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the date given in the parameter.

        Arguments:
                |  Sheet Name (string)               | The selected sheet that the cell will be modified from.                                                            |
                |  Column (int)                      | The column integer value that will be used to modify the cell.                                                     |
                |  Row (int)                         | The row integer value that will be used to modify the cell.                                                        |
                |  Value (int)                       | The integer value containing a date that will be added to the specified sheetname at the specified column and row. |
        Example:

        | *Keywords*           |  *Parameters*                                                               |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |            |
        | Put Date To Cell     |  TestSheet1                                        |  0  |  0  |  12.3.1999 |

        """
        if self.workbook:
            my_sheet_index = self.sheet_names.index(sheetname)
            cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
            if cell.ctype is XL_CELL_DATE:
                self.workbook.sheets()
                if not self.copied_workbook:
                    self.copied_workbook = copy(self.workbook)
        if self.copied_workbook:
            print(value)
            date_str = value.split('.')
            date_arr = [int(date_str[2]), int(date_str[1]), int(date_str[0])]
            print(date_str, date_arr)
            ymd = datetime(*date_arr)
            plain = easyxf('', num_format_str='d.M.yyyy')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), ymd, plain)

    def modify_cell_with(self, sheetname, column, row, op, val):
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
        my_sheet_index = self.sheet_names.index(sheetname)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        curval = cell.value
        if cell.ctype is XL_CELL_NUMBER:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            plain = easyxf('')
            modexpr = str(curval) + op + val
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), eval(modexpr), plain)

    def add_to_date(self, sheetname, column, row, numdays):
        """
        Using the sheet name the number of days are added to the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.                                                                          |
                |  Column (int)                    | The column integer value that will be used to modify the cell.                                                                   |
                |  Row (int)                       | The row integer value that will be used to modify the cell.                                                                      |
                |  Number of Days (int)            | The integer value containing the number of days that will be added to the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Add To Date          |  TestSheet1                                        |  0  |  0  |  4  |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        if cell.ctype is XL_CELL_DATE:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            curval = datetime(*xldate_as_tuple(cell.value, self.workbook.datemode))
            newval = curval + timedelta(int(numdays))
            plain = easyxf('', num_format_str='d.M.yyyy')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), newval, plain)

    def subtract_from_date(self, sheetname, column, row, numdays):
        """
        Using the sheet name the number of days are subtracted from the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.                                                                                 |
                |  Column (int)                    | The column integer value that will be used to modify the cell.                                                                          |
                |  Row (int)                       | The row integer value that will be used to modify the cell.                                                                             |
                |  Number of Days (int)            | The integer value containing the number of days that will be subtracted from the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Subtract From Date   |  TestSheet1                                        |  0  |  0  |  7  |

        """
        my_sheet_index = self.sheet_names.index(sheetname)
        cell = self.workbook.get_sheet(my_sheet_index).cell(int(row), int(column))
        if cell.ctype is XL_CELL_DATE:
            self.workbook.sheets()
            if not self.copied_workbook:
                self.copied_workbook = copy(self.workbook)
            curval = datetime(*xldate_as_tuple(cell.value, self.workbook.datemode))
            newval = curval - timedelta(int(numdays))
            plain = easyxf('', num_format_str='d.M.yyyy')
            self.copied_workbook.get_sheet(my_sheet_index).write(int(row), int(column), newval, plain)

    def save_excel(self, filename, use_temp_dir=False):
        """
        Saves the Excel file indicated by file name, the `use_temp_dir` can be set to true if the user needs the file saved in the temporary directory.
        If the boolean `use_temp_dir` is set to true, depending on the operating system of the computer running the test the file will be saved in the Temp directory if the operating system is Windows or tmp directory if it is not.

        Arguments:
                |  File Name (string)                      | The name of the of the file to be saved.  |
                |  Use Temporary Directory (default=False) | The file will not be saved in a temporary directory by default. To activate and save the file in a temporary directory, pass 'True' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Save Excel           |  NewExcelRobotTest.xls                             |

        """
        file_path = self.__get_file_path(filename, use_temp_dir)
        print('*DEBUG* Got file path %s' % file_path)
        self.copied_workbook.save(file_path)

    def create_sheet(self, sheet_name):
        """
        Creates and appends new Excel worksheet using the new sheet name to the current workbook.

        Arguments:
                |  New Sheet name (string)  | The name of the new sheet added to the workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Create New Sheet     |  NewSheet                                          |

        """
        self.copied_workbook = copy(self.workbook)
        self.copied_workbook.add_sheet(sheet_name)

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
        self.copied_workbook = Workbook()
        self.copied_workbook.add_sheet(sheet_name)
