# Robotframework-excel for Robot Framework

## Introduction

Robotframework-excel is a Robot Framework Library that provides keywords to allow opening, reading, writing and saving Excel files. The Robotframework-excel leverages two other python libraries [xlutils](https://pypi.python.org/pypi/xlutils/2.0.0) and [natsort](https://pypi.python.org/pypi/natsort/5.2.0). Xlutils installs [xlrd](https://pypi.python.org/pypi/xlrd) that reads data from an Excel file and [xlwt](https://pypi.python.org/pypi/xlwt) that can write to an Excel file.

- Information about Robotframework-excel keywords can be found on the [ExcelRobot-Keyword Documentation](https://zero-88.github.io/robotframework-excel/docs/ExcelRobot-KeywordDocumentation.html) page.
- Information about working with Excel files in Python can be found on the [Python Excel](http://www.python-excel.org/) page.
- Useful pdf for practical use with Excel files [here](http://www.simplistix.co.uk/presentations/python-excel.pdf).

## Requirements

- Python >= 2.7 | Python >= 3.3
- Robot Framework >= 3.0
- xlutils 2.0.0. Access the downloads [here](https://pypi.python.org/pypi/xlutils/1.7.1), or use pip install xlutils.
- natsort 5.2.0. Access the downloads [here](https://pypi.python.org/pypi/natsort/5.2.0), or use pip install natsort.

## Installation

The recommended installation tool is [pip](http://pip-installer.org).

Install pip. Enter the following (Append `--upgrade` to update both the library and all its dependencies to the latest version):

```bash
pip install robotframework-excel --upgrade
```

To install a specific version enter:

```bash
pip install robotframework-excel==(version)
```

### Uninstall

To uninstall Robotframework-excel use the following pip command:

```bash
pip uninstall robotframework-excel
```

## Project structure

- `ExcelRobot/base.py`: The Robot Python Library defines excel operation keyword.
- `tests/unit/*.py`: Unit test
- `tests/acceptance/ExcelRobotTest.robot`: Example robot test file to display what various keywords from Robotframework-excel accomplish
- `docs/ExcelRobot-KeywordDocumentation.html`: Keyword documentation for the Robotframework-excel.

## Usage

To write tests with Robot Framework and Robotframework-excel, `ExcelRobot` must be imported into your Robot test suite.
See [Robot Framework User Guide](http://code.google.com/p/robotframework/wiki/UserGuide) for more information.

## Running the Demo

The test file `ExcelRobotTest.robot`, is an easily executable test for Robot Framework using Robotframework-excel.

For in depth detail on how the keywords function, read the Keyword documentation found here : [Keyword Documentation](https://zero-88.github.io/robotframework-excel/docs/ExcelRobot.html)

To run the test navigate to the Tests directory in C:\Python folder. Open a command prompt within the `tests/acceptance` folder and run:

```bash
pybot ExcelRobotTest.robot -d "./out"
```

## Release Note

[Release Note Documentation](https://zero-88.github.io/robotframework-excel/docs/release-notes.md)

## Limitation

- When using the keyword `Add New Sheet` the user cannot perform any functions before or after this keyword on the currently open workbook. The changes that other keywords make will not be saved when the keyword `Add New Sheet` is used. They must add a sheet then save the workbook before using any other keyword. If they want to use any other keywords on the workbbok they must open the workbook again to do so.

## Getting Help

The [user group for Robot Framework](http://groups.google.com/group/robotframework-users) is the best place to get help. Include in the post:

- Contact the [Python-Excel google group](https://groups.google.com/forum/#!forum/python-excel)
- Full description of what you are trying to do and expected outcome
- Version number of Robotframework-excel and Robot Framework
- Traceback or other debug output containing error information