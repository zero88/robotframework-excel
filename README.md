# Robotframework-excel for Robot Framework

[![Version](https://img.shields.io/pypi/v/robotframework-excel.svg)](https://img.shields.io/pypi/v/robotframework-excel.svg)
[![License](https://img.shields.io/pypi/l/robotframework-excel.svg)](https://img.shields.io/pypi/l/robotframework-excel.svg)
[![Status](https://img.shields.io/pypi/status/robotframework-excel.svg)](https://img.shields.io/pypi/status/robotframework-excel.svg)
[![PyVersion](https://img.shields.io/pypi/pyversions/robotframework-excel.svg)](https://img.shields.io/pypi/pyversions/robotframework-excel.svg)

[![Build Status](https://travis-ci.org/zero-88/robotframework-excel.svg?branch=master)](https://travis-ci.org/zero-88/robotframework-excel)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=robotframework-excel&metric=coverage)](https://sonarcloud.io/component_measures?id=robotframework-excel&metric=coverage)
[![Quality Status](https://sonarcloud.io/api/project_badges/measure?project=robotframework-excel&metric=alert_status)](https://sonarcloud.io/dashboard?id=robotframework-excel)

## Introduction

Robotframework-excel is a Robot Framework Library that provides keywords to allow opening, reading, writing and saving Excel files.

- Information about Robotframework-excel keywords can be found on the [ExcelRobot-Keyword Documentation](https://zero88.github.io/robotframework-excel/docs/ExcelRobot.html) page.
- Information about working with Excel files in Python can be found on the [Python Excel](http://www.python-excel.org/) page.

## Requirements

- Python >= 2.7 | Python >= 3.3
- Robot Framework >= 3.0
- xlutils 2.0.0. Access the downloads [here](https://pypi.python.org/pypi/xlutils/1.7.1), or use pip install xlutils.
  - [xlrd](https://pypi.python.org/pypi/xlrd) that reads data from an Excel file
  - [xlwt](https://pypi.python.org/pypi/xlwt) that can write to an Excel file.
- openpyxl 1.0.2
- natsort 5.2.0. Access the downloads [here](https://pypi.python.org/pypi/natsort/5.2.0), or use pip install natsort.
- enum34 1.1.6

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
- `docs/ExcelRobot.html`: Keyword documentation for the Robotframework-excel.

## Usage

To write tests with Robot Framework and Robotframework-excel, `ExcelRobot` must be imported into your Robot test suite.
See [Robot Framework User Guide](http://code.google.com/p/robotframework/wiki/UserGuide) for more information.

## Running the Demo

The test file `ExcelRobotTest.robot`, is an easily executable test for Robot Framework using Robotframework-excel.

For in depth detail on how the keywords function, read the Keyword documentation found here : [Keyword Documentation](https://zero88.github.io/robotframework-excel/docs/ExcelRobot.html)

Open a command prompt within the `tests/acceptance` folder and run:

```bash
pybot ExcelRobotTest.robot -d "./out"
```

## Release Note

[Release Note Documentation](https://zero88.github.io/robotframework-excel/docs/release-notes.md)

## Limitation

- Lack `DataType` is `CURRENCY` and `PERCENTAGE`
- Not yet optimize performance when saving Excel file after modifying itself

## Contribution

The [user group for Robot Framework](http://groups.google.com/group/robotframework-users) is the best place to get help. Include in the post:

- Contact the [Python-Excel google group](https://groups.google.com/forum/#!forum/python-excel)
- Full description of what you are trying to do and expected outcome
- Version number of Robotframework-excel and Robot Framework
- Traceback or other debug output containing error information
