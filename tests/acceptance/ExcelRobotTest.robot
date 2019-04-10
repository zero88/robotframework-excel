*** Settings ***
Library             ExcelRobot
Library             OperatingSystem
Suite Setup         Init Test

*** Variables ***
${Names}
${Num}
${Excel_File}       ExcelRobotTest
${xlsx}             xlsx
${xls}              xls
${type}             ${xlsx}
${Test_Data_Path}   ${CURDIR}${/}..${/}data${/}
${Out_Data_Path}    ${TEMPDIR}${/}excelrobot${/}
${SheetName}        Graph Data
${NewSheetName}     NewSheet
${NewSheetName}     NewSheet

*** Test Cases ***
Read Excel
    Get Values          excel_type=${type}

Write Excel 1
    Create New Excel    excel_type=${type}

Write Excel 2
    Create Excel From Existing File      excel_type=${type}

Write Excel 3
    Create New Sheet     excel_type=${type}

Write Excel 4
    Write New Value      excel_type=${type}

*** Keywords ***
Init Test
    Wait Until Keyword Succeeds     3x      2sec    Remove Directory    ${Out_Data_Path}    True
    Create Directory    ${Out_Data_Path}
    Copy Files          ${Test_Data_Path}${/}*  ${Out_Data_Path}

Get Values
    [Arguments]    ${excel_type}
    Open Excel     ${Out_Data_Path}${Excel_File}.${excel_type}
    ${Names}=      Get Sheet Names
    Set Suite Variable   ${Names}
    ${Num}=        Get Number of Sheets
    Set Suite Variable   ${Num}
    ${Col}=        Get Column Count    TestSheet1
    ${Row}=        Get Row Count       TestSheet1
    ${ColVal}=     Get Column Values   TestSheet2   1
    ${RowVal}=     Get Row Values      TestSheet2   1
    ${Sheet}=      Get Sheet Values    DataSheet
    Log   ${Sheet}
    ${Workbook}=   Get Workbook Values   False
    Log   ${Workbook}
    ${ByName}=     Read Cell Data By Name       GraphSheet   B2
    ${ByCoords}=   Read Cell Data               GraphSheet   1   1
    Check Cell Type      TestSheet1   0   1   TEXT


Create Excel From Existing File
    [Arguments]    ${excel_type}
    Open Excel To Write     ${Out_Data_Path}${Excel_File}.${excel_type}  ${Out_Data_Path}Clone_${Excel_File}.${excel_type}   True
    Save Excel

Create New Excel
    [Arguments]    ${excel_type}
    Open Excel To Write     ${Out_Data_Path}NewExcelSheet.${excel_type}
    Save Excel

Create New Sheet
    [Arguments]    ${excel_type}
    Open Excel To Write     ${Out_Data_Path}NewExcelSheet.${excel_type}
    Create Sheet            ${NewSheetName}
    Save Excel

Write New Value
    [Arguments]    ${excel_type}
    Open Excel To Write     ${Out_Data_Path}WriteExcelSheet.${excel_type}
    Create Sheet            ${NewSheetName}
    Write To Cell By Name   ${NewSheetName}     A1   abc
    Write To Cell By Name   ${NewSheetName}     A2   34
    Write To Cell By Name   ${NewSheetName}     A3   True
    Write To Cell           ${NewSheetName}     0       4       xx      TEXT
    Save Excel
