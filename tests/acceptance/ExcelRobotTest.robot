*** Settings ***
Library             ExcelRobot
Library             OperatingSystem
Suite Setup         Init Test

*** Variables ***
${Names}
${Num}
# ${Excel_File}     ExcelRobotTest.xlsx
${Excel_File}       ExcelRobotTest.xls
# ${Excel_File}     a.txt
${Test_Data_Path}   ${CURDIR}${/}..${/}data${/}
${Out_Data_Path}    ${TEMPDIR}${/}excelrobot${/}
${SheetName}        Graph Data
${NewSheetName}     NewSheet

*** Test Cases ***
Read Excel
    Get Values

Write Excel 1
    Create New Excel

Write Excel 2
    Create Excel From Existing File

Write Excel 3
    Create New Sheet

*** Keywords ***
Init Test
    Wait Until Keyword Succeeds     3x      2sec    Remove Directory    ${Out_Data_Path}    True
    Create Directory    ${Out_Data_Path}
    Copy Files          ${Test_Data_Path}${/}*  ${Out_Data_Path}

Get Values
    Open Excel     ${Out_Data_Path}${Excel_File}
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
    Open Excel To Write     ${Out_Data_Path}${Excel_File}  ${Out_Data_Path}Clone_${Excel_File}   True
    Save Excel

Create New Excel
    Open Excel To Write     ${Out_Data_Path}NewExcelSheet.xls
    Save Excel

Create New Sheet
    Open Excel To Write     ${Out_Data_Path}NewExcelSheet.xls
    Create Sheet            ${NewSheetName}
    Save Excel
