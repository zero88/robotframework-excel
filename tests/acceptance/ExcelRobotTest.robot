*** Settings ***
Library 			ExcelRobot
Library 			Collections

*** Variables ***
${Names}
${Num}
# ${Excel_File}		ExcelRobotTest.xlsx
${Excel_File}		ExcelRobotTest.xls
# ${Excel_File}		a.txt
${Test_Data_Path}   ${CURDIR}${/}..${/}data${/}
${Out_Data_Path}	${CURDIR}${/}..${/}..${/}out${/}
${SheetName}		Graph Data
${NewSheetName}		NewSheet

*** Test Cases ***
Excel Test
	Get Values and Modify Spreadsheet
	Add Date To Sheet
	Perform Function and Change Date
	Create a New Excel
	Create a New Sheet
	Check New Sheet Values

*** Keywords ***
Get Values and Modify Spreadsheet
	Open Excel   ${Test_Data_Path}${Excel_File}
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
	${ByName}=     Read Cell Data By Name          GraphSheet   B2
	${ByCoords}=   Read Cell Data By Coordinates   GraphSheet   1   1
	Check Cell Type      TestSheet1   0   1
	Put Number To Cell   TestSheet1   1   1   90
	Put String To Cell   TestSheet3   1   1   yellow
	Put Date To Cell     TestSheet2   1   1   1.4.1989
	Put Date To Cell     TestSheet2   1   2   12.10.1991
	Save Excel           ${Out_Data_Path}TestExcel.xls

Add Date To Sheet
	Open Excel        ${Out_Data_Path}TestExcel.xls
	Add To Date       TestSheet2   1   2   5
    Check Cell Type   TestSheet2   1   2
	Save Excel        ${Out_Data_Path}NewDateExcel.xls

Perform Function and Change Date
	Open Excel           ${Out_Data_Path}NewDateExcel.xls
	Modify Cell With     TestSheet1   1   1   *   45
	Subtract From Date   TestSheet2   1   1   1
	Save Excel           ${Out_Data_Path}FunctionExcel.xls

Create a New Excel
	Create Workbook		NewExcelSheet
	Save Excel			${Out_Data_Path}NewExcel.xls

Create a New Sheet
	Open Excel		${Out_Data_Path}FunctionExcel.xls
	Create Sheet	${NewSheetName}
	Save Excel      ${Out_Data_Path}NewSheetExcel.xls

Check New Sheet Values
	Open Excel     ${Out_Data_Path}NewSheetExcel.xls
	${NewNames}=   Get Sheet Names
	${NewNum}=     Get Number of Sheets
	Should Not Be Equal As Strings    ${Names}   ${NewNames}
	Should Not Be Equal As Integers   ${Num}     ${NewNum}
	${Sheet}=      Get Sheet Values   TestSheet3   False
	Log            ${Sheet}
	${stringList}=   Convert To String   ${Sheet}
	Should Contain   ${stringList}   yellow
