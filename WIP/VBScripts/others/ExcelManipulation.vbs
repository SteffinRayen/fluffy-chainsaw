Function ExcelManipulation()

	'Open a workbook
	filePath = InputBox("Enter the file path","Enter Value")


	If CreateObject("Scripting.FileSystemObject").FileExists(filePath) Then

		'Launch Excel
		Set objExcel = CreateObject("Excel.Application")

		'Set Excel to be visible
		objExcel.Application.Visible = True
	
		Set objWorkbook = objExcel.Workbooks.Open(filePath)

		'Select a worksheet
		Set objWorksheet = objWorkbook.Worksheets(1)

		'Get the value of cell A1
		strCellValue = objExcel.Cells(1, 1).Value

		'Put the value of strCellValue into cell A2
		objWorksheet.Cells(1, 2).Value = strCellValue

		'Save the workbook,
		objWorkbook.Save

		'Quit Excel
		objExcel.Quit

	Else
		MsgBox "File not found in the given path",vbOKOnly+vbCritical,"Incorrect path"
	End If


End Function