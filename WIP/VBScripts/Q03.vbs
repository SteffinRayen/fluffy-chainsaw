Function button03()

	'Open a workbook
	filePath = InputBox("Enter the file path","Enter Value")

	'Launch Excel
	Set objExcel = CreateObject("Excel.Application")

	'Set Excel to be visible
	objExcel.Application.Visible = True

	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	for each sheet in objWorkbook.Worksheets

		sheet.Cells(1, 1).Value = "Yo :)"
		
	next

	'Save the workbook,
	objWorkbook.Save

	'Quit Excel
	objExcel.Quit

	MsgBox "Data Split Successfully",vbInformation

End Function