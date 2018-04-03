Function button07()
	'Open a workbook
	filePath = InputBox("Enter the file path","Enter Value")

	'Launch Excel
	Set objExcel = CreateObject("Excel.Application")

	'Set Excel to be visible
	objExcel.Application.Visible = True

	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	'Select a worksheet
	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)
	Set objWorksheet4 = objWorkbook.Worksheets(4)

	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	for intRow  = 2 to rowCount step 1
		FirstName = objWorksheet1.Cells(intRow, 1).Value
		LastName = objWorksheet2.Cells(intRow, 1).Value
		DOB = CDate(objWorksheet3.Cells(intRow, 1).Value)
		age = Round((Now() - dob) / 365.2425)
		objWorksheet4.Cells(intRow, 1).Value = FirstName&" "&LastName&" is of the age "&age
		
	next

	'Save the workbook,
	objWorkbook.Save

	'Quit Excel
	objExcel.Quit

	MsgBox "Data validated succesfully",vbInformation
End Function