Function button07()

	filePath = document.getElementById("Q07_1").value
	filePath2 = document.getElementById("Q07_2").value

	Set objExcel = CreateObject("Excel.Application")
	Set objExcel2 = CreateObject("Excel.Application")

	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePath)
	
	objExcel2.Application.Visible = True
	Set objWorkbook2 = objExcel2.Workbooks.Open(filePath2)

	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)
	
	Set objWorksheet4 = objWorkbook2.Worksheets(1)

	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	for intRow  = 2 to rowCount step 1
		FirstName = objWorksheet1.Cells(intRow, 1).Value
		LastName = objWorksheet2.Cells(intRow, 1).Value
		DOB = CDate(objWorksheet3.Cells(intRow, 1).Value)
		age = Round((Now() - dob) / 365.2425)
		objWorksheet4.Cells(intRow, 1).Value = FirstName&" "&LastName&" is of the age "&age		
	next

	objWorkbook2.Save
	objExcel.Quit
	objExcel2.Quit

	MsgBox "Data populated succesfully",vbInformation

End Function