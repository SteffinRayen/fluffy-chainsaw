Function button08()
	
	filePath = document.getElementById("Q08_1").value
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)
	Set objWorksheet4 = objWorkbook.Worksheets(4)

	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	for intRow  = 1 to rowCount step 1
		
		If objWorksheet1.Cells(intRow, 1).Value + objWorksheet2.Cells(intRow, 1).Value = objWorksheet3.Cells(intRow, 1).Value Then
			objWorksheet4.Cells(intRow, 1).Value = "Pass"
			objWorksheet4.Cells(intRow, 1).Interior.ColorIndex = 4
		Else
			objWorksheet4.Cells(intRow, 1).Value = "Fail"
			objWorksheet4.Cells(intRow, 1).Interior.ColorIndex = 3
		End If
	next

	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data validated succesfully",vbInformation
End Function