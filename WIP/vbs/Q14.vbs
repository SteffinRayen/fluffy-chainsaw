Function button14()
	filePath = document.getElementById("Q14_1").value

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	
	Set objWorkbook = objExcel.Workbooks.Open(filePath)
	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)

	rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	intRow3 = 2
	for intRow1 = 2 to rowCount step 1
		for intRow2 = 2 to rowCount step 1
			if objWorksheet1.Cells(intRow1, 1).Value = objWorksheet2.Cells(intRow2, 1).Value Then
				objWorksheet3.Cells(intRow3, 1).Value = objWorksheet1.Cells(intRow1, 1).Value
				objWorksheet3.Cells(intRow3, 2).Value = objWorksheet1.Cells(intRow1, 2).Value
				objWorksheet3.Cells(intRow3, 3).Value = objWorksheet2.Cells(intRow2, 2).Value
				intRow3 = intRow3 + 1
				Exit For
			End If
		Next
	Next

	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data Collated Successfully",vbInformation
End Function