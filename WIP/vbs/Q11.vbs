Function button11()

	
	ExcelFilePath = document.getElementById("Q11").value
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True

	Set objWorkbook = objExcel.Workbooks.Open(ExcelFilePath)
	Set objWorksheet = objWorkbook.Worksheets(1)

	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	for intRow  = 1 to rowCount step 1

		'Get the value of cell A1
		strCellValue = Split(objExcel.Cells(introw, 1).Value)
		count = 2

		for each i in strCellValue

			'Put the value of strCellValue into cell A2
			temp = right(i,1)
			length_of_word = instrrev(i,temp)
			objWorksheet.Cells(intRow, count).Value = i &" "& length_of_word
			count = count + 1

		next
	next

	'Save the workbook,
	objWorkbook.Save
	'Quit Excel
	objExcel.Quit

	MsgBox "Data Split Successfully",vbInformation

End Function