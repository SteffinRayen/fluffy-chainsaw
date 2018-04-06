Function button11()
	
	'ExcelFilePath = document.getElementById("Q11_1").value
	ExcelFilePath = InputBox ("Enter .xlsx file path")
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(ExcelFilePath)
	Set objWorksheet = objWorkbook.Worksheets(1)
	
	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	
	for intRow  = 1 to rowCount step 1
	'Get the value of cell A1
		strCellValue = Split(objExcel.Cells(introw, 1).Value)
		count = 2
	
		for each word in strCellValue
			'Put the value of strCellValue into cell A2
			length_of_word = instrrev(word,right(word,1))
			objWorksheet.Cells(intRow, count).Value = word &" "& length_of_word
			count = count + 1
		next
	next
	
	objWorkbook.Save
	objExcel.Quit
	
	MsgBox "Data Split Successfully",vbInformation

End Function