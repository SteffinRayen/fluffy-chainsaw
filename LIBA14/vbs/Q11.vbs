Function button11()
	
	ExcelFilePath = document.getElementById("Q11_1").value
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(ExcelFilePath)
	Set objWorksheet = objWorkbook.Worksheets(1)
	
	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	
	for intRow  = 1 to rowCount step 1
		strCellValue = Split(objExcel.Cells(introw, 1).Value)
		count = 2
		for each word in strCellValue
			With (New RegExp)
				.Global = True
				.Pattern = "\W" 
				word = .Replace(word, "") 
			End With
			objWorksheet.Cells(intRow, count).Value = word &" "& instrrev(word,right(word,1))
			count = count + 1
		next
	next
	
	objWorkbook.Save
	objExcel.Quit
	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing
	
	MsgBox "Data Split Successfully",vbInformation

End Function