Function button11()

	'Get File path from hml text input using ID
	ExcelFilePath = document.getElementById("Q11_1").value
	
	'Create Excel Object
	Set objExcel = CreateObject("Excel.Application")

	'Make it visible
	objExcel.Application.Visible = True

	'Load an existing Workbook in the Excel object
	Set objWorkbook = objExcel.Workbooks.Open(ExcelFilePath)

	'Select the Worksheet to manipulate
	Set objWorksheet = objWorkbook.Worksheets(1)
	
	'Get the no of used rows in the sheet
	rowCount = objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	
	'Iterate over the rows of sentences
	for intRow  = 1 to rowCount step 1

		'Split sentence into array of words using space
		strCellValue = Split(objExcel.Cells(introw, 1).Value)

		'Keeping track of which cell to enter the word and the length
		count = 2

		'Iterate over the words in the sentence
		for each word in strCellValue

			'Remove non letters
			With (New RegExp)
				.Global = True
				.Pattern = "\W" '[A-Za-z0-9]
				word = .Replace(word, "") 
			End With

			'Storing value
			objWorksheet.Cells(intRow, count).Value = word &" "& instrrev(word,"")
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