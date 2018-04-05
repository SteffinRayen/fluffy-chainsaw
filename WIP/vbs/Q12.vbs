Function button12()

	TextFilePath = InputBox("Enter the file path .txt" ,"Enter Value")
	Set TextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(TextFilePath)

	ExcelFilePath =  InputBox("Enter the file path .xlsx","Enter Value")
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(ExcelFilePath)
	Set objWorksheet = objWorkbook.Worksheets(1)

	CounterA = 2
	CounterB = 2

	do while not TextFile.AtEndOfStream

		strLine = TextFile.ReadLine()
		strWords = Split(Trim(strLine))

		for each word in strWords

			If IsNumeric(word) Then
				objWorksheet.Cells(CounterA, 1).Value = word
				CounterA = CounterA + 1
			Else
				objWorksheet.Cells(CounterB, 2).Value = word
				CounterB = CounterB + 1
			End if
	
		next
	loop
	TextFile.Close
	Set TextFile = Nothing
	objWorkbook.Save
	objExcel.Quit
	MsgBox "Data categorized Successfully",vbInformation


End Function