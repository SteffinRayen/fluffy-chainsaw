Function button15()

	filePathText = InputBox("Enter the file path .txt" ,"Enter Value")
	Set TextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile (filePathText, 1)

	filePathExcel = InputBox("Enter the file path .xlsx","Enter Value")
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePathExcel)
	Set objWorksheet = objWorkbook.Worksheets(1)

	row = 0
	Do Until TextFile.AtEndOfStream
	  line = TextFile.Readline
	  objWorksheet.Cells(row, 1).Value = line
	  row = row + 1
	Loop

	TextFile.Close
	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data Split Successfully",vbInformation


End Function