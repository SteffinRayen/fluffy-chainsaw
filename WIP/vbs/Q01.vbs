Function button01()
	
	filePathText = InputBox("Enter the file path .txt" ,"Enter Value")
	Set TextFile = CreateObject("Scripting.FileSystemObject").CreateTextFile (filePathText, 1)
	TextFile.Write "Automation"&vbCrLf&"of"&vbCrLf&"txt file"&vbCrLf&"Using"&vbCrLf&"VBS"
	TextFile.close

	MsgBox "Data Created Succesfuly Successfully",vbInformation
	 
	'The Excel file to be created
	strOutput = InputBox("Enter the file path .xlsx without .xlsx" ,"Enter Value")
	 
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = TRUE
	 
	Set objWorkbook = objExcel.Workbooks.Open(filePathText)
	 
	objExcel.ActiveWorkbook.SaveAs strOutput, 1
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit

	MsgBox "Data Tranfered Succesfuly Successfully",vbInformation

End Function