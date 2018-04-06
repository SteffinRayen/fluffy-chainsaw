Function button01()
	
	filePathText = document.getElementById("Q01_1").value
	Set TextFile = CreateObject("Scripting.FileSystemObject").CreateTextFile (filePathText, 1)
	TextFile.Write "Automation"&vbCrLf&"of"&vbCrLf&"txt file"&vbCrLf&"Using"&vbCrLf&"VBS"
	TextFile.close

	MsgBox "Data Created Succesfuly Successfully",vbInformation
	 
	'The Excel file to be created
	strOutput = document.getElementById("Q01_2").value 
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = TRUE	 

	Set objWorkbook = objExcel.Workbooks.Open(filePathText)
	 
	objExcel.ActiveWorkbook.SaveAs strOutput, 1
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit

	MsgBox "Data Tranfered Succesfuly Successfully",vbInformation

End Function