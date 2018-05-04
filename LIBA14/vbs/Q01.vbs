Function button01()
	
	filePathText = document.getElementById("Q01_1").value
	fileNameText = document.getElementById("Q01_2").value
	absoluteFilePathText = filePathText&"\"&fileNameText&".txt"
	
	Set TextFile = CreateObject("Scripting.FileSystemObject").CreateTextFile (absoluteFilePathText, 1)

	TextFile.Write "Automation"&vbCrLf&"of"&vbCrLf&"txt file"&vbCrLf&"Using"&vbCrLf&"VBS"
	
	TextFile.close

	MsgBox "Text File Created Succesfuly",vbInformation
	 
	'The Excel file to be created
	filePathExcel = document.getElementById("Q01_3").value 
	fileNameExcel = document.getElementById("Q01_4").value 
	absoluteFilePathExcel = filePathExcel&"\"&fileNameExcel

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = TRUE	 

	Set objWorkbook = objExcel.Workbooks.Open(absoluteFilePathText)
	 
	objExcel.ActiveWorkbook.SaveAs absoluteFilePathExcel,51
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit

	MsgBox "Data Tranfered Succesfuly",vbInformation

End Function