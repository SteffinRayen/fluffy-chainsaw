Function button04()

	Input = document.getElementById("Q04_1").value
	Output = document.getElementById("Q04_4").value

	TextName = document.getElementById("Q04_2").value
	XcelName = document.getElementById("Q04_5").value
	Count = document.getElementById("Q04_3").value
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Workbooks.Add
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	objSheet.Name = "Customer Details"

	'Output Should be without file name
	strExcelPath = Output+"\"+XcelName+".xlsx"

	objSheet.Cells(1, 1).Value = "Account Number" 'Row 1 Column 1 (A)
	objSheet.Cells(1, 2).Value = "Customer Name" 'Row 1 Column 2 (B)
	objSheet.Cells(1, 3).Value = "Customer ID" 'Row 1 Column 3 (C)

	Set FSO = CreateObject("Scripting.FileSystemObject")
	for textFile = 1 to Count

		'Input should be till before \
		Set ReadTextFile = FSO.OpenTextFile(Input+"\"&TextName&" ("&textFile&").txt", 1)
		Do Until ReadTextFile.AtEndOfStream

			Textline = ReadTextFile.Readline()
			
			If Instr(Textline, "Account Number:") Then ' If textline contain string "Account Number :"
			  objSheet.Cells(textFile +1, 1).Value = Split(Textline, ":")(1) ' Split the textline  with ":" indicator
			End If
			
			If Instr(Textline, "Customer Name:") Then
			  objSheet.Cells(textFile +1, 2).Value = Split(Textline, ":")(1)
			End If
			
			If Instr(Textline, "Customer ID:") Then
			   objSheet.Cells(textFile +1, 3).Value = Split(Textline, ":")(1)
			End If

		Loop
	next

	MsgBox "Data collated Successfully",vbInformation

	
	objExcel.ActiveWorkbook.SaveAs strExcelPath
	objExcel.ActiveWorkbook.Close
	
	objExcel.Application.Quit
	Set objSheet = Nothing
	Set objExcel = Nothing



End Function