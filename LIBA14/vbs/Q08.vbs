Function button08()

	filePath = document.getElementById("Q08_1").value
	fileName = document.getElementById("Q08_2").value
	absoluteFilePath = filePath&"\"&fileName&".xlsx"
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(absoluteFilePath)

	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)
	Set objWorksheet4 = objWorkbook.Worksheets(4)

	rowCount = InputBox("How many numbers do you want to enter?")

	For i=1 To rowCount
		objWorksheet1.Cells(i, 1).Value=InputBox("Enter First Number (Sheet1): "&i)
		objWorksheet2.Cells(i, 1).Value=InputBox("Enter Second Number (Sheet2): "&i)
		objWorksheet3.Cells(i, 1).Value=InputBox("Enter the sum of "&objWorksheet1.Cells(i, 1).Value&" "&objWorksheet1.Cells(i, 1).Value)
	Next

	for intRow  = 1 to rowCount step 1
		If objWorksheet1.Cells(intRow, 1).Value + objWorksheet2.Cells(intRow, 1).Value = objWorksheet3.Cells(intRow, 1).Value Then
			objWorksheet4.Cells(intRow, 1).Value = "Pass"
			objWorksheet4.Cells(intRow, 1).Interior.ColorIndex = 4
		Else
			objWorksheet4.Cells(intRow, 1).Value = "Fail"
			objWorksheet4.Cells(intRow, 1).Interior.ColorIndex = 3
		End If
	next

	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data validated succesfully",vbInformation

End Function