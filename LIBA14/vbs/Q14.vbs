Function button14()
	Dim objExcel, strExcelPath, objSheet1,objSheet2,objSheet3,companyname,totalrows,totalrows1,Name,tr,count
	
	strExcelPath = document.getElementById("Q14_1").value'
	
	' Open specified spreadsheet and select the first worksheet.
	Set objExcel = CreateObject("Excel.Application")
	objExcel.WorkBooks.Open strExcelPath
	Set objSheet1 = objExcel.ActiveWorkbook.Worksheets(1)
	Set objSheet2 = objExcel.ActiveWorkbook.Worksheets(2)
	Set objSheet3 = objExcel.ActiveWorkbook.Worksheets(3)
	totalrows=objSheet2.UsedRange.Rows.Count
	totalrows1=objSheet1.UsedRange.Rows.Count
	tr=objSheet3.UsedRange.Rows.Count
	
	Name = document.getElementById("Q14_2").value
	' Modify a cell.
	count=0
	For i=2 To totalrows
		If objSheet2.Cells(i,1).Value = Name Then
			tr=tr+1
			objSheet3.Cells(tr,1).Value = Name
			objSheet3.Cells(tr,2).Value = objSheet2.Cells(i,2)
			count=count+1
		End If
	Next
	tr=tr-count
	For i=2 To totalrows1
		If objSheet1.Cells(i,1).Value = Name Then
			tr=tr+1
			'objSheet3.Cells(tr,1).Value = Name
			objSheet3.Cells(tr,3).Value = objSheet1.Cells(i,2)
		End If
	Next

	MsgBox "Done",vbInformation
	
	' Save and quit.
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
End Function