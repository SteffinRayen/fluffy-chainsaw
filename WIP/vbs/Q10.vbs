Function button10()

	filePath = document.getElementById("Q10_3").value

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePath)
	Set objWorksheet = objWorkbook.Worksheets(1)

	rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	colCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Columns.count  

	ReDim ColArray(0)  'empty array
	
	For i = 1 to colCount Step 2
	  ReDim Preserve ColArray(UBound(ColArray)+1)
	  ColArray(UBound(ColArray)) = objWorksheet.Cells(1, i).Value
	Next

	ColName = document.getElementById("Q11_2").value
	Order = document.getElementById("Q11_1").value

	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data Sorted Successfully",vbInformation
End Function