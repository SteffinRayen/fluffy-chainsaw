Function button10()
	'Open a workbook
	filePath = InputBox("Enter the file path","Enter Value")

	'Launch Excel
	Set objExcel = CreateObject("Excel.Application")

	'Set Excel to be visible
	objExcel.Application.Visible = True

	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	'Select a worksheet
	Set objWorksheet = objWorkbook.Worksheets(1)

	rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	colCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Columns.count  

	ReDim ColArray(0)  'empty array
	For i = 1 to colCount Step 2
	  ReDim Preserve ColArray(UBound(ColArray)+1)
	  ColArray(UBound(ColArray)) = objWorksheet.Cells(1, i).Value
	Next

	ColName = Input ("What should we order by ?"& Join(ColArray, vbNewLine))
	Order = Indup ("Ascending / Descending (A/D)")


	'Save the workbook,
	objWorkbook.Save

	'Quit Excel
	objExcel.Quit

	MsgBox "Data Sorted Successfully",vbInformation
End Function