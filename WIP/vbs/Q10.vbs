Function button10()

	filePath = document.getElementById("Q10_3").value

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePath)
	Set objWorksheet = objWorkbook.Worksheets(1)

	rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
	colCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Columns.count 
	

	ReDim ColArray(colCount) 
	
	For i = 1 to colCount Step 2
	  ColArray(UBound(ColArray)) = objWorksheet.Cells(1, i).Value
	Next

	ColName = document.getElementById("Q10_2").value
	
	If document.getElementById("Q10_1").value = "A" Then
		Order = 1
	Else
		Order = 2
	End If

	ColNum = 0

	for i = 1 to colCount step 1
		if objWorksheet.Cells(1, i).Value = ColName Then
			ColNum = i
		End If
	Next

	intRow = 0
	Select Case Order
		Case 1
			switching = 1
			Do 
				switching = 0
				for intRow = 2 to rowCount-1 step 1
					shouldSwitch = 0
					'Ascending
					if StrComp(objWorksheet.Cells(intRow, ColNum).Value, objWorksheet.Cells(intRow + 1, ColNum).Value) = 1 Then			
						shouldSwitch= 1
						Exit For
					End If
				Next
				if shouldSwitch Then
					For intCol = 1 to colCount step 1
						temp = objWorksheet.Cells(intRow , intCol).Value
						objWorksheet.Cells(intRow , intCol).Value = objWorksheet.Cells(intRow + 1, intCol).Value
						objWorksheet.Cells(intRow + 1, intCol).Value = temp
					Next

					switching = 1
				End If

			Loop while switching
		Case 2
			switching = 1
			Do 
				switching = 0
				for intRow = 2 to rowCount step 1
					shouldSwitch = 0
					'Descending
					if StrComp(objWorksheet.Cells(intRow, ColNum).Value, objWorksheet.Cells(intRow + 1, ColNum).Value) = -1 Then
						shouldSwitch= intRow
						Exit For
					End If
				Next
				if shouldSwitch > 0 Then
					For intCol = 1 to colCount step 1
						temp = objWorksheet.Cells(intRow , intCol).Value
						objWorksheet.Cells(intRow , intCol).Value = objWorksheet.Cells(intRow + 1, intCol).Value
						objWorksheet.Cells(intRow + 1, intCol).Value = temp
					Next
					switching = 1
				End If
			Loop while switching
	End Select

	objWorkbook.Save
	objExcel.Quit

	MsgBox "Data Sorted Successfully",vbInformation
End Function