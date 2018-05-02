Function button16()

	excelPath= document.getElementById("Q16_1").value
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(excelPath)
	Dim character(5)
	
	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)

	rowCount1 = objWorksheet1.UsedRange.Rows.Count
	
	For row = 2 to (rowCount1)

	'Name
		tempWord = objWorksheet1.Cells(row,1).Value

		Set re = New RegExp
		With re
		  .Pattern    = "^[a-zA-Z\s]+$"
		  .IgnoreCase = False
		  .Global     = False
		End With

		' Test method returns TRUE if a match is found
		If re.Test( tempWord ) AND Len(tempWord)<=60 Then
			objworksheet3.cells(row,1).Value="pass"
		Else
			objworksheet3.cells(row,1).Value="fail"
		End If

	'DOB
		tempWord = objWorksheet1.Cells(row,2).Value

		if (IsDate(tempWord)) then
			objworksheet3.cells(row,2).Value="pass"
		else 
			objworksheet3.cells(row,2).Value="fail"
		end if

	'Address- one line consist of 30 characters. assuming four lines containing
		tempWord = objWorksheet1.Cells(row,3).Value

		if (UBound(Split(tempWord,Chr(10))) = 3 ) then
			objworksheet3.cells(row,3).Value="pass"
		else 
			objworksheet3.cells(row,3).Value="fail"
		end if

	'City
		For row1 = 1 to (objWorksheet2.UsedRange.Rows.Count)
			if (objworksheet1.cells(row,4).Value = objworksheet2.cells(row1,1).Value) then
				objworksheet3.cells(row,4).Value="pass"
				Exit For
			else 
				objworksheet3.cells(row,4).Value="fail" 
			end if
		Next

	'Emp_ID
		tempWord = objWorksheet1.Cells(row,5).Value
		character(0)= Mid(tempWord,1,1) 
		character(1)= Mid(tempWord,2,1)
		character(2)= Mid(tempWord,3,1)
		character(3)= Mid(tempWord,4,1)
		character(4)= Mid(tempWord,5,1)
		character(5)= Mid(tempWord,6,1)
		if (Len(tempWord)=6 AND VarType(character(0))=8 AND IsNumeric(character(1)) AND IsNumeric(character(2)) AND IsNumeric(character(3)) AND IsNumeric(character(4)) AND IsNumeric(character(5)) ) then
			objworksheet3.cells(row,5).Value="pass"
		else 
			objworksheet3.cells(row,5).Value="fail"
		end if

	Next


	objWorkbook.Save
	objExcel.Quit	
	
	MsgBox "Data Validated Successfully",vbInformation 	  

End Function