Function button16()

	excelPath= document.getElementById("Q16_1").value
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(excelPath)
	
	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)

	rowCount1 = objWorksheet1.UsedRange.Rows.Count
	
	'Name
	For row = 2 to (rowCount1-1)

		word = objWorksheet1.Cells(row,1).Value

		if (Len(word)<=60 AND VarType(word)=8) then
			objworksheet3.cells(row,1).Value="pass"
		else 
			objworksheet3.cells(row,1).Value="fail"
		end if
	Next

	'DOB
	For row = 2 to (rowCount1-1)

		word1 = objWorksheet1.Cells(row,2).Value

		if (IsDate(word1)) then
			objworksheet3.cells(row,2).Value="pass"
		else 
			objworksheet3.cells(row,2).Value="fail"
		end if
	Next

	'Address- one line consist of 30 characters. assuming four lines containing '120characters.
	For row = 2 to (rowCount1-1)

		word2 = objWorksheet1.Cells(row,3).Value

		if (VarType(word2)=8 AND Len(word2)>120 ) then
			objworksheet3.cells(row,3).Value="pass"
		else 
			objworksheet3.cells(row,3).Value="fail"
		end if
	Next

	'City
	For row = 2 to (rowCount1)
		For row1 = 1 to (objWorksheet2.UsedRange.Rows.Count)
			if (objworksheet1.cells(row,4).Value = objworksheet2.cells(row1,1).Value) then
				objworksheet3.cells(row,4).Value="pass"
				Exit For
			else 
				objworksheet3.cells(row,4).Value="fail" 
			end if
		Next
	Next

	'Emp_ID
	Dim words(5)
	For row = 2 to (rowCount1-1)
		word3 = objWorksheet1.Cells(row,5).Value
		words(0)= Mid(word3,1,1) 
		words(1)= Mid(word3,2,1)
		words(2)= Mid(word3,3,1)
		words(3)= Mid(word3,4,1)
		words(4)= Mid(word3,5,1)
		words(5)= Mid(word3,6,1)
		if (Len(word3)=6 AND VarType(words(0))=8 AND IsNumeric(words(1)) AND IsNumeric(words(2)) AND IsNumeric(words(3)) AND IsNumeric(words(4)) AND IsNumeric(words(5)) ) then
			objworksheet3.cells(row,5).Value="pass"
		else 
			objworksheet3.cells(row,5).Value="fail"
		end if
	Next

	objWorkbook.Save
	objExcel.Quit	
	
	MsgBox "Data Validated Successfully",vbInformation 	  

End Function