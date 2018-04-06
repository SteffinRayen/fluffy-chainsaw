Function button13()

	FilePath1 = document.getElementById("Q13_1").value
	FilePath2 = document.getElementById("Q13_2").value
	
	Set objExcel1 = CreateObject("Excel.Application")
	objExcel1.Application.Visible = True
	Set objWorkbook1 = objExcel1.Workbooks.Open(FilePath1)
	Set objWorksheet1 = objWorkbook1.Worksheets(1)
	
	rowCount1 = objExcel1.ActiveWorkbook.Sheets(1).UsedRange.Rows.count

	ReDim CountryArray (rowCount1 - 1)

	Set objExcel2 = CreateObject("Excel.Application")
	objExcel2.Application.Visible = True
	Set objWorkbook2 = objExcel2.Workbooks.Open(FilePath2)
	Set objWorksheet2 = objWorkbook2.Worksheets(1)

	for intRow  = 1 to rowCount1 step 1
		CountryArray(introw -1) = objWorksheet1.Cells(intRow, 2).Value		
	next
	
	Dim UniqCountry : UniqCountry = uniq(CountryArray)
	NoOfCountries = uBound (UniqCountry) + 1
	
	Dim CountryDict : Set CountryDict = CreateObject("Scripting.Dictionary")
	count = 1
	for each word in UniqCountry
		objWorksheet2.Cells(count, 1).Value = word
		objWorksheet2.Cells(count, 2).Value = 0
		count = count + 1
	next

	for intRow1 = 1 to rowCount1 step 1
		
		for intRow2 = 1 to NoOfCountries step 1
			If objWorksheet1.Cells(intRow1, 2).Value = objWorksheet2.Cells(intRow2, 1).Value then
				objWorksheet2.Cells(intRow2, 2).Value = objWorksheet2.Cells(intRow2, 2).Value + 1
				temp = objWorksheet2.Cells(intRow2, 2).Value
				objWorksheet2.Cells(intRow2, temp + 2).Value = objWorksheet1.Cells(intRow1, 1).Value 
			End If
		next

	next
	
	
	objWorkbook1.Save
	objExcel1.Quit
	objWorkbook2.Save
	objExcel2.Quit

	MsgBox "Data sorted Successfully",vbInformation

End Function

Function uniq(array)
  Dim dicTemp : Set dicTemp = CreateObject("Scripting.Dictionary")
  Dim xItem
  For Each xItem In array
      dicTemp(xItem) = 0
  Next
  uniq = dicTemp.Keys()
End Function
