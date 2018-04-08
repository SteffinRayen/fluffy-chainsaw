Function button12()
	
	TextFilePath = document.getElementById("Q12_1").value
	Set TextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(TextFilePath)
	
	ExcelFilePath = document.getElementById("Q12_2").value
	ExcelFile = document.getElementById("Q12_3").value
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	objExcel.Workbooks.Add
	Set objWorksheet = objExcel.ActiveWorkbook.Worksheets(1)
	
	strExcelPath = ExcelFilePath+"\"+ExcelFile+".xlsx"

	objWorksheet.Name = "Content"
	objWorksheet.Cells(1, 1).Value = "String" 		'Row 1 Column 1 (A)
	objWorksheet.Cells(1, 2).Value = "Integer" 		'Row 1 Column 2 (B)
	objWorksheet.Cells(1, 3).Value = "Long" 		'Row 1 Column 3 (C)
	objWorksheet.Cells(1, 4).Value = "Double" 		'Row 1 Column 4 (D)
	objWorksheet.Cells(1, 5).Value = "Date" 		'Row 1 Column 5 (E)
	objWorksheet.Cells(1, 6).Value = "Time" 		'Row 1 Column 6 (F)
	'Read Lines & and Store into Array
	'``````````````````````````````````	
	Dim a()
	i = 0
	Do Until TextFile.AtEndOfStream
		Redim Preserve a(i)
		a(i) = TextFile.ReadLine
		i = i + 1
	Loop
	TextFile.Close
	m=0
	'Splitting Line into Array (word by Word)
	'`````````````````````````````````````````
	For Each strLine in a
		inputText = strLine
		outputArray = Split(inputText)
		For Each x in outputArray
			Redim Preserve copy(m)
			copy(m)=x
			m=m+1
		Next
	Next
	'Regular Expressions
	'````````````````````````
	'Date
	'```````````````````````````
	Set rDate1 = New RegExp
	With rDate1
		.Pattern    = "^(19|20)\d\d[(-|/|\|.) /.](0[1-9]|1[012])[(-|/|\|.) /.](0[1-9]|[12][0-9]|3[01])$"
		.IgnoreCase = True
		.Global     = True
	End With
	Set rDate2 = New RegExp
	With rDate2
		.Pattern    = "^(0[1-9]|[12][0-9]|3[01])[(-|/|\|.) /.](0[1-9]|1[012])[(-|/|\|.) /.](19|20)\d\d$"		
		.IgnoreCase = True
		.Global     = True
	End With
	'Integer
	'```````````````````````````
	Set rInteger = New RegExp
	With rInteger
		.Pattern    = "^([0-9]{1,4}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])$"
		.IgnoreCase = True
		.Global     = True
	End With
	'Long
	'`````````````````````````
	Set rLong = New RegExp
	With rLong
		.Pattern    = "^[1-9][2-9](?!00$)[0-9][1-9]?\d+$"
		.IgnoreCase = True
		.Global     = True
	End With
	'Double
	'`````````````````````````
	Set rFloat = New RegExp
	With rFloat
		.Pattern    = "^(\d)*[./.](\d)*$"
		.IgnoreCase = True
		.Global     = True
	End With
	'Time 
	'````````````````````````
	Set rTime = New RegExp
	With rTime			' HH			:	09		19/(1-5)(0-9)	
		.Pattern    = "^(0[0-9]|1[012]|(1[3-9)|(2[0123]))[:/.](0[0-9]|[1-5][0-9])$"
		.IgnoreCase = True
		.Global     = True
	End With
	' String & Character
	'````````````````````````
	Set rString = New RegExp
	With rString
		.Pattern    = "[^\d\s]$"
		.IgnoreCase = True
		.Global     = True
	End With
	'Variables to Count
	StrCount=2
	IntCount=2	
	LongCount=2
	DoubleCount=2
	DateCount=2
	TimeCount=2
	'DataType wise
	'```````````````````````````````````````````````
	for k=0 to UBound(copy)
		if(rString.Test(copy(k))) then
		objWorksheet.Cells(StrCount, 1).Value = copy(k)
		StrCount=StrCount+1
	End if
	If(rInteger.Test(copy(k))) then
		objWorksheet.Cells(IntCount, 2).Value = copy(k)
		IntCount=IntCount+1
	End If
	If(rLong.Test(copy(k))) then
		objWorksheet.Cells(LongCount, 3).Value = copy(k)
		LongCount=LongCount+1
	End If
	if(rFloat.Test(copy(k))) then
		objWorksheet.Cells(DoubleCount, 4).Value = copy(k)
		DoubleCount=DoubleCount+1
	End If
	if(rDate1.Test(copy(k))) then
		objWorksheet.Cells(DateCount, 5).Value = copy(k)
		DateCount=DateCount+1
	End If
	If(rDate2.Test(copy(k))) then
		objWorksheet.Cells(DateCount, 5).Value = copy(k)
		DateCount=DateCount+1
	End If
	if(rTime.Test(copy(k))) then
		objWorksheet.Cells(TimeCount, 6).Value = copy(k)
		TimeCount=TimeCount+1
	End if
	Next
	'FORMAT
	'`````````````````````````````````
	objWorksheet.Range("A1:F1").Font.Italic = True 		'Italic 
	objWorksheet.Range("A1:F1").Font.Size = 12			'Size to 12
	objWorksheet.Range("A1:F1").Font.Bold = True		'Bold
	objWorksheet.Range("E:E").NumberFormat = "m/d/yyyy" 'Date Format
	
	objExcel.Application.Visible = True
	objExcel.ActiveWorkbook.SaveAs strExcelPath
	
	TextFile.Close
	Set TextFile = Nothing
	Set objWorksheet = Nothing
	Set objExcel = Nothing
	
	MsgBox "Data categorized Successfully",vbInformation

End Function