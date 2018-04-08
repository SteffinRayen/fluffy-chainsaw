Function button12()

	TextFilePath = document.getElementById("Q12_1").value
	Set TextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(TextFilePath)

	ExcelFilePath = document.getElementById("Q12_2").value
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	objExcel.Workbooks.Add
	ExcelFile = document.getElementById("Q12_3").value
	Set objWorksheet = objExcel.ActiveWorkbook.Worksheets(1)
	strExcelPath = ExcelFilePath+"\"+ExcelFile+".xlsx"
	objWorksheet.Name = "Content"
	
'Column Heading
'`````````````````````````````````	
	objWorksheet.Cells(1, 1).Value = "String" 		'Row 1 Column 1 (A)
	objWorksheet.Cells(1, 2).Value = "Integer" 		'Row 1 Column 2 (B)
	objWorksheet.Cells(1, 3).Value = "Long" 		'Row 1 Column 3 (C)
	objWorksheet.Cells(1, 4).Value = "Double" 		'Row 1 Column 4 (D)
	objWorksheet.Cells(1, 5).Value = "Date" 		'Row 1 Column 5 (E)
	objWorksheet.Cells(1, 6).Value = "Time" 		'Row 1 Column 6 (F)
	
'Read Lines & and Store into Array
'``````````````````````````````````	
	Dim aLine()
	i = 0
	Do Until TextFile.AtEndOfStream
		Redim Preserve aLine(i)
		aLine(i) = TextFile.ReadLine
		i = i + 1
	Loop
	TextFile.Close
	m=0

'Splitting Line into Array (word by Word)
'`````````````````````````````````````````
	For Each strLine in aLine
		inputText = strLine
		outputArray = Split(inputText)
	
		'This allows each word to stored in new array seperately
		For Each x in outputArray
			Redim Preserve Word(m)
			Word(m)=x
			while (Right(Word(m),1)="." Or Right(Word(m),1)="," Or Right(Word(m),1)="?" Or Right(Word(m),1)=":" Or Right(Word(m),1)=";" or Right(Word(m),1)="/")
				Word(m)=Left(Word(m), len(Word(m))-1)
			Wend
			m=m+1
		
		Next
	Next

'Regular Expressions
'````````````````````````
'Date
'`````
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
	Set rDouble = New RegExp
	With rDouble
      .Pattern    = "^(\d)*[./.](\d)*$"
      .IgnoreCase = True
      .Global     = True
	End With
  'Time 
'````````````````````````
	Set rTime = New RegExp
	With rTime			' HH			:	09		19/(1-5)(0-9)	
      .Pattern    = "^(0[0-9]|(1[0-9])|(2[0-3]))[:/.](0[0-9]|[1-5][0-9])$"
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
'````````````````````````
	StrCount=2
	IntCount=2	
	LongCount=2
	DoubleCount=2
	DateCount=2
	TimeCount=2
	PuncCount=2
	
'DataType wise
'```````````````````````````````````````````````
	for k=0 to UBound(Word)

		If(rString.Test(Word(k))) then
			objWorksheet.Cells(StrCount, 1).Value = Word(k)
			StrCount=StrCount+1
		End if

		If(rInteger.Test(Word(k))) then
			objWorksheet.Cells(IntCount, 2).Value = Word(k)
			IntCount=IntCount+1
		End If
	
		If(rLong.Test(Word(k))) then
			objWorksheet.Cells(LongCount, 3).Value = Word(k)
			LongCount=LongCount+1
		End If
	
		if(rDouble.Test(Word(k))) then
			if(not(Word(k)=".")) then
				objWorksheet.Cells(DoubleCount, 4).Value = Word(k)
				DoubleCount=DoubleCount+1
			End If
		End If
	
		if(rDate1.Test(Word(k))) then
			objWorksheet.Cells(DateCount, 5).Value = Word(k)
			DateCount=DateCount+1
		End If
	
		If(rDate2.Test(Word(k))) then
			objWorksheet.Cells(DateCount, 5).Value = Word(k)
			DateCount=DateCount+1
		End If
	
		if(rTime.Test(Word(k))) then
			objWorksheet.Cells(TimeCount, 6).Value = Word(k)
			TimeCount=TimeCount+1
		End if
	Next

	
'FORMAT
'`````````````````````````````````
	objWorksheet.Range("A1:G1").Font.Italic = True 		'Italic 
	objWorksheet.Range("A1:G1").Font.Size = 12			'Size to 12
	objWorksheet.Range("A1:G1").Font.Bold = True		'Bold
	objWorksheet.Range("E:E").NumberFormat = "m/d/yyyy" 'Date Format
	objWorksheet.Range("C:C").NumberFormat = "General"  'NumberFormat

'Saving Excel File
'`````````````````````````````````
	objExcel.Application.Visible = True
	objExcel.ActiveWorkbook.SaveAs strExcelPath
 
'Cleaning memory of Objects used for TextFile & Excel
'````````````````````````````````````````````````````
	Set TextFile = Nothing
	Set objWorksheet = Nothing
	Set objExcel = Nothing
	
	MsgBox "Data categorized Successfully",vbInformation

End Function