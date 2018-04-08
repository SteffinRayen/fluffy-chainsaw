Function button10()

	Dim a(100),b(100),Order,ColNum,i,d
	Order = CInt(document.getElementById("Q10_1").value)
	ColNum = CInt(document.getElementById("Q10_2").value)
	fileName = document.getElementById("Q10_3").value
	Set obj = createobject("Excel.Application")
	obj.visible=True
	Set obj1=obj.Workbooks.open(fileName)
	Set obj2=obj1.Worksheets("Sheet1")
	Set celRange = obj2.Range("A1:C60")
	For i=1 To 60
		b(i)=obj2.Cells(i,ColNum).Value
	Next
	
	If Order = 1 Then
	'Ascending
		For i=1 To 60
			For j=i+1 To 60
				If (b(i) > b(j)) Then
					d =  b(i)
					b(i) = b(j)
					b(j) = d
				End If
			Next
		Next
	ElseIf Order = 2 Then
	
	'descending
		For i=1 To 60
			For j=i+1 To 60
				If (b(i) < b(j)) Then
					d =  b(i)
					b(i) = b(j)
					b(j) = d
				End If
			Next
		Next
	Else 
		If Order > 2 Or Order < 1 Then
			MsgBox "enter 1 or 2 as input"
		End If 
	End If
	For i=1 To 60
		obj2.Cells(i,ColNum).Value= b(i)
	Next
	
	obj1.Save
	obj.Quit
	Set obj1= Nothing
	Set obj2= Nothing
	Set obj=Nothing	

 
	MsgBox "Data Sorted Successfully",vbInformation
End Function