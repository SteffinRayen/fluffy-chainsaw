Function button04()
	Input = InputBox("Enter the location of the customer detail files")
	Output = InputBox("Enter the location for exporting the excel file")
	Xcelname = InputBox("Enter a name for the output file")
	Count = Inputbox("Enter the total number of customers")
	'Bind to the Excel object
	Set objExcel = CreateObject("Excel.Application")
	 
	'Create a new workbook.
	objExcel.Workbooks.Add
	 
	'Bind to worksheet.
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	 
	'Name the worksheet
	objSheet.Name = "Customer Details"
	'Set the save location
	strExcelPath = Output+"\"+Xcelname+".xlsx"
	 
	'--------------------------------------------------------
	'Populate the worksheet with data
	'--------------------------------------------------------
	'   objSheet.Cells(row, column).Value = "Whatever"
	 
	'Add some titles to row 1
	objSheet.Cells(1, 1).Value = "Account Number" 'Row 1 Column 1 (A)
	objSheet.Cells(1, 2).Value = "Customer Name" 'Row 1 Column 2 (B)
	objSheet.Cells(1, 3).Value = "Customer ID" 'Row 1 Column 3 (C)


	Set FSO = CreateObject("Scripting.FileSystemObject")
	for textFile = 1 to Count

		Set ReadTextFile = FSO.OpenTextFile(Input+"\Q04 ("&textFile&").txt", 1)
		Do Until ReadTextFile.AtEndOfStream

		Textline = ReadTextFile.Readline()
		If Instr(Textline, "Account Number:") Then ' If textline contain string "Account Number :"
		  AccountNumber = Split(Textline, ":")(1) ' Split the textline  with ":" indicator
		End If
		If Instr(Textline, "Customer Name:") Then
		  CustomerName = Split(Textline, ":")(1)
		End If
		If Instr(Textline, "Customer ID:") Then
		   ID = Split(Textline, ":")(1)
		End If
		objSheet.Cells(textFile +1, 1).Value = AccountNumber
		objSheet.Cells(textFile +1, 2).Value = CustomerName
		objSheet.Cells(textFile +1, 3).Value = ID

		
		Loop
	next

	MsgBox "Data collated Successfully",vbInformation

	'--------------------------------------------------------
	' Save the spreadsheet and close the workbook
	'--------------------------------------------------------
	objExcel.ActiveWorkbook.SaveAs strExcelPath
	objExcel.ActiveWorkbook.Close
	 
	'Quit Excel
	objExcel.Application.Quit
	 
	'Clean Up
	Set objSheet = Nothing
	Set objExcel = Nothing



End Function