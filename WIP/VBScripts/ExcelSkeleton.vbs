Function ExcelSkeleton()

  'ReadExcel Using Search                    
  Set objExcel = Wscript.CreateObject("Excel.Application")

  filePath = InputBox("Enter the file path","Enter Value")  
  Set objWorkbook = objExcel.Workbooks.Open(filePath)   
  objExcel.visible=True
  
  rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
  colCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Columns.count  
  Msgbox("Rows    :" & rowCount)
  Msgbox("Columns :" & colCount) 
  
  a=inputbox("Enter the serial number","Search") 
   intRow = 2
   intCol = 2
    for intRow=2 to rowCount  step 1 'for (intRow = 0; intRow < rowCount; intRow++) {
       if ( CInt(a) = CInt(objExcel.Cells(intRow, 1).Value) ) then        
         for intCol=1 to colCount step 1  '(intCol = 0; intCol < colCount; intCol++) {
             c = c & "    " & (objExcel.Cells(intRow, intCol).Value) 
          next 
             sp=Split(c,";")
              b=ubound(sp)
           for i=0 to b
              Msgbox(sp(i))
           Next
       End if
          c=null
    next
  objExcel.Quit

End Function