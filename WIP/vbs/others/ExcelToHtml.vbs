Function ExcelToHtml ()  

  Set objExcel = Wscript.CreateObject("Excel.Application")

  filePath = InputBox("Enter the file path","Enter Value")  
  Set objWorkbook = objExcel.Workbooks.Open(filePath)   
  objExcel.visible=True
    objWorkbook.Windows(1).Visible =True  
  Set XlSheet =objWorkbook.Sheets (1)  
  XlSheet.Activate  
  iRow = 1  
  With objExcel  
      Do while .Cells (iRow, 1).value <> ""  
          .Cells (iRow, 1).activate  
          iRow = iRow + 1  
      Loop  
         .Cells (iRow, 1).value=Document.GetElementsByName ("simpleText") (0).Value 
         MsgBox "Data Added Successfully”, vbinformation  

         Document.GetElementsByName ("simpleText") (0).Value="" 
   End With  
   ObjWorkbook. Save  
   ObjWorkbook. Close  
   Set objWorkbook = Nothing  
   Set objExcel = Nothing  
End Function   