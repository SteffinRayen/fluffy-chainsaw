Function HtmlToExcel()
                  
	Set objExcel = Wscript.CreateObject("Excel.Application")

	filePath = InputBox("Enter the file path","Enter Value")  
	Set objWorkbook = objExcel.Workbooks.Open(filePath)   
	objExcel.visible=True
    objWorkbook.Windows(1).Visible =True  
    
    set XlSheet =objWorkbook.Sheets(1)  
    XlSheet.Activate  
    Set tab=document.getElementsByTagName("table")(0)  
    mytable = document.getElementsByTagName("table")(0).rows.length  
    mytable1= document.getElementsByTagName("table")(0).rows(0).cells.length  
    For n = 0 to (mytable-1)  
        For j = 0 To (mytable1-1)  
            XlSheet.Cells (n + 1, j + 1).Value = tab.Rows(n).Cells(j).innertext   
        Next  

    Next   
    MsgBox "Data Exported Successfully",vbInformation  
    objWorkbook. Save  
    objWorkbook. Close  
    Set objWorkbook = Nothing  
    Set objExcel = Nothing  


End Function