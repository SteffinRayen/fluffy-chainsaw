Function button05()

    Set ObjXML = CreateObject ("Microsoft.XMLDOM")   
	filePathXML = document.getElementById("Q05_1").value
    ObjXML.load(filePathXML)  
    ObjXML.async = False   

	filePathExcel = document.getElementById("Q05_2").value
    Set objExcel = CreateObject("Excel.Application")  
    Set objWorkbook =  objExcel.Workbooks.Open (filePathExcel)  
    objExcel.Application.Visible = True  
    objWorkbook.Windows(1).Visible = True  
    Set objWorksheet =objWorkbook.Sheets(1)  
    objWorksheet.Activate  
    
    Set CNodes = ObjXML.SelectNodes("/CATALOG/PLANT/COMMON/text()")   
    Set BNodes = ObjXML.SelectNodes("/CATALOG/PLANT/BOTANICAL/text()")   
    Set ZNodes = ObjXML.SelectNodes("/CATALOG/PLANT/ZONE/text()")   
    Set LNodes = ObjXML.SelectNodes("/CATALOG/PLANT/LIGHT/text()")   
    Set PNodes = ObjXML.SelectNodes("/CATALOG/PLANT/PRICE/text()")   
    Set ANodes = ObjXML.SelectNodes("/CATALOG/PLANT/AVAILABILITY/text()")   
    
    objWorksheet.Range ("A" & 1).Value = "COMMON Name"  
    objWorksheet.Range ("B" & 1).Value = "BOTANICAL Name"  
    objWorksheet.Range ("C" & 1).Value = "ZONE"  
    objWorksheet.Range ("D" & 1).Value = "LIGHT"  
    objWorksheet.Range ("E" & 1).Value = "PRICE"  
    objWorksheet.Range ("F" & 1).Value = "AVAILABILITY"  
    For i = 0 To (CNodes.Length - 1)
        objWorksheet.Range("A" & i + 2).Value = CNodes(i).NodeValue   
        objWorksheet.Range("B" & i + 2).Value = BNodes(i).NodeValue  
        objWorksheet.Range("C" & i + 2).Value = ZNodes(i).NodeValue  
        objWorksheet.Range("D" & i + 2).Value = LNodes(i).NodeValue  
        objWorksheet.Range("E" & i + 2).Value = PNodes(i).NodeValue  
        objWorksheet.Range("F" & i + 2).Value = ANodes(i).NodeValue  
    Next  

    objWorkbook.save  
    objWorkbook.close  
    Set objWorkbook = Nothing  
    Set objExcel = Nothing 

    MsgBox "Data Transfered Successfully", vbInformation  

End Function