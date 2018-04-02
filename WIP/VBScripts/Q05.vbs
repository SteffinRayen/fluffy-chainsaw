Function button05()

    Set ObjXML = CreateObject ("Microsoft.XMLDOM")   
	filePathXML = InputBox("Enter the file path of XML","Enter Value")
    ObjXML.load(filePathXML)  
    ObjXML.async = False   

	filePathExcel = InputBox("Enter the file path of Excel","Enter Value")
    Set objExcel = CreateObject("Excel.Application")  
    Set objWorkbook =  objExcel.Workbooks.Open (filePathExcel)  
    objExcel.Application.Visible = True  
    objWorkbook.Windows(1).Visible = True  
    Set XlSheet =objWorkbook.Sheets(1)  
    XlSheet.Activate  
    
    Set CNodes = ObjXML.SelectNodes("/CATALOG/PLANT/COMMON/text()")   
    Set BNodes = ObjXML.SelectNodes("/CATALOG/PLANT/BOTANICAL/text()")   
    Set ZNodes = ObjXML.SelectNodes("/CATALOG/PLANT/ZONE/text()")   
    Set LNodes = ObjXML.SelectNodes("/CATALOG/PLANT/LIGHT/text()")   
    Set PNodes = ObjXML.SelectNodes("/CATALOG/PLANT/PRICE/text()")   
    Set ANodes = ObjXML.SelectNodes("/CATALOG/PLANT/AVAILABILITY/text()")   
    
    XlSheet.Range ("A" & 1).Value = "COMMON Name"  
    XlSheet.Range ("B" & 1).Value = "BOTANICAL Name"  
    XlSheet.Range ("C" & 1).Value = "ZONE"  
    XlSheet.Range ("D" & 1).Value = "LIGHT"  
    XlSheet.Range ("E" & 1).Value = "PRICE"  
    XlSheet.Range ("F" & 1).Value = "AVAILABILITY"  
    For i = 0 To (CNodes.Length - 1)  
        COMMON = CNodes(i).NodeValue  
        BOTANICAL = BNodes(i).NodeValue  
        ZONE = ZNodes(i).NodeValue  
        LIGHT = LNodes(i).NodeValue  
        PRICE = PNodes(i).NodeValue  
        AVAILABILITY = ANodes(i).NodeValue  

        XlSheet.Range("A" & i + 2).Value = COMMON  
        XlSheet.Range("B" & i + 2).Value = BOTANICAL 
        XlSheet.Range("C" & i + 2).Value = ZONE 
        XlSheet.Range("D" & i + 2).Value = LIGHT 
        XlSheet.Range("E" & i + 2).Value = PRICE 
        XlSheet.Range("F" & i + 2).Value = AVAILABILITY 
    Next  

    objWorkbook.save  
    objWorkbook.close  
    Set objWorkbook = Nothing  
    Set objExcel = Nothing  
    MsgBox "Data Transfered Successfully", vbInformation  

End Function