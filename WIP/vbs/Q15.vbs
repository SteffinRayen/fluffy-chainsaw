Function button15()

	sFolder = document.getElementById("Q15_1").value  	'path for Screenshots
	If sFolder = "" Then
	  MsgBox ("No Folder parameter was passed")
	End If
	
	filePath = document.getElementById("Q15_2").value                'path to save Excel file
	'------------------------------------------------------------------------------------------------------------------------------------------------
	
	Set objExcel = CreateObject("Excel.Application")                                 'Object creation for Excel Application
	objExcel.Application.Visible = True                                              'Making  Excel object Visible
	Set objWorkbook = objExcel.Workbooks.Add()                                       'Adding a new Workbook

	'----------------------------------------------------------------------------------------------------------------------------------------------------
	Set folder = CreateObject("Scripting.FileSystemObject").GetFolder(sFolder)       ' Creating Folder Object for Path mentioned
	Set files = folder.Files                                                         ' Declaring variable files for for all Files in that folder
	TestCase = 0
	ScreenShot = 0
	TestCaseCount = 0
	'--------------------------------------------------------------------------------------------------------------------------------------------------------
	
	
	For each folderIdx In files                                             ' Loop to Read all Files in folder one by one ,Will always read in Ascending order
		TestCase = Mid (folderIdx.Name,9,1)                                 ' To capture 9th letter of file name
		ScreenShot = Mid (folderIdx.Name, 21 ,1)                            ' To capture 21th letter of file name
		    IF TestCaseCount <> TestCase Then
			    TestCaseCount = TestCase
			    Set objWorksheet = objWorkbook.Worksheets.Add               'Adding a new Worksheet
			    objWorksheet.Name="Testcase"&TestCase                       ' Naming the sheet 
			    objWorkbook.Sheets(objWorksheet.Name).Activate              ' Activting current Worksheet
                                               
			END IF
		IF TestCase = ScreenShot Then                                      
		   i=0                                                             
		END IF
		     i=i+10                                                    
		
              objWorksheet.Cells(i,1).Select                                ' Selecting cell to insert
	             
				 With objExcel.ActiveCell.Worksheet.Pictures.Insert(sFolder&"\"&folderIdx.Name).ShapeRange     'inserting Picture
                       .Height = 200                                                                           'Setting Dimension For picture
                       .Width = 200
                 End With
                
	Next
  '-----------------------------------------------------------------------------------------------------------------------------------------------------------
	
	objWorkbook.SaveAs(filepath)                                               'Saving in Destination File Path
	objExcel.Quit                                                              ' Quitting Excel Application
	MsgBox "Screenshots sorted Successfully",vbInformation			

End Function	
