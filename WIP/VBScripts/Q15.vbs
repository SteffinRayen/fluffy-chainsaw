Function button15()

	dim tag: set tag = document.getElementById("simpleText")
	sFolder = InputBox("Enter the folder path to the screen shots","Enter Value")

	If sFolder = "" Then
	  tag.InnerHtml ("No Folder parameter was passed")
	End If

	'Open a workbook
	filePath = InputBox("Enter the file path to .xlsx","Enter Value")
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.Visible = True
	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	'Select a worksheet
	Set objWorksheet1 = objWorkbook.Worksheets(1)
	Set objWorksheet2 = objWorkbook.Worksheets(2)
	Set objWorksheet3 = objWorkbook.Worksheets(3)
	Set objWorksheet4 = objWorkbook.Worksheets(4)

	Set folder = CreateObject("Scripting.FileSystemObject").GetFolder(sFolder)
	Set files = folder.Files
	TestCase = 0
	ScreenShot = 0
	Counter = 0
	For each folderIdx In files
		'Will always read in Ascending order
		
		TestCase = Mid (folderIdx.Name,9,1)
		ScreenShot = Mid (folderIdx.Name, 21 ,1)

		SELECT case TestCase
			CASE 1
				objWorksheet1.Cells(ScreenShot,1).Value = folderIdx.Name
				'Add Image in this cell
				'Image addition not working in Excel 10 :(
				objWorksheet1.Cells(ScreenShot,2).Value = Trim(sFolder&"\"&folderIdx.Name)
				
				
			CASE 2
				objWorksheet2.Cells(ScreenShot,1).Value = folderIdx.Name
				objWorksheet2.Cells(ScreenShot,2).Value = Trim(sFolder&"\"&folderIdx.Name)
				
			CASE 3
				objWorksheet3.Cells(ScreenShot,1).Value = folderIdx.Name
				objWorksheet3.Cells(ScreenShot,2).Value = Trim(sFolder&"\"&folderIdx.Name)
				
			CASE 4
				objWorksheet4.Cells(ScreenShot,1).Value = folderIdx.Name
				objWorksheet4.Cells(ScreenShot,2).Value = Trim(sFolder&"\"&folderIdx.Name)
				

		END SELECT
		tag.InnerHtml = (tag.InnerHtml & TestCase &" "& ScreenShot &" "& folderIdx.Name &" A"&ScreenShot + Counter*2 &":B"&ScreenShot + Counter*2 + 1 &" <br>")
		Counter = Counter + 1
	Next

	'Save the workbook,
	objWorkbook.Save

	objExcel.Quit
	MsgBox "Screenshots sorted Successfully",vbInformation


End Function

