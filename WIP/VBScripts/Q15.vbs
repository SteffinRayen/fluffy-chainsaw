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
	Set objWorksheet = objWorkbook.Worksheets(1)

	Set folder = CreateObject("Scripting.FileSystemObject").GetFolder(sFolder)
	Set files = folder.Files
	TestCase = 0
	ScreenShot = 0
	For each folderIdx In files
		'Will always read in Ascending order
		
		TestCase = Mid (folderIdx.Name,9,1)
		ScreenShot = Mid (folderIdx.Name, 21 ,1)

		Worksheets(TestCase).Activate 

		With ActiveSheet.Pictures.insert(sFolder&"\"&folderIdx.Name)
	        With .ShapeRange
	            .LockAspectRatio = msoTrue
	            .Width = 50
	            .Height = 70
	        End With
	        .Left = ActiveSheet.Range("A" & ScreenShot).Left
	        .Top = ActiveSheet.Range("A" & ScreenShot).Top
	        .Placement = 1
	        .PrintObject = True
	    End With

		tag.InnerHtml = (tag.InnerHtml & TestCase &" "& ScreenShot &" "& folderIdx.Name &" <br>")
	Next

	
	objExcel.Quit
	MsgBox "Screenshots sorted Successfully",vbInformation


End Function