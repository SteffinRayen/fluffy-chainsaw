Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
wscript.echo sFileSelected
'------ To Open a Excel Workbook------' 
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(sFileSelected)
objExcel.Application.Visible=True 
'--- To Take screenshot of each Worksheet in the Excel Workbook ------'
dim s
For each s in objExcel.Worksheets
	s.Activate
'----------- Taking Screenshot using word object---------------------'
	Set oWordBasic=CreateObject("Word.Basic")
	oWordBasic.SendKeys"{prtsc}"
	oWordBasic.AppClose"Microsoft Word"
	Set oWordBasic=Nothing
	Wscript.Sleep 2000
	set Wshshell=CreateObject("WScript.shell")
	set shl=createobject("shell.application")
	shl.MinimizeAll
	WScript.Sleep 1000
	shl.UndoMinimizeAll
	Set shl=Nothing
	WScript.Sleep 1000
	WshShell.SendKeys "^v"
	WScript.Sleep 500
	MsgBox "Screenshot taken"
Next
i=i+1
'--- To Rename the Worksheets ------------'
newName = objExcel.Application.InputBox("Rename the Worksheets as :")
For i = 1 To objExcel.Application.Sheets.Count
	objExcel.Application.Sheets(i).Name = newName & i
Next