Function Directory()
	dim fso: set fso = CreateObject("Scripting.FileSystemObject")
    dim CurrentDirectory
    CurrentDirectory = fso.GetAbsolutePathName(".")
    dim Directory
    Directory = fso.BuildPath(CurrentDirectory, "Rest of the directory till your working directory\")
	Msgbox("The current directory is "& Directory)



	Set objShell = CreateObject("Wscript.Shell")
	strPath = ScriptFullName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(strPath)
	strFolder = objFSO.GetParentFolderName(objFile) 
	strPath = "explorer.exe /e," & strFolder
	objShell.Run strPath



End Function