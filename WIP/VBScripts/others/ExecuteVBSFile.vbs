Function ExecuteVBSFile(fileName)

		Public fileDrive,folderpath
		folderPath= CreateObject("Scripting.FileSystemObject").BuildPath(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName("."), "GIT_WORKSPACE\fluffy-chainsaw\WIP\VBScripts")
		fileDriveSplit=Split(folderPath,":")
		fileDrive=fileDriveSplit(0) & ":"

		Dim executionMode,Command,fso1,fullPath1
		executionMode="cscript " & fileName & ".vbs"
		Set fso1=CreateObject("Scripting.FileSystemObject")
		Set WshShell = CreateObject("WScript.Shell")
		fullPath1=folderPath & fileName & ".vbs"
		If fso1.FileExists(fullPath1) Then
			Command = "cmd /K" & fileDrive & "&" & "cd " & folderPath & "&" & executionMode
			WshShell.Run Command
		Else
			MsgBox "File not found in the given path",vbOKOnly+vbCritical,"Incorrect path"
		End If
		Set WshShell = Nothing
End Function
