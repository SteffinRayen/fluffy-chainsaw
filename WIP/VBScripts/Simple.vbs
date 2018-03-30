Function Simple()

	
	Set fso = CreateObject("Scripting.FileSystemObject")
	sFolder = CreateObject("Scripting.FileSystemObject").BuildPath(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName("."), "GIT_WORKSPACE\fluffy-chainsaw\WIP\") 
	document.write ("Directory : " &sFolder& "<br />")
	
	If sFolder = "" Then
	  Msgbox ("No Folder parameter was passed")
	End If

	Dim FileNameArr()
	Set folder = fso.GetFolder(sFolder)
	Set files = folder.Files
	counter = 0
	For each folderIdx In files
		redim preserve FileNameArr(counter)
		FileNameArr(counter) = folderIdx.Name
		counter = counter + 1
		document.write(counter &" "& FileNameArr(counter - 1) &" <br />")
	Next

End Function