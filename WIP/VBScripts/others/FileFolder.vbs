Function FileFolder()

	dim tag: set tag = document.getElementById("simpleText")
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	sFolder = CreateObject("Scripting.FileSystemObject").BuildPath(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName("."), "GIT_WORKSPACE\fluffy-chainsaw\WIP\VBScripts") 
	tag.InnerHtml = ("Directory : " &sFolder& "<br />")
	
	If sFolder = "" Then
	  tag.InnerHtml ("No Folder parameter was passed")
	End If


	'Get file names and display using InnerHtml
	Dim FileNameArr()
	Set folder = fso.GetFolder(sFolder)
	Set files = folder.Files
	counter = 0
	For each folderIdx In files
		redim preserve FileNameArr(counter)
		FileNameArr(counter) = folderIdx.Name
		counter = counter + 1
		tag.InnerHtml = (tag.InnerHtml & counter &" "& FileNameArr(counter - 1) &" <br>")
	Next

End Function