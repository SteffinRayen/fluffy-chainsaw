Function button09()
	Dim objFSO, colFiles, objFile, strDestFolder, objNewestFile,filename, countofA, countofB, countofC, countofD, countofE, countofF, countofG, countofH, countofI, countofJ, countofK, countofL, countofM, countofN, countofO, countofP, countofQ, countofR, countofS, countofT, countofU, countofV, countofW, countofX, countofY, countofZ, countofFOLDER, strFolder
	countofA=0
	countofB=0
	countofC=0
	countofD=0
	countofE=0
	countofF=0
	countofG=0
	countofH=0
	countofI=0
	countofJ=0
	countofK=0
	countofL=0
	countofM=0
	countofN=0
	countofO=0
	countofP=0
	countofQ=0
	countofR=0
	countofS=0
	countofT=0
	countofU=0
	countofV=0
	countofW=0
	countofX=0
	countofY=0
	countofZ=0
	countofFOLDER=0
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set colFiles = objFSO.GetFolder("F:\Reader")
	filename = ("F:\A")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\A_Ref\"
	Else
		objFSO.CreateFolder("F:\A")
		strDestFolder = "F:\A\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "a" or Left(objFile.Name, 1)="A" Then
			countofA=countofA + 1
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\B")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\B_Ref\"
	Else
		objFSO.CreateFolder("F:\B")
		strDestFolder = "F:\B\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "b" or Left(objFile.Name, 1)="B" Then
			countofB=countofB + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\C")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\C_Ref\"
	Else
		objFSO.CreateFolder("F:\C")
		strDestFolder = "F:\C\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "c" or Left(objFile.Name, 1)="C" Then
			countofC=countofC + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\D")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\D_Ref\"
	Else
		objFSO.CreateFolder("F:\D")
		strDestFolder = "F:\D\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "d" or Left(objFile.Name, 1)="D" Then
			countofD=countofD + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\E")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\E_Ref\"
	Else
		objFSO.CreateFolder("F:\E")
		strDestFolder = "F:\E\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "e" or Left(objFile.Name, 1)="E" Then
			countofE=countofE + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\F")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\F_Ref\"
	Else
		objFSO.CreateFolder("F:\F")
		strDestFolder = "F:\F\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "f" or Left(objFile.Name, 1)="F" Then
			countofF=countofF + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\G")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\G_Ref\"
	Else
		objFSO.CreateFolder("F:\G")
		strDestFolder = "F:\G\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "g" or Left(objFile.Name, 1)="G" Then
			countofG=countofG + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\H")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\H_Ref\"
	Else
		objFSO.CreateFolder("F:\H")
		strDestFolder = "F:\H\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "h" or Left(objFile.Name, 1)="H" Then
			countofH=countofH + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\I")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\I_Ref\"
	Else
		objFSO.CreateFolder("F:\I")
		strDestFolder = "F:\I\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "i" or Left(objFile.Name, 1)="I" Then 
			countofI=countofI + 1                  
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\J")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\J_Ref\"
	Else
		objFSO.CreateFolder("F:\J")
		strDestFolder = "F:\J\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "j" or Left(objFile.Name, 1)="J" Then
			countofJ=countofJ + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\K")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\K_Ref\"
	Else
		objFSO.CreateFolder("F:\K")
		strDestFolder = "F:\K\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "k" or Left(objFile.Name, 1)="K" Then
			countofK=countofK + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\L")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\L_Ref\"
	Else
		objFSO.CreateFolder("F:\L")
		strDestFolder = "F:\L\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "l" or Left(objFile.Name, 1)="L" Then
			countofL=countofL + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\M")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\M_Ref\"
	Else
		objFSO.CreateFolder("F:\M")
		strDestFolder = "F:\M\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "m" or Left(objFile.Name, 1)="M" Then
			countofM=countofM + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\N")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\N_Ref\"
	Else
		objFSO.CreateFolder("F:\N")
		strDestFolder = "F:\N\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "n" or Left(objFile.Name, 1)="N" Then
			countofN=countofN + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\O")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\O_Ref\"
	Else
		objFSO.CreateFolder("F:\O")
		strDestFolder = "F:\O\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "o" or Left(objFile.Name, 1)="O" Then
			countofO=countofO + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\P")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\P_Ref\"
	Else
		objFSO.CreateFolder("F:\P")
		strDestFolder = "F:\P\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "p" or Left(objFile.Name, 1)="P" Then
			countofP=countofP + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\Q")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\Q_Ref\"
	Else
		objFSO.CreateFolder("F:\Q")
		strDestFolder = "F:\Q\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "q" or Left(objFile.Name, 1)="Q" Then
			countofQ=countofQ + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\R")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\R_Ref\"
	Else
		objFSO.CreateFolder("F:\R")
		strDestFolder = "F:\R\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "r" or Left(objFile.Name, 1)="R" Then
			countofR=countofR + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\S")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\S_Ref\"
	Else
		objFSO.CreateFolder("F:\S")
		strDestFolder = "F:\S\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "s" or Left(objFile.Name, 1)="S" Then
			countofS=countofS + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\T")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\T_Ref\"
	Else
		objFSO.CreateFolder("F:\T")
		strDestFolder = "F:\T\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "t" or Left(objFile.Name, 1)="T" Then
			countofT=countofT + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\U")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\U_Ref\"
	Else
		objFSO.CreateFolder("F:\U")
		strDestFolder = "F:\U\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "u" or Left(objFile.Name, 1)="U" Then
			countofU=countofU + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\V")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\V_Ref\"
	Else
		objFSO.CreateFolder("F:\V")
		strDestFolder = "F:\V\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "v" or Left(objFile.Name, 1)="V" Then
			countofV=countofV + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\W")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\W_Ref\"
	Else
		objFSO.CreateFolder("F:\W")
		strDestFolder = "F:\W\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "w" or Left(objFile.Name, 1)="W" Then
			countofW=countofW + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\X")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\X_Ref\"
	Else
		objFSO.CreateFolder("F:\X")
		strDestFolder = "F:\X\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "x" or Left(objFile.Name, 1)="X" Then
			countofX=countofX + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\Y")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\Y_Ref\"
	Else
		objFSO.CreateFolder("F:\Y")
		strDestFolder = "F:\Y\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "y" or Left(objFile.Name, 1)="Y" Then
			countofY=countofY + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True
		End If
	Next
	filename = ("F:\Z")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		strDestFolder = "F:\Z_Ref\"
	Else
		objFSO.CreateFolder("F:\Z")
		strDestFolder = "F:\Z\"
	END If
	For each objFile In colFiles.Files
		If Left(objFile.Name, 1) = "z" or Left(objFile.Name, 1)="Z" Then
			countofZ=countofZ + 1                   
			objFSO.CopyFile objFile.Path,strDestFolder,True 
		End If
	Next
	filename = ("F:\folder")
	folderexists = objFSO.FolderExists(filename)
	If objFSO.FolderExists(filename) Then
		newname = filename & String(1, "_")
		newname = newname & String(1, "R")
		newname = newname & String(1, "e")
		newname = newname & String(1, "f")
		objFSO.FolderExists(newname)
		objFSO.CreateFolder(newname)
		objFSO.copyFolder"F:\Reader\*","F:\folder_Ref"
		MsgBox "countof FOLDER=" & objFSO.GetFolder("F:\folder_Ref").Subfolders.Count
	Else
		objFSO.CreateFolder("F:\folder")
		objFSO.copyFolder"F:\Reader\*","F:\folder"
		MsgBox "countof FOLDER=" & objFSO.GetFolder("F:\folder").Subfolders.Count
	END If
	MsgBox "A= " & countofA	& vbCrlf & "B= " & countofB	& vbCrlf & "C= " & countofC	& vbCrlf & "D= " & countofD	& vbCrlf & "E= " & countofE	& vbCrlf & "F= " & countofF	& vbCrlf & "G= " & countofG	& vbCrlf & "H= " & countofH	& vbCrlf & "I= " & countofI	& vbCrlf & "J= " & countofJ	& vbCrlf & "K= " & countofK	& vbCrlf & "L= " & countofL	& vbCrlf & "M= " & countofM	& vbCrlf & "N= " & countofN	& vbCrlf & "O= " & countofO	& vbCrlf & "P= " & countofP	& vbCrlf & "Q= " & countofQ	& vbCrlf & "R= " & countofR	& vbCrlf & "S= " & countofS	& vbCrlf & "T= " & countofT	& vbCrlf & "U= " & countofU	& vbCrlf & "V= " & countofV	& vbCrlf & "W= " & countofW	& vbCrlf & "X= " & countofX	& vbCrlf & "Y= " & countofY	& vbCrlf & "Z= " & countofZ	& vbCrlf & "Done."
End Function