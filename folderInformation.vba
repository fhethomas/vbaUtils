Sub folderInfoLoop()
	Dim strFolder, strLoc As String
	Dim i As Integer
	Dim outputSht As Worksheet
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	i = 2
	Set outputSht = Worksheets("Output")
	strLoc = "\\C:\Users\User\Documents\"
	strFolder = Dir(strLoc, vbDirectory)
	Do While Len(strFolder) > 0
		outputSht.Range("A" & i).Value = strFolder
		Set fold = fso.getFolder(strLoc & strFolder & "\")
		outputSht.Range("B" & i).Value = fold.DateLastModified
		i = i + 1
		strFolder = Dir
	Loop
End Sub