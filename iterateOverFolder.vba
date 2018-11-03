Sub folderLoop()
    Dim directory As String, fName As String
    Dim wkbk As Workbook    
    Application.ScreenUpdating = False
    directory = "C:\Users\"
    ' set file name to Dir
    fName = Dir(directory)
    While fName <> ""
        ' Check to ensure file is an .xl type, but not the file itself
        If InStr(1, fName, ".xl") <> 0 And fName <> ThisWorkbook.Name Then
            Set wkbk = Workbooks.Open(directory & fName, UpdateLinks:=False)
            Debug.Print ("Found file: " & wkbk.Name)
            wkbk.Close SaveChanges:=False
        End If
        ' reset file name to Dir to iterate over files
        fName = Dir
    Wend
Application.ScreenUpdating = True
End Sub
