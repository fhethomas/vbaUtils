Sub IterateOverTable()
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim ctrlWkb, openWkb As Workbook
    Dim ctrlSht As Worksheet
    Dim schoolFacRange, workrange, templateFile, cellRng As Range
    Set ctrlWkb = ThisWorkbook
    Dim facStr, schoolStr, studentFactStr, studentSchoolStr, fileNameStr, folderStr, templateStr As String
    Dim PT As PivotTable
    Set ctrlSht = ctrlWkb.Worksheets("CtrlSht")
    Dim iRow, x As Long
    Dim rowH As Range
    Dim facCOl, schoolCol, i  As Integer
    Dim delRng As Range
    
    Application.DisplayAlerts = False
    ' Define your Excel Doc
    templateStr="C:\My Docs\My Excel.xlsx"
    
    ' we're deleting everything that's not equal to this school
    schoolStr = "MY SCHOOL!"

        ' open the template workbook
        Set openWkb = Workbooks.Open(templateStr)
            ' iterate over each sheet and table and delete values that don't match School
            For Each sht In openWkb.Worksheets
                For Each tbl In sht.ListObjects
                   
                        ' find the column with school as headers
                        i = 1
                        schoolCol = 0
                        For Each rowH In tbl.HeaderRowRange
                            If rowH.Value = "School" Then
                                schoolCol = i
                            ElseIf schoolCol > 0 Then
                                Exit For
                            End If
                            i = i + 1
                        Next rowH
                        
                        Debug.Print ("*****" & tbl.Name & "*****")
                        
                        ' Filter on not My School!
                        tbl.Range.AutoFilter field:=schoolCol, Criteria1:="<>" & schoolStr
			' Create a range that equals visible cells
                        Set delRng = tbl.Range.Offset(1).Resize(tbl.Range.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
                        ' Clear the filter
                        tbl.Range.AutoFilter
			' Delete the range
                        delRng.Delete                    
                Next tbl
            Next sht

            ' refresh pivot tables
            For Each sht In openWkb.Worksheets
                 On Error Resume Next
                For Each PT In sht.PivotTables        '<~~ Loop all pivot tables in worksheet
                    PT.PivotCache.Refresh
                Next PT
            Next sht
        openWkb.SaveAs folderStr & "\" & fileNameStr
        openWkb.Close saveChanges:=False

    
    
    Application.DisplayAlerts = True

End Sub