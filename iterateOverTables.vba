Sub tblLoop()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim ctrlWkb, openWkb As Workbook

        ' open the template workbook
        Set openWkb = Workbooks.Open(templateStr)

            For Each sht In openWkb.Worksheets
                For Each tbl In sht.ListObjects
                    Debug.Print(tbl.name)
                Next tbl
            Next sht
        
        openWkb.Close saveChanges:=False

End Sub

