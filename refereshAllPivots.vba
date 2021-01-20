Sub pivotLoop()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim ctrlWkb, openWkb As Workbook
    Dim PT as PivotTable

        ' open the template workbook
        Set openWkb = Workbooks.Open(templateStr)

            For Each sht In openWkb.Worksheets
                For Each PT In sht.PivotTables       '<-- Loop all pivot tables in worksheet
                    PT.PivotCache.Refresh
                Next PT
            Next sht
        

        
        openWkb.Close saveChanges:=False

End Sub

