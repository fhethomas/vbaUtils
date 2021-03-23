Sub matchColumns()
    ' Match the columns from Sheet 2 to Sheet1 by title
    ' Copy all the data into Sheet1 from Sheet2
    ' If columns in Sheet2 are in a different order to Sheet1 it puts the data into Sheet1 in the correct order
    Dim wkSht1, wkSht2 As Worksheet
    
    Set wkSht1 = Worksheets(1)
    Set wkSht2 = Worksheets(2)
    
    Dim titleStr As String
    'Dim colFoundBool As Boolean
    Dim colCount, rowCount, colCheckCount, colCountSht1 As Long
    
    Dim colInt, colCheckInt, rowInt As Integer
    
    Dim workrange As Range
    
    Set workrange = Range(wkSht2.Range("A1"), wkSht2.Range("A1").End(xlToRight))
    colCount = workrange.Count
    Set workrange = Range(wkSht1.Range("A1"), wkSht1.Range("A1").End(xlToRight))
    Set colCountSht1 = workrange.Count
    
    Set workrange = wkSht2.UsedRange
    rowCount = workrange.Rows.Count
    
    For colInt = 1 To colCount
        colFoundBool = False
        titleStr = wkSht2.Cells(1, colInt).Value
        'Debug.Print (titleStr)
        For colCheckInt = 1 To colCountSht1
            If titleStr = wkSht1.Cells(1, colCheckInt).Value Then
                'colFoundBool = True
                Range(wkSht1.Cells(2, colCheckInt), wkSht1.Cells(rowCount, colCheckInt)).Value = Range(wkSht2.Cells(2, colInt), wkSht2.Cells(rowCount, colInt)).Value
                Exit For
            End If
        Next colCheckInt
    Next colInt
    

End Sub
