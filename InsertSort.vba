
Function InsertSort(xArr)

    Dim boolInserted As Boolean
    Dim newArr() As Integer
    ReDim newArr(0 To 1)
    Dim i, j As Integer
    
    If xArr(0) < xArr(1) Then
    
        newArr(0) = xArr(0)
        newArr(1) = xArr(1)
        
    Else
        newArr(1) = xArr(0)
        newArr(0) = xArr(1)
    
    End If
    
    For i = 0 To UBound(xArr)
        boolInserted = False
        
        If i >= 2 Then
            
            For j = 0 To UBound(newArr)
            
                If xArr(i) <= newArr(j) Then
                    newArr = insertArray(xArr(i), newArr, j)
                    boolInserted = True
                    Exit For
                End If
            
            Next j
            If boolInserted = False Then
                newArr = insertArray(xArr(i), newArr, UBound(newArr) + 1)
            End If
        End If
    Next i
    InsertSort = newArr
End Function
