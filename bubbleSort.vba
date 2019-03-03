
Function BubbleSort(xArr)
    Dim boolSorted As Boolean
    Dim intArrLength As Integer
    Dim i, tempInt, chgCntInt, LoopCntInt As Integer
    
    intArrLength = UBound(xArr)
    LoopCntInt = 0
    boolSorted = False
    Do While boolSorted = False
    
        chgCntInt = 0
        For i = 1 To intArrLength
        
            If xArr(i - 1) > xArr(i) Then
                tempInt = xArr(i)
                xArr(i) = xArr(i - 1)
                Arr(i - 1) = tempInt
                chgCntInt = chgCntInt + 1
            End If
        
        Next i
        
        If chgCntInt = 0 Then
            boolSorted = True
        End If
        LoopCntInt = LoopCntInt + 1
    
    Loop
    
    BubbleSort = xArr
    
End Function
