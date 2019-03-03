
Sub Quicksort(xArr As Variant, arrLbound As Long, arrUbound As Long)

    Dim pivotVal As Variant
    Dim vSwap As Variant
    Dim tmpLow, tmpHi As Long
    
    tmpLow = arrLbound
    tmpHi = arrUbound
    pivotVal = xArr((arrLbound + arrUbound) \ 2)
    
    While (tmpLow <= tmpHi)
        While (xArr(tmpLow) < pivotVal And tmpLow < arrUbound)
            tmpLow = tmpLow + 1
        Wend
        While (pivotVal < xArr(tmpHi) And tmpHi > arrLbound)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            vSwap = xArr(tmpLow)
            xArr(tmpLow) = xArr(tmpHi)
            xArr(tmpHi) = vSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (arrLbound < tmpHi) Then Quicksort xArr, arrLbound, tmpHi
    
    If (tmpow < arrUbound) Then Quicksort xArr, tmpLow, arrUbound
    
End Sub
