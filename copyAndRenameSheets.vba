Sub genShts()

    Dim wkSht, newSht, lastSht As Worksheet
    
    Dim nameArr As Variant
    
    nameArr = Array("test1","test2","test3")
    
    Set wkSht = Worksheets("Target Sheet")
    
    Dim i, iLgth As Integer
    
    iLgth = UBound(nameArr)
    Set lastSht = Worksheets(Worksheets.Count)
    
    For i = 0 To iLgth
        
        wkSht.Copy After:=lastSht
        Set lastSht = Worksheets(Worksheets.Count)
        lastSht.Name = nameArr(i)
        Debug.Print nameArr(i)
    
    Next i

End Sub