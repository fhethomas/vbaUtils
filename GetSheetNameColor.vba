Sub getShtNamesAndColor()

    Dim wksht As Worksheet
    
    For Each wksht In Worksheets
            ' print the worksheet name and also the color index
            Debug.Print wksht.Name
            Debug.Print wksht.Tab.ColorIndex
        
    Next wksht

End Sub
