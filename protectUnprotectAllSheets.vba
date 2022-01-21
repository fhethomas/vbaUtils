Sub protectAll()

    Dim wkSht As Worksheet
    
    For Each wkSht In Worksheets
        wkSht.Protect Password:="Test"
    Next wkSht

End Sub
Sub unprotectAll()

    Dim wkSht As Worksheet
    
    For Each wkSht In Worksheets
        wkSht.Unprotect Password:="Test"
    Next wkSht

End Sub
