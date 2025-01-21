Sub formatPartString(wkRng, targetStr)
    ' This sub loops through a range - finds a string and formats it
    ' *** Inputs ***
    ' wkRng - a range of cells
    ' targetStr - the string you want to make red and bold
    
    Dim cellRng As Range
    Dim startPosition, formatLength As Long
    Dim rowLength, iLong As Long
    Dim targetLength As Integer
    Dim iterationCounter As Integer
    Dim exampleStr As String

    rowLength = wkRng.Rows.Count
    targetLength = Len(targetStr)


    ' loop through the workrange
    For Each cellRng In wkRng
        ' make it uppercase & remove grammar
        exampleStr = UCase(cellRng.Value)
        exampleStr = removeGrammar(exampleStr)
        
        ' we've got a max iteration set at 300 incase while loop goes too far...
        iterationCounter = 0
        ' set first postion
        startPosition = InStr(1, exampleStr, targetStr, vbTextCompare)

        Do While startPosition > 0
            
            iterationCounter = iterationCounter + 1
            ' Format the cells
            With cellRng.Characters(Start:=startPosition, Length:=targetLength).Font
                .Bold = True ' Make text bold
                .Size = 14 ' Change font size
                .Color = RGB(255, 0, 0) ' Change text color (optional: red)
            End With
            ' Find the next start position
            startPosition = InStr(startPosition + targetLength, exampleStr, targetStr, vbTextCompare)
            If iterationCounter > 300 Then
                Exit Do
            End If
        Loop
    Next

End Sub