Sub addWksht()
    ActiveWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = "NewSheet"
End Sub 
