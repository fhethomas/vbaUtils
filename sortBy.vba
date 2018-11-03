Sub sortWksht()
	Dim wksht As Worksheet
	Set wksht = Worksheets("Data")
	Dim sortRange As Range
	Set sortRange=Range(wksht.Range(“A1”),wksht.Range(“A1”).end(xlDown))
	' clear existing sort
	wksht.Sort.SortFields.Clear

	' Sort by data in columns P and O
	wkSht.Sort.SortFields.Add Key:=Range(“P2:P” & sortRange.count),SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
	wkSht.Sort.SortFields.Add Key:=Range(“O2:O” & sortRange.count),SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
	With wkSht.Sort
		.SetRange Range(“A1:U” & sortRange.count)
		.Header=xlYes
		.MatchCase=False
		.Orientation=xlTopToBottom
		.SortMethod=xlPinYin
		.Apply
	End With
End Sub