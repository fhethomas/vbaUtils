Function InsertArray(xValue, xArr, xPosition)
	Dim i, tempInt As Integer
	Dim insertedBool As Boolean
	insertedBool = False
	ReDim Preserve xArr(0 to Ubound(xArr) + 1)

	For i = 0 to Ubound(xArr)
		If i >= xPosition And insertedBool = False Then
			insertedBool = True
			tempInt = xArr(i)
			xArr(i) = xValue
		ElseIf i >= xPosition Then
			xValue = tempInt
			tempInt = xArr(i)
			xArr(i) = xValue
		End If
	Next
	
	InsertArray = xArr
End Function