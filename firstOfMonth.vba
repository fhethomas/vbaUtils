Function firstOfMonth(theDate)
    
    firstOfMonth = DateAdd("d", -(DatePart("d", theDate) - 1), theDate)

End Function