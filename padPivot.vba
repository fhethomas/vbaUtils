Function padPivot(piv, wkSht)
    addrStr = piv.TableRange1.Address
    topStr = Left(addrStr, InStr(1, addrStr, ":") - 1)
    bottomStr = Right(addrStr, Len(addrStr) - InStr(1, addrStr, ":"))
    wkSht.Range(bottomStr).Offset(1, 0).EntireRow.Insert Shift:=xlDown
    wkSht.Range(topStr).Offset(0, 0).EntireRow.Insert Shift:=xlUp
    padPivot = addrStr
End Function