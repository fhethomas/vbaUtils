# VBA Utilities
This repository contains useful and frequently used functions. 

Below are a list of useful fragments of code that could prove useful.

# Get Username
```
strName =Environ("username")
```

# Modulo
```
i Mod 35
```

# Errors

Resumes next in Loop 

On Error Resume Next

Causes break in code

On Error GoTo 0

Goes to specified section and runs code

On Error GoTo alertMsg

alertMsg:
	Msgbox(“Alert!”)

# Set Workbook
Set activeWkb=ThisWorkbook
  
# Status bar updating
Application.DisplayStatusBar=True
Application.StatusBar = x & “%”
Application.StatusBar = False

# Create New Workbook
Workbooks.Add

# Worksheet Function
Debug.Print ("length of columns: " & Application.WorksheetFunction.CountA(dataSht.Range("A:A")))

# Count worksheets
MsgBox Worksheets.Count

# Select Multiple Ranges
Range(“A1:A2,B2:C4”).value=10

# Cells in range:
Range(Cells(1,1),Cells(4,1))=5

# EmptyClipboard
Application.CutCopyMode=False

# Clear Contents
Workrange.ClearContents

# Filter – returns an array of elements in the filtered array that match
Arr_tem=Filter(Original_array, Some_variable_to_look_up)

# Join – join strings from array
Join(Array_temp,”+”)

# Offset
Range(“A1”).offset(row,column).value=”offsetRow”

# Split (text,delimeter)
Split(“A brand new string”,” “)

# Set cell formulas & filldown
Wksht.range(“A1”).formula=”=Sum(B1:D1)”
Wksht.range(“A1:A” & 50).filldown

# Yes/No
Dim Answ As Integer
Answ=MsgBox(“Continue?”,vbYesNo)


# Find Substring
If InStr(1, Range("A1”).Value, "Chicago") <> 0 Then
    Debug.print("Range A1 contains Chicago")
End If

# Put a break into your code
Stop

# Frequently used Formulas
Non-Data Analysis Regression
Select 3 cells horizontally and press ctrl+Shift+Enter – gives x2+x+Constant
=Linest(y-range,xrange^{1,2},TRUE,FALSE)
Count of Unique Values
=Sumproduct(1/countif(a2:a5,a2:a5))
Weekday
=text(A1,”dddd”)   - Where A1=01/01/2017
Add a Month
=Edate(oldDate,12)
Array Formula
{=Max(If(A2:A4=”Dog”,B2:B4))}
Indirect Formula
=indirect(“Calculation!A1”)
File name
=cell(“filename”)
Get worksheet name
=replace(cell(“filename”),1,find(“]”,cell(“filename”)),””)
Date Difference
=DateDif(a1,a2,”m”) – can also be d/m/y
Year Frac
=yearfrac(a1,a2)
Errors?
=iferror(A1,”There’s an error”)
Format text
=text(a1,”00”)
Classic index match
=index(D:D,match(“Cat”,”A:A,0))
Substitute
=substitute(A1,”ReplacementString”,A2)
Data Analysis
1.	Click the File tab, click Options, and then click the Add-Ins category.
2.	In the Manage box, select Excel Add-ins and then click Go.
3.	In the Add-Ins available box, select the Analysis ToolPak check box, and then click OK.

