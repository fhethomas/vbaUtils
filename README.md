# vbaUtils
This repository contains useful and frequently used functions. 

Below are a list of useful fragments of code that could prove useful.

# Modulo
i Mod 35

# Errors
' Resumes next in Loop
On Error Resume Next
' Causes break in code
On Error GoTo 0
' Goes to specified section and runs code
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

