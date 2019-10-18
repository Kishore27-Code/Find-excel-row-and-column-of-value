Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set xlVbscript = objExcel.WorkBooks.Open("c:\script.xlsx")

Set FindName = xlVbscript.Sheets(1).Usedrange.Find("implementation process ")

For each FindName in xlVbscript.Sheets(1).UsedRange
If FindName = "implementation process" Then
NameCol = FindName.Column
NameRow = FindName.Row
MsgBox(NameRow & "," & NameCol)
End If
Next
