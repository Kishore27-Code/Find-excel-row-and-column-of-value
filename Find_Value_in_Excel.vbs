Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set xlVbscript = objExcel.WorkBooks.Open("C:\Users\ECE-Kishore27\Desktop\Probotiq_Solution\VBAscript\VBScript Basics\Vbscript.xlsx")

Set FindName = xlVbscript.Sheets(1).Usedrange.Find("Probotiq Solutions")

For each FindName in xlVbscript.Sheets(1).UsedRange
If FindName = "Probotiq Solutions" Then
NameCol = FindName.Column
NameRow = FindName.Row
MsgBox(NameRow & "," & NameCol)
End If
Next