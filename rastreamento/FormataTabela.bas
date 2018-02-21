Attribute VB_Name = "FormataTabela"
Public Sub insereTab(name As String)

Dim wb As Workbook
Set wb = ThisWorkbook

Application.DisplayAlerts = False
For i = wb.Sheets.Count To 2 Step -1
    wb.Sheets(i).Delete
Next i
Application.DisplayAlerts = True

wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)
wb.Sheets(wb.Sheets.Count).name = name

formata (name)

End Sub


Public Sub limpaTab()
If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
End If
i = ActiveSheet.UsedRange.Rows.Count
If i > 2 Then Rows("3:" & i).Delete
Range("A3").Select

End Sub

Sub formata(nometable As String)
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(nometable)
ws.Range("A1") = "RASTREADO " & DateValue(Now)
ws.Range("A1:I1").Merge

With ws.Range("a1:i1")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .MergeCells = True
    .Font.Bold = True
End With

ws.Range("A1").Interior.ThemeColor = xlThemeColorAccent2

ws.Cells(2, 1) = "Nº DO OBJETO"
ws.Cells(2, 2) = "SITUAÇÃO"
ws.Cells(2, 3) = "DATA"
ws.Cells(2, 4) = "CIDADE"
ws.Cells(2, 5) = "UF"
ws.Cells(2, 6) = "UNIDADE"
ws.Cells(2, 7) = "DR"
ws.Cells(2, 8) = "STATUS"
ws.Cells(2, 9) = "TIPO"

ws.Range("a2:i2").AutoFilter


End Sub

