Attribute VB_Name = "SalvaSheet"
Private Function savePath() As String
Dim fd As FileDialog
Dim path As String
Set fd = Application.FileDialog(msoFileDialogSaveAs)
path = "D:\CARTORIO\PI\PI TOTAL CARTORIO"

With fd
    .InitialFileName = path
    .FilterIndex = 1
    .Title = "Salvar como"
    If .Show <> 0 Then
        savePath = .SelectedItems(1)
    End If
    
End With

End Function

Sub copySave()
Dim ws As Worksheet, wsave As Worksheet
Dim wb As Workbook
Dim sp As String

Set ws = ThisWorkbook.Sheets("RETORNO_PI")
Dim i As Long
i = ws.UsedRange.Rows.Count

ws.Range("A2:J" & i).Copy

Set wb = Workbooks.Add

Set wsave = wb.Sheets(1)
wsave.Name = "RETORNO_PIS"
wsave.Range("A1").Select
wsave.Paste
wsave.Range("D:J").WrapText = True
wsave.Range("A:A").ColumnWidth = 16
wsave.Range("E:E").ColumnWidth = 39
wsave.Range("D:D").ColumnWidth = 39
wsave.Range("J:J").ColumnWidth = 31
wsave.Range("F:I").ColumnWidth = 17
wsave.Cells.EntireRow.AutoFit
wsave.Range("A3:J3").AutoFilter
wsave.Range("A4").Select
sp = savePath
If sp <> "" Then
    wsave.SaveAs filename:=sp
End If

End Sub


