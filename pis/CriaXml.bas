Attribute VB_Name = "CriaXml"
Option Explicit
    
Public Sub criarXml()
Dim str As String
Dim wb As Workbook
Dim objStream As ADODB.Stream
Dim ws1 As Worksheet, ws2 As Worksheet
Dim Flag As Integer, Maximo As Integer, j As Integer, idx As Integer
Dim i As Long

If apagaArquivos Then
    MsgBox "Programa parou!", vbDefaultButton1, "Parada"
    Exit Sub
End If

Set objStream = New ADODB.Stream
    objStream.Charset = "iso-8859-1"
  
Set wb = ThisWorkbook

Set ws1 = wb.Sheets("RELTEMP")
Set ws2 = wb.Sheets("campos")

Flag = 0
Maximo = Range("S1")
idx = 1
str = ""
For i = 3 To ws1.UsedRange.Rows.Count
    If Flag = 0 Then
        objStream.Open
        str = "<?xml version=""1.0"" encoding=""iso-8859-1""?>" & vbCrLf & "<pis>" & vbCrLf
    End If
    
    Flag = Flag + 1
        
    str = str & "<pi id=""" & Flag & """>" & vbCrLf
    
    For j = 1 To ws2.UsedRange.Rows.Count
            str = str & "   <" & ws2.Cells(j, 1) & ">" & _
            WorksheetFunction.Trim(ws1.Cells(i, j)) & "</" & ws2.Cells(j, 1) & ">" & vbCrLf
    Next j

    str = str & "</pi>" & vbCrLf
    If Flag = Maximo Then
        str = str & "</pis>"
        objStream.WriteText str
        objStream.SaveToFile wb.path & "\plan_pedido_" & idx & ".xml", 2
        idx = idx + 1
        Flag = 0
        objStream.Close
    ElseIf i = ws1.UsedRange.Rows.Count Then
        str = str & "</pis>"
        objStream.WriteText str
        objStream.SaveToFile wb.path & "\plan_pedido_" & idx & ".xml", 2
        objStream.Close
    End If
Next i
End Sub


Public Sub Limpar()
Dim t As Long
Dim ws As Worksheet
Dim r As Long
Set ws = ThisWorkbook.Sheets("RELTEMP")
r = ws.UsedRange.Rows.Count

If r > 2 Then ws.Rows("3:" & r).delete
ws.Range("a3").Select
End Sub


Private Function apagaArquivos() As Boolean
Dim resposta As Integer

On Error GoTo erro
resposta = MsgBox("Deseja deletar arquivos?", vbYesNo)
If resposta = 6 Then
    Kill ThisWorkbook.path & "\plan_pedido*"
    MsgBox "Arquivos deletados com sucesso!"
    apagaArquivos = False
    Exit Function
End If

apagaArquivos = True
    
Exit Function
erro:
apagaArquivos = False
'MsgBox Err.Description
End Function
