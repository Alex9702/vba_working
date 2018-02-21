Attribute VB_Name = "RetornoPi"
Option Explicit

Sub xmlToSheet()
Dim path As String, filename As String

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("RETORNO_PI")

Dim r  As Long

r = ws.UsedRange.Rows.Count + 1

Dim xDoc As MSXML2.DOMDocument60
Dim nodelist As MSXML2.IXMLDOMNodeList, n As IXMLDOMNode

Set xDoc = New MSXML2.DOMDocument60
xDoc.async = False
xDoc.validateOnParse = False
path = getPath
If path = "" Then Exit Sub

If xDoc.Load(path) Then
    Set nodelist = xDoc.getElementsByTagName("pi")
    
    For Each n In nodelist
        ws.Cells(r, 1) = n.ChildNodes(0).Text 'OBJETO
        ws.Cells(r, 2) = n.ChildNodes(1).Text 'COD PI
        ws.Cells(r, 3) = n.ChildNodes(2).Text 'COD ERRO
        If ws.Cells(r, 3) = 900 Then
            ws.Cells(r, 4) = mensagens 'MENS ERRO
        Else
            ws.Cells(r, 4) = "" 'MENS ERRO
        End If
        ws.Cells(r, 5) = n.ChildNodes(3).Text 'MENS RETORNO
        ws.Cells(r, 6) = n.ChildNodes(4).Text 'DT REGISTRO
        ws.Cells(r, 7) = n.ChildNodes(5).Text 'DT ULTIMA OCORRENCIA
        ws.Cells(r, 8) = n.ChildNodes(6).Text 'PRAZO RESP
        ws.Cells(r, 9) = n.ChildNodes(7).Text 'DATA RESP
        ws.Cells(r, 10) = n.ChildNodes(8).Text 'RESPOSTA
        r = r + 1
    Next n
    ws.Range("D:J").WrapText = True
    ws.Range("A:A").ColumnWidth = 16
End If
bordas ws, r
End Sub

Private Function getPath() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogOpen)

With fldr
    .Title = "Selecione o arquivo excel"
    .AllowMultiSelect = False
    .InitialFileName = "D:\CARTORIO\PROJETO\Nova pasta\"
    .FilterIndex = 5
    If .Show <> -1 Then GoTo nextcode
    sItem = .SelectedItems(1)
End With

nextcode:
getPath = sItem
Set fldr = Nothing

End Function

Private Function mensagens() As String
mensagens = "Caso já exista um PI cadastrado com o mesmo número de objeto, verifique o nome retornado pelo sistema, caso seja diferente do desejado, entre em contato com a GECAC através do e-mail SPMCOORDNAC@correios.com.br para as providências.."
End Function


Sub deleteSheetPis()
Dim ws As Worksheet
Dim r As Long
Set ws = ThisWorkbook.Sheets("RETORNO_PI")
r = ws.UsedRange.Rows.Count

If r > 5 Then ws.Rows("5:" & r).delete
ws.Range("a5").Select

End Sub

Private Sub bordas(ws As Worksheet, ByRef r As Long)

With ws.Range("A5:j" & r - 1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ws.Range("A5:j" & r - 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ws.Range("A5:j" & r - 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ws.Range("A5:j" & r - 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
     With ws.Range("A5:j" & r - 1).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ws.Range("A5:j" & r - 1).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub
