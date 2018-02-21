Attribute VB_Name = "sro"
Option Explicit

'PEGA OS NÚMEROS DOS OBJETOS PARA RASTREIO NO SRO
Private Function getObjetos(ByRef contador As Long, ws As Worksheet) As String

getObjetos = ""
Dim j As Long
Dim i As Integer
i = 0

For j = contador To ws.UsedRange.Rows.Count Step 1
    i = i + 1
    If getObjetos = "" Then
        getObjetos = Trim(ws.Range("A" & j))
    Else
        getObjetos = getObjetos & ";" & Trim(ws.Range("A" & j))
    End If
    contador = contador + 1
    If i = 50 Then Exit For
Next j

End Function

'BUSCA SRO
Sub buscasro(resultado As String)
Dim wb As Workbook
Set wb = ThisWorkbook
Dim i, x As Long
Dim j As Long
Dim flag As Boolean
j = 3
i = 3
flag = True
Dim ws As Worksheet
Set ws = wb.Sheets(1)
Dim wsSro As Worksheet

insereTab ("Rastreamento")

Set wsSro = wb.Sheets("Rastreamento")

Dim xDoc As MSXML2.DOMDocument
Set xDoc = New MSXML2.DOMDocument
xDoc.async = False
xDoc.validateOnParse = False
Dim url As String

Dim total As Long
total = ws.UsedRange.Rows.Count

Dim nodelist As MSXML2.IXMLDOMNodeList

For x = 2 To total
    statusBar total, x
    url = "http://webservicesro/sro_xml/xml?Tipo=L&Resultado=" & resultado & "&Evento=&Objetos=" & getObjetos(j, ws)
    If xDoc.Load(url) Then
        Set nodelist = xDoc.getElementsByTagName("sroxml/objeto")
        Dim n, e As IXMLDOMNode
        For Each n In nodelist
        
            For Each e In n.SelectNodes("evento")
            
                wsSro.Cells(i, 1) = n.ChildNodes(0).Text 'Nº DO OBJETO
                
                If flag = True Then
                     ws.Hyperlinks.Add Anchor:=wsSro.Cells(i, 1), _
                    Address:="http://websro2.correiosnet.int/rastreamento/" & _
                    "sro?opcao=PESQUISA&objetos=" & n.ChildNodes(0).Text
                    flag = False
                End If
                wsSro.Cells(i, 2) = e.ChildNodes(4).Text 'DESCRIÇÃO
                wsSro.Cells(i, 3) = e.ChildNodes(2).Text 'DATA
                wsSro.Cells(i, 4) = e.ChildNodes(5).ChildNodes(2).Text 'CIDADE
                wsSro.Cells(i, 5) = e.ChildNodes(5).ChildNodes(3).Text 'UF
                wsSro.Cells(i, 6) = e.ChildNodes(5).ChildNodes(0).Text 'LOCAL / UNIDADE
                wsSro.Cells(i, 7) = e.ChildNodes(5).ChildNodes(6).Text 'DR
                wsSro.Cells(i, 8) = e.ChildNodes(1).Text 'STATUS
                wsSro.Cells(i, 9) = e.ChildNodes(0).Text 'TIPO
                i = i + 1
            Next e
            flag = True
        Next n
    End If
    x = j
Next
Cells.EntireColumn.AutoFit
Application.statusBar = False
End Sub

Private Function statusBar(total As Long, ByVal contado As Long)
    Dim conta As Long
    conta = 100 * contado / total
    Application.statusBar = "Total de busca: " & conta & "%"
    
End Function

Private Sub Teste()
statusBar 1520, 352
End Sub
