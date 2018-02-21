Attribute VB_Name = "Teste"
Private Sub Teste()
Dim xDoc As MSXML2.DOMDocument60
Set xDoc = New MSXML2.DOMDocument60
xDoc.async = False
xDoc.validateOnParse = False

If xDoc.Load("D:\CARTORIO\Arquivos Temporários\download.xml") Then
    Set listas = xDoc.DocumentElement
    
    For Each listaNode In listas.ChildNodes
        For Each fieldNode In listaNode.ChildNodes
            Debug.Print "[" & fieldNode.BaseName & "] = [" & fieldNode.Text & "]"
        Next fieldNode
    Next listaNode
End If

Set xDoc = Nothing
End Sub
