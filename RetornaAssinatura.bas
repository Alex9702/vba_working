Attribute VB_Name = "RetornaAssinatura"
Public Function Assinatura() As String
    Dim pathHtml As String: pathHtml = "....\AppData\Roaming\Microsoft\" & _
    "Assinaturas\Assinatura.htm"
    
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(pathHtml).OpenAsTextStream(1, -2)
    Assinatura = ts.ReadAll
    ts.Close
End Function

