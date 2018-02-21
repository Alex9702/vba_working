Attribute VB_Name = "RetornaObjetos"
Option Explicit
Public Objetos As Variant

Public Sub GetObjetos(ByVal path As String)
Dim i As Integer: i = 0
Dim j As Integer: j = 1
Dim selItem As Variant

Dim fd As Office.FileDialog

Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False

Set fd = xlApp.Application.FileDialog(msoFileDialogFilePicker)

With fd
    .InitialFileName = path
    .Title = "Salvar como"
    If .Show <> 0 Then
        For Each selItem In fd.SelectedItems
           i = i + 1
        Next
        
        ReDim Objetos(1 To i)
        
        For Each selItem In fd.SelectedItems
           Objetos(j) = selItem
           j = j + 1
        Next
    End If
End With

Set fd = Nothing
xlApp.Quit
Set xlApp = Nothing

End Sub

