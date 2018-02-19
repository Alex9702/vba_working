Attribute VB_Name = "Negativas"
Private Sub NegativasMensagem()
Dim objMsg As MailItem
Dim strHtml As String
Dim cumprimento As String
Dim data As Date, hora As Integer

data = DateValue(Now)
hora = Int(Hour(Now))

If hora > 12 Then
    cumprimento = "<p class=MsoNormal><span style='font-size:10.0pt;" & _
    "font-family:""Century Gothic"",""sans-serif""'>Boa Tarde.</span></p>"
Else
    cumprimento = "<p class=MsoNormal><span style='font-size:10.0pt;" & _
    "font-family:""Century Gothic"",""sans-serif""'>Bom Dia.</span></p>"
End If

strHtml = "<p class=MsoNormal><span style='font-size:10.0pt;" & _
    "font-family:""Century Gothic"",""sans-serif\'""'>" & _
    "Seguem Negativas realizadas dia " & data & ".</span></p>" & _
    "<p class=MsoNormal><span style='font-size:10.0pt;" & _
    "font-family:""Century Gothic"",""sans-serif""'>Atenciosamente,</span></p>"

Set objMsg = Application.CreateItem(olMailItem)

With objMsg
    .To = "email@email.com"
    .CC = "email@email.com"
    .Subject = "PI"
    .Categories = "test"
    '.VotingOptions = "Yes;No;Maybe;"
    .BodyFormat = olFormatHTML
    .HTMLBody = cumprimento & strHtml & Assinatura
    '.Importance = olImportanceHigh
    '.Sensitivity = olConfidential
    GetObjetos "D:\"
    On Error Resume Next
    For Each o In Objetos
        .Attachments.Add o
    Next

    
    '.ExpiryTime = DateAdd("m", 6, Now)
    '.DeferredDeliveryTime = #8/2/2018 6:00:00 PM#
    .Display
    '.Send 'use to send it automatically
End With

Set objMsg = Nothing

End Sub

