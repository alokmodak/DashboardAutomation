Attribute VB_Name = "Send_Mail"
Sub Send_Email_Via_OutlookInbox(flName As String, toEmailAdd As String, subject As String, txtBody As String)

    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = toEmailAdd
        '.CC = Cells(i, 2).Value
        .subject = subject
        .Body = txtBody
        .Attachments.Add flName
        .Display  'or use .Send
    End With
    
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
