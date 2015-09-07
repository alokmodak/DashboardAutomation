Attribute VB_Name = "Send_Mail"
Sub CDO_Mail_Small_Text()
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim Flds As Variant

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

        iConf.Load -1    ' CDO Source Defaults
        Set Flds = iConf.Fields
        With Flds
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
            '.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mumcas.igate.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .Update
        End With

    strbody = "Hi there"

    With iMsg
        Set .Configuration = iConf
        .To = "Jitendra.deshmukh@igate.com"
        .CC = ""
        .BCC = ""
        .From = "Jitendra.deshmukh@igate.com"
        .Subject = "New figures"
        .TextBody = strbody
        .Send
    End With

End Sub
