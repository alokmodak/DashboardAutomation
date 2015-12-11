Attribute VB_Name = "DataDownload"
Public Sub Data_Download()
On Error Resume Next
Const READYSTATE_COMPLETE As Integer = 4

  Dim objIE As Object
  
  Set objIE = CreateObject("InternetExplorer.Application")

  With objIE
    .Visible = True
    .silent = True
    .Navigate "https://apmpw63.hsec.emeadc001.philips.com:8463/sap/bw/BEx?SAP-LANGUAGE=EN&BOOKMARK_ID=CU4A96HFEQL9ONOMSTDAUOMPU&VARIABLE_SCREEN=X"
    Do Until .ReadyState = READYSTATE_COMPLETE
      DoEvents
    Loop
  End With
  Set objIE = Nothing

End Sub


Sub Test()
    Set IE = CreateObject("InternetExplorer.application")
    IE.Visible = True
    IE.Navigate ("https://webmail.igate.com/owa/auth/logon.aspx?replaceCurrent=1&url=https%3a%2f%2fwebmail.igate.com%2fowa%2f")
    Do
        If IE.ReadyState = 4 Then
            IE.Visible = True
            Exit Do
        Else
            DoEvents
        End If
    Loop
    IE.document.Forms(0).all("Domain\user name").value = "me"
    IE.document.Forms(0).all("Password").value = "mypasssword"
    IE.document.Forms(0).submit
    
End Sub
