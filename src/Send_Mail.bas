Attribute VB_Name = "Send_Mail"
'Date           Who     What

'**********************************
'Set Microsoft CDO for windows 2000 library as reference if the below code doesnt work


Sub SendEmailUsingYahoo()

Set myMail = CreateObject("CDO.Message")

Dim NewMail As CDO.Message
   
Set NewMail = New CDO.Message
  
'Enable SSL Authentication
NewMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
  
'Make SMTP authentication Enabled=true (1)
  
NewMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
  
'Set the SMTP server and port Details
'To get these details you can get on Settings Page of your Yahoo Account
  
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.mail.yahoo.com"
  
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
  
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  
'Set your credentials of your Gmail Account
  
NewMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "vishwamitra01@yahoo.com"
  
NewMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "**********"
  
'Update the configuration fields
NewMail.Configuration.Fields.Update
   
'Set All Email Properties
  
With NewMail
  .Subject = ""
  .From = ""
  .To = ""
  .CC = ""
  .BCC = ""
  .textbody = ""
  '.AddAttachment "C:\ABC.xls"
End With
  
NewMail.Send
MsgBox ("Mail has been Sent")
  
'Set the NewMail Variable to Nothing
Set NewMail = Nothing
   

End Sub
