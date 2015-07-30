
Private Sub Workbook_Open()

Call ImportCode

'putting values in UI boxes
Call Automation_Tool_File_Open
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

   Call hideTipsForWorkbookEvent
   
   Call exportVbaCode
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Call hideTipsForWorkbookEvent
    
    Call exportVbaCode
    
End Sub

Public Sub SetEnvironmentVariable()

Dim objUserEnvVars As Object
Dim strVar As String
Set objUserEnvVars = CreateObject("WScript.Shell").Environment("User")
strVar = objUserEnvVars.Item("Dashboard_Automation")
MsgBox strVar
End Sub
