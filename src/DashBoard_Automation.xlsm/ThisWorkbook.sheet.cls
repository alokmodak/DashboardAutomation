
Private Sub Workbook_Open()

'putting values in UI boxes
Call Automation_Tool_File_Open
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

   Call hideTipsForWorkbookEvent
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Call hideTipsForWorkbookEvent

End Sub
