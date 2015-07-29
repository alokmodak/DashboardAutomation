Option Explicit
Public exitLoop As Integer


Private Sub chkDashboard_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call showInfoTip(lblDashboard)
End Sub

'output file generation
Private Sub cmbOutput_Click()

Dim startTime As Integer
Dim endTime As Integer
Dim processTime As Integer
Dim minstartTime As Integer
Dim minEndTime As Integer
Dim minProcessTime As Integer

minstartTime = Minute(Now()) 'start time for process time
startTime = Second(Now())

    If Sheet1.chkDashboard.value = True Then
    
            Call Error_Messagebox
            
            Call Generate_Dashboard_Output 'Generating output file
            
            Call Calculate_Innovation 'Fleching values for # of Patterns
                 
            Call Complaints_Escalations_Calculation
            
            Call FCO_Calculations
            
            Call Installation_Cost_Calculation
            
            Call Warranty_Cost_Calculations
            
            Call Service_Information
            
            Call CQ_Calculations
            
    ElseIf Sheet1.chkRevenue.value = True Then
            
            Call Revenue_Graph_Creation
    
    ElseIf Sheet1.chkCTS.value = True Then
    
    Else
    MsgBox "Please select an output option"
    
    End If
    

minEndTime = Minute(Now())
endTime = Second(Now())
minProcessTime = minEndTime - minstartTime
processTime = endTime - startTime
If endTime < startTime Then
startTime = 60 - startTime
processTime = endTime + startTime
End If
Sheet1.processTime.value = processTime
Sheet1.minProcessTime.value = minProcessTime
End Sub

'UI Date Control
Private Sub cmbIncreaseMonth_Click()

Call Increase_Year_And_Month_Value

End Sub

Private Sub cmbDecreaseMonth_Click()

Call Decrease_Year_And_Month_Value
End Sub


Private Sub cmb6NC2OK_Click()
On Error Resume Next
    Sheet1.lstBx6NC.Visible = False
    Sheet1.cmb6NC2OK.Visible = False
    Sheet1.comb6NC2.value = ""
    Dim i As Integer
    For i = 0 To 20
        If Sheet1.lstBx6NC.Selected(i) = True Then
            Sheet1.comb6NC2.value = Sheet1.comb6NC2.value & " " & Sheet1.lstBx6NC.List(i) & " "
        End If
    Next
End Sub


Private Sub comb6NC2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Sheet1.lstBx6NC.Visible = False Then
        Sheet1.lstBx6NC.Visible = True
        Sheet1.cmb6NC2OK.Visible = True
    End If
End Sub

Private Sub helpMenu_Click()
On Error Resume Next
Dim wordFl As Object
Dim helpPath As String
helpPath = ThisWorkbook.Path & "\" & "DashboardAutomation_Help.doc"
Set wordFl = CreateObject("word.application")
wordFl.documents.Open helpPath
wordFl.Visible = True
wordFl.Activate
End Sub


Private Sub chkAllGroups_Click()
    If Sheet1.chkAllGroups.value = True Then
        Sheet1.comb6NC1.Enabled = False
        Sheet1.comb6NC2.Enabled = False
        Sheet1.combProductGroup.value = ""
        Sheet1.combProductGroup.Enabled = False
        Sheet1.lstBx6NC.MultiSelect = fmMultiSelectSingle
        Sheet1.lstBx6NC.value = ""
        Sheet1.lstBx6NC.MultiSelect = fmMultiSelectMulti
        Sheet1.comb6NC2.value = ""
    Else
        Sheet1.comb6NC1.Enabled = True
        Sheet1.comb6NC2.Enabled = True
        Sheet1.combProductGroup.Enabled = True
        Sheet1.combProductGroup.value = "Select Product Group"
    End If
End Sub
