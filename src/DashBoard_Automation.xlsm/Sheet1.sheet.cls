Option Explicit

Private Sub chkDashboard_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

    If Sheet1.chkDashboard.Value = True Then
    
            Call Error_Messagebox
            
            Call Generate_Dashboard_Output 'Generating output file
            
            Call Calculate_Innovation 'Fleching values for # of Patterns
                 
            Call Complaints_Escalations_Calculation
            
            Call FCO_Calculations
            
            Call Installation_Cost_Calculation
            
            Call Warranty_Cost_Calculations
            
            Call CQ_Calculations
            
    ElseIf Sheet1.chkRevenue.Value = True Then
    'Call Revenue_Product_Group_Select
    
    ElseIf Sheet1.chkCTS.Value = True Then
    
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
Sheet1.processTime.Value = processTime
Sheet1.minProcessTime.Value = minProcessTime
End Sub

'UI Date Control
Private Sub cmbIncreaseMonth_Click()

Call Increase_Year_And_Month_Value

End Sub

Private Sub cmbDecreaseMonth_Click()

Call Decrease_Year_And_Month_Value
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
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.comb6NC1.Enabled = False
        Sheet1.comb6NC2.Enabled = False
        Sheet1.combProductGroup.Value = ""
        Sheet1.combProductGroup.Enabled = False
    Else
        Sheet1.comb6NC1.Enabled = True
        Sheet1.comb6NC2.Enabled = True
        Sheet1.combProductGroup.Enabled = True
        Sheet1.combProductGroup.Value = "Select Product Group"
    End If
End Sub