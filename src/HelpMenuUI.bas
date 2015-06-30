Attribute VB_Name = "HelpMenuUI"
'Date           Who     What


Option Private Module

Public clearInfoTime As Date            'the time hideInfoTips is scheduled for
                                        '(needed to cancel scheduled clear in event of workbook close)
                                        
Public infoClearScheduled As Boolean    'flag to tell Workbook_BeforeClose event that a clear is scheduled
                                        '(testing clearInfoTime is not reliable)

Sub hideInfoTips(Optional Dummy As Boolean)
'Dummy argument stops this macro from appearing in user macros list

    With Sheet1
        .lblDashboard.Visible = False
    End With
    
    Application.Cursor = xlDefault
    
    infoClearScheduled = False

End Sub

Public Sub showInfoTip(currentTip As Object)
    
    Static previousTip As Object
    
    On Error Resume Next
    
    'for labels Input Files
Dim lblMonthCheck As String
Dim lblText As String
Dim warrantyVal As String

warrantyVal = Replace(Sheet1.combYear.value, "-", "")
 lblMonthCheck = Format(Sheet1.combYear.value, "mmmyy")
 
 lblText = "Input files Required Are -" & vbCrLf & _
"1) Service Scorecard F 6.1_" & lblMonthCheck & ".xlsm" & vbCrLf & _
"2) Install SPAN P95_" & lblMonthCheck & ".xlsx" & vbCrLf & _
"3) FCO OP review file_" & lblMonthCheck & ".xlsx" & vbCrLf & _
"4) Escalations_Overview_ALL BIUs_" & lblMonthCheck & ".xlsx" & vbCrLf & _
"5) Customer escalations (Weekly Review) Complaints_" & lblMonthCheck & ".xlsx" & vbCrLf & _
"6) " & Sheet1.combYear.value & " Installation spend L2-report" & lblMonthCheck & ".xlsb" & vbCrLf & _
"7) Level 4 Warranty Spend Analysis - " & warrantyVal & " @ " & warrantyVal - 1 & " BS Rate_IGT" & vbCrLf & _
"8) Level 4 Warranty Spend Analysis - " & warrantyVal & " @ " & warrantyVal - 1 & " BS Rate_DI" & vbCrLf & _
"9) KPI dashboard_Innovation_" & lblMonthCheck & ".xlsx" & vbCrLf & _
"10) CQ_Data_SPM.xlsx"

Sheet1.lblDashboard.Caption = lblText


    
    Application.Cursor = xlNorthwestArrow 'stops mouse pointer from flickering between hourglass and arrow
    
    If Not currentTip Is previousTip Or Not currentTip.Visible Then
        
        previousTip.Visible = False
        currentTip.Visible = True
        Set previousTip = currentTip
    
    End If
    
    Application.OnTime clearInfoTime, "hideInfoTips", , False 'ditch any previously scheduled clear
    
    clearInfoTime = Now + TimeSerial(0, 0, 10) 'approx time to show infotip before clearing (h, m, s)
    
    infoClearScheduled = True
    
    Application.OnTime clearInfoTime, "hideInfoTips"

End Sub

Public Sub hideTipsForWorkbookEvent()
    
    If infoClearScheduled Then
    
        On Error Resume Next
    
        Application.OnTime clearInfoTime, "hideInfoTips", , False

        With Sheet1
            .lblDashboard.Visible = False
        End With
    
    End If

    Application.Cursor = xlDefault

End Sub
