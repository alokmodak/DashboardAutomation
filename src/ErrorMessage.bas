Attribute VB_Name = "ErrorMessage"
Option Explicit

'Date           Who     What
'10-June-2015   Name    First file - Purpose is to to perform error handling on UI
'16-June-2015   Name    Added validation for checkbox


Public Sub Error_Messagebox()

'Messagebox if no drive is selected
If Sheet1.rdbLocalDrive.value = False Then
    If Sheet1.rdbSharedDrive.value = False Then
            MsgBox "Please select a Data Source!" & vbCrLf & "Local or Shared or Sharepoint"
            End
    End If
End If

'Message box to select year
If Sheet1.combYear.value = "Select" Then
MsgBox "Please select a Year/Month Value!"
End
End If


'message box for option select CTS, Revenue or Dashboard
If Sheet1.chkDashboard.value = False Then
    If Sheet1.chkCTS.value = False Then
        If Sheet1.chkRevenue.value = False Then
            MsgBox "Please select Output option"
            End
        End If
    End If
End If

'Messagebox if no values selected in product group or CTS or revenue
If Sheet1.combProductGroup.value = "Select Product Group" Then
 MsgBox "Please select a value in Dashboard Product group"
        End
    ElseIf Sheet1.comb6NC2.value = "Select Product Group" Then
     MsgBox "Please select a value in Product group"
        End
        ElseIf Sheet1.comb6NC1.value = "Select Product Group" Then
         MsgBox "Please select a value in Product group"
        End
End If


'message box for file name and Selected date correction
    Dim monthChk As String
    Dim yearChk As String
    Dim inputItem As String
    Dim flNameCheckDate As String
    Dim selectedDate As String
    Dim fstMonthChk As String
    Dim servicefileName As String
    Dim innovationFileName As String
    Dim outputPath As String
    Dim installFileOpen As String
    Dim TestStr As String
    
    outputPath = ThisWorkbook.Path & "\" & "Mos KPI Summary.xlsx" 'output file path
    TestStr = ""
    TestStr = Dir(outputPath)
    On Error GoTo 0
    If TestStr = "" Then
        MsgBox "Output File with name " & Chr(34) & "Mos KPI Summary.xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'for service scorecard
    fstMonthChk = Format(Sheet1.combYear.value, "mmmyy")
    servicefileName = ""
    servicefileName = Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*")
    
    If servicefileName = "" Then
    MsgBox "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End
    End If
    
    'for innovation file
    innovationFileName = ""
    innovationFileName = Dir(ThisWorkbook.Path & "\" & "KPI dashboard_Innovation_" & fstMonthChk & "*.xls*")
    
    If innovationFileName = "" Then
    MsgBox "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End
    End If

    'checking Install Hrs file exists
    installFileOpen = ""
    installFileOpen = Dir(ThisWorkbook.Path & "\" & "Install SPAN P95_" & fstMonthChk & "*.xls*")
    If installFileOpen = "" Then
        MsgBox "Input File with name " & Chr(34) & "Install SPAN P95_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
        End
    End If

    'Checking for FCO OP review file.xlsx

    Dim fcoFileOpen As String
    fcoFileOpen = ""
    fcoFileOpen = Dir(ThisWorkbook.Path & "\" & "FCO OP review file_" & fstMonthChk & "*.xls*")
    If fcoFileOpen = "" Then
        MsgBox "Input File with name " & Chr(34) & "FCO OP review file_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'Escalations_Overview_ALL BIUs.xlsx
    Dim escOFileOpen As String
    escOFileOpen = ""
    escOFileOpen = Dir(ThisWorkbook.Path & "\" & "Escalations_Overview_ALL BIUs_" & fstMonthChk & "*.xls*")
    If escOFileOpen = "" Then
        MsgBox "Input File with name " & Chr(34) & "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'Customer escalations (Weekly Review) Complaints.xlsx

    Dim compOFileOpen As String
    compOFileOpen = ""
    compOFileOpen = Dir(ThisWorkbook.Path & "\" & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & "*.xls*")
    If compOFileOpen = "" Then
        MsgBox "Input File with name " & Chr(34) & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    '2015-05 Installation spend L2-report.xlsb
    Dim inscostFileOpen As String
    inscostFileOpen = ""
    inscostFileOpen = Dir(ThisWorkbook.Path & "\" & Sheet1.combYear.value & " " & "Installation spend L2-report" & "*.xls*")
        If inscostFileOpen = "" Then
        MsgBox "Input File with name " & Chr(34) & Sheet1.combYear.value & " " & "Installation spend L2-report" & ".xlsb" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'warranty cost file
    Dim warrantyCostFile1 As String
    Dim warrantyCostFile2 As String
    Dim valFind As String, found As String
    Dim found2 As String
    
    valFind = Replace(Sheet1.combYear.value, "-", "")
    warrantyCostFile1 = Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*IGT.xls*")
    warrantyCostFile2 = Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*DI.xls*")
    found = InStr(1, warrantyCostFile1, valFind, vbTextCompare)
    found = InStr(1, warrantyCostFile2, valFind, vbTextCompare)
    
    If found = "0" Then
    MsgBox "Input File with name " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb" & Chr(34) & " or " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb" & Chr(34) & vbCrLf & " doesn't exist!"
    End
    End If
End Sub
