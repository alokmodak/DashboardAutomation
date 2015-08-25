Attribute VB_Name = "ErrorMessage"
Option Explicit

'Date           Who     What
'10-June-2015   Name    First file - Purpose is to to perform error handling on UI
'16-June-2015   Name    Added validation for checkbox

Public Sub Error_Messagebox()

Dim fileNotPresent As String

'Messagebox if no drive is selected
If Sheet1.rdbLocalDrive.Value = False Then
    If Sheet1.rdbSharedDrive.Value = False Then
            MsgBox "Please select a Data Source!" & vbCrLf & "Local or Shared or Sharepoint"
            End
    End If
End If

'Message box to select year
If Sheet1.combYear.Value = "Select" Then
MsgBox "Please select a Year/Month Value!"
End
End If


'message box for option select CTS, Revenue or Dashboard
If Sheet1.chkDashboard.Value = False Then
    If Sheet1.chkCTS.Value = False Then
        If Sheet1.chkRevenue.Value = False Then
            MsgBox "Please select Output option"
            End
        End If
    End If
End If

'Messagebox if no values selected in product group or CTS or revenue
If Sheet1.combProductGroup.Value = "Select Product Group" Then
 MsgBox "Please select a value in Dashboard Product group"
        End
    ElseIf Sheet1.comb6NC2.Value = "Select Product Group" Then
     MsgBox "Please select a value in Product group"
        End
        ElseIf Sheet1.comb6NC1.Value = "Select Product Group" Then
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
    Dim CQDataFile As String
    Dim serviceInfoDataFile As String
    
If Sheet1.rdbLocalDrive.Value = True Then 'Input files check for Local drive

    outputPath = ThisWorkbook.Path & "\" & "KPI Summary.xlsx" 'output file path
    TestStr = ""
    TestStr = Dir(outputPath)
    On Error GoTo 0
    If TestStr = "" Then
        MsgBox "Output File with name " & Chr(34) & "KPI Summary.xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'for service scorecard
    fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")
    servicefileName = ""
    servicefileName = Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*")
    
    If servicefileName = "" Then
    fileNotPresent = Chr(34) & "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm" & Chr(34) & vbCrLf
    End If
    
    'for innovation file
    innovationFileName = ""
    innovationFileName = Dir(ThisWorkbook.Path & "\" & "KPI dashboard_Innovation_" & fstMonthChk & "*.xls*")
    
    If innovationFileName = "" Then
    fileNotPresent = fileNotPresent & Chr(34) & "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
    End If

    'checking Install Hrs file exists
    installFileOpen = ""
    installFileOpen = Dir(ThisWorkbook.Path & "\" & "Install SPAN P95.xlsx")
    If installFileOpen = "" Then
        fileNotPresent = fileNotPresent & Chr(34) & "Install SPAN P95.xlsx" & Chr(34) & vbCrLf
    End If

    'Checking for FCO OP review file.xlsx

    Dim fcoFileOpen As String
    fcoFileOpen = ""
    fcoFileOpen = Dir(ThisWorkbook.Path & "\" & "FCO OP review file_" & fstMonthChk & "*.xls*")
    If fcoFileOpen = "" Then
        fileNotPresent = fileNotPresent & Chr(34) & "FCO OP review file.xlsx" & Chr(34) & vbCrLf
    End If
    
    'Escalations_Overview_ALL BIUs.xlsx
    Dim escOFileOpen As String
    escOFileOpen = ""
    escOFileOpen = Dir(ThisWorkbook.Path & "\" & "Escalations_Overview_ALL BIUs_" & fstMonthChk & "*.xls*")
    If escOFileOpen = "" Then
        fileNotPresent = fileNotPresent & Chr(34) & "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
    End If
    
    'Customer escalations (Weekly Review) Complaints.xlsx

    Dim compOFileOpen As String
    compOFileOpen = ""
    compOFileOpen = Dir(ThisWorkbook.Path & "\" & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & "*.xls*")
    If compOFileOpen = "" Then
        fileNotPresent = fileNotPresent & Chr(34) & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
    End If
    
    '2015-05 Installation spend L2-report.xlsb
    Dim inscostFileOpen As String
    inscostFileOpen = ""
    inscostFileOpen = Dir(ThisWorkbook.Path & "\" & Sheet1.combYear.Value & " " & "Installation spend L2-report" & "*.xls*")
        If inscostFileOpen = "" Then
        fileNotPresent = fileNotPresent & Chr(34) & Sheet1.combYear.Value & " " & "Installation spend L2-report" & ".xlsb" & Chr(34) & vbCrLf
    End If
    
    'warranty cost file
    Dim warrantyCostFile1 As String
    Dim warrantyCostFile2 As String
    Dim valFind As String, found As String
    Dim found2 As String
    
    valFind = Replace(Sheet1.combYear.Value, "-", "")
    warrantyCostFile1 = Dir(ThisWorkbook.Path & "\" & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb")
    warrantyCostFile2 = Dir(ThisWorkbook.Path & "\" & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb")
    found = InStr(1, warrantyCostFile1, valFind, vbTextCompare)
    found = InStr(1, warrantyCostFile2, valFind, vbTextCompare)
    
    If found = "0" Then
    fileNotPresent = fileNotPresent & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb" & Chr(34) & " or " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb" & Chr(34) & vbCrLf
    End If
    
    'For CQ Data File Validation
    CQDataFile = ""
    CQDataFile = Dir(ThisWorkbook.Path & "\" & "CQ_Data_SPM.xlsx")
    
    If CQDataFile = "" Then
    fileNotPresent = fileNotPresent & Chr(34) & "CQ_Data_SPM.xlsx" & Chr(34) & vbCrLf
    End If
    
    'For service information Data File Validation
    serviceInfoDataFile = ""
    serviceInfoDataFile = Dir(ThisWorkbook.Path & "\" & "Service_Information_Quality_Completion.xlsx")
    
    If CQDataFile = "" Then
    fileNotPresent = fileNotPresent & Chr(34) & "Service_Information_Quality_Completion.xlsx" & Chr(34) & vbCrLf
    End If

Dim flNt As Integer
Dim msg1
        msg1 = MsgBox("Following Files are not present at" & " " & Chr(34) & ThisWorkbook.Path & Chr(34) & vbCrLf & fileNotPresent & vbCrLf & " Do you want to Continue?", vbYesNo)
        If msg1 = vbNo Then
        End
        End If
End If

    
'validation for files present over shared drive
If Sheet1.rdbSharedDrive.Value = True Then
    
    Dim fileNotFoundShared As String
    
    SharedDrive_Path "KPI Summary.xlsx"
    
    If fileExists = False Then
        MsgBox "Output File with name " & Chr(34) & "KPI Summary.xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'for service scorecard
    fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")
    SharedDrive_Path ("Service Scorecard F 6.1_" & fstMonthChk & ".xlsm")
    If fileExists = False Then
    fileNotFoundShared = Chr(34) & "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm" & Chr(34) & vbCrLf
    End If
    
    'for innovation file
    innovationFileName = "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx"
    SharedDrive_Path innovationFileName
    
    If fileExists = False Then
    fileNotFoundShared = fileNotFoundShared & Chr(34) & "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
    End If

    'checking Install Hrs file exists
    installFileOpen = "Install SPAN P95.xlsx"
    SharedDrive_Path installFileOpen
    
    If fileExists = False Then
        fileNotFoundShared = fileNotFoundShared & Chr(34) & "Install SPAN P95.xlsx" & Chr(34) & vbCrLf
        
    End If

    'Checking for FCO OP review file.xlsx
    fcoFileOpen = "FCO OP review file.xlsx"
    SharedDrive_Path fcoFileOpen
    
    If fileExists = False Then
        fileNotFoundShared = fileNotFoundShared & Chr(34) & "FCO OP review file.xlsx" & Chr(34) & vbCrLf
        
    End If
    
    'Escalations_Overview_ALL BIUs.xlsx
    escOFileOpen = "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx"
    SharedDrive_Path escOFileOpen
    
    If fileExists = False Then
        fileNotFoundShared = fileNotFoundShared & Chr(34) & "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
        
    End If
    
    'Customer escalations (Weekly Review) Complaints.xlsx
    compOFileOpen = "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx"
    SharedDrive_Path compOFileOpen
    
    If fileExists = False Then
        fileNotFoundShared = fileNotFoundShared & Chr(34) & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" & Chr(34) & vbCrLf
        
    End If
    
    '2015-05 Installation spend L2-report.xlsb
    inscostFileOpen = Sheet1.combYear.Value & " " & "Installation spend L2-report" & ".xlsb"
    SharedDrive_Path inscostFileOpen
    
    If fileExists = False Then
        fileNotFoundShared = fileNotFoundShared & Chr(34) & Sheet1.combYear.Value & " " & "Installation spend L2-report" & ".xlsb" & Chr(34) & vbCrLf
        
    End If
    
    'warranty cost file
    
    valFind = Replace(Sheet1.combYear.Value, "-", "")
    warrantyCostFile1 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb"
    warrantyCostFile2 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb"
    SharedDrive_Path warrantyCostFile1
    SharedDrive_Path warrantyCostFile2
    
    If fileExists = False Then
    fileNotFoundShared = fileNotFoundShared & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb" & Chr(34) & " or " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb" & Chr(34) & vbCrLf
    
    End If
    
    
    'For PLCM Data File Validation
    Dim PLCMDataFile As String
    PLCMDataFile = "CS_Dashboard.xlsx"
    SharedDrive_Path PLCMDataFile
    
    If fileExists = False Then
    fileNotFoundShared = fileNotFoundShared & Chr(34) & "CS_Dashboard.xlsx" & Chr(34) & vbCrLf
    End If
    
    
    'For CQ Data File Validation
    CQDataFile = "CQ_Data_SPM.xlsx"
    SharedDrive_Path CQDataFile
    
    If fileExists = False Then
    fileNotFoundShared = fileNotFoundShared & Chr(34) & "CQ_Data_SPM.xlsx" & Chr(34) & vbCrLf
    
    End If
    
    'For service information Data File Validation
    serviceInfoDataFile = "Service_Information_Quality_Completion.xlsx"
    SharedDrive_Path "Service_Information_Quality_Completion.xlsx"
    
    If fileExists = False Then
    fileNotFoundShared = fileNotFoundShared & Chr(34) & "Service_Information_Quality_Completion.xlsx" & Chr(34) & vbCrLf
    
    End If
        
        msg1 = MsgBox("Following files are not present at the selected location " & Chr(34) & Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & Chr(34) & " , Do you want to Continue?" & vbCrLf & fileNotFoundShared, vbYesNo)
        If msg1 = vbNo Then
        End
        End If
End If

End Sub

'Error message for Revenue file not present
Public Sub ErrorMessage_RevenueFiles(inputFile As String)
On Error Resume Next
Dim inputRevenue As String
inputRevenue = inputFile

If Sheet1.rdbSharedDrive.Value = True Then
    SharedDrive_Path inputRevenue
Else
    sharedDrivePath = ThisWorkbook.Path & "\" & inputRevenue
End If

Dim flPresent As String
flPresent = ""
flPresent = Dir(sharedDrivePath)
If flPresent = "" Then
MsgBox Chr(34) & inputRevenue & Chr(34) & " File not Found! Please select Appropriate Path"
End
End If
marketInputFile = "Market_Groups_Markets_Country.xlsx"

If Sheet1.rdbSharedDrive.Value = True Then
    SharedDrive_Path marketInputFile
Else
    sharedDrivePath = ThisWorkbook.Path & "\" & marketInputFile
End If

flPresent = ""
flPresent = Dir(sharedDrivePath)
If flPresent = "" Then
MsgBox Chr(34) & marketInputFile & Chr(34) & " File not Found! Please select Appropriate Path"
End
End If

End Sub
