Attribute VB_Name = "ErrorMessage"
Option Explicit

'Date           Who     What
'10-June-2015   Name    First file - Purpose is to to perform error handling on UI
'16-June-2015   Name    Added validation for checkbox


Public Sub Error_Messagebox()

Dim fileNotPresent(10) As String

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
    Dim CQDataFile As String
    Dim serviceInfoDataFile As String
    
If Sheet1.rdbLocalDrive.value = True Then 'Input files check for Local drive

    outputPath = ThisWorkbook.Path & "\" & "KPI Summary.xlsx" 'output file path
    TestStr = ""
    TestStr = Dir(outputPath)
    On Error GoTo 0
    If TestStr = "" Then
        MsgBox "Output File with name " & Chr(34) & "KPI Summary.xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'for service scorecard
    fstMonthChk = Format(Sheet1.combYear.value, "mmmyy")
    servicefileName = ""
    servicefileName = Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*")
    
    If servicefileName = "" Then
    fileNotPresent(1) = "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End If
    
    'for innovation file
    innovationFileName = ""
    innovationFileName = Dir(ThisWorkbook.Path & "\" & "KPI dashboard_Innovation_" & fstMonthChk & "*.xls*")
    
    If innovationFileName = "" Then
    fileNotPresent(2) = "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End If

    'checking Install Hrs file exists
    installFileOpen = ""
    installFileOpen = Dir(ThisWorkbook.Path & "\" & "Install SPAN P95_" & fstMonthChk & "*.xls*")
    If installFileOpen = "" Then
        fileNotPresent(3) = "Input File with name " & Chr(34) & "Install SPAN P95_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
    End If

    'Checking for FCO OP review file.xlsx

    Dim fcoFileOpen As String
    fcoFileOpen = ""
    fcoFileOpen = Dir(ThisWorkbook.Path & "\" & "FCO OP review file_" & fstMonthChk & "*.xls*")
    If fcoFileOpen = "" Then
        fileNotPresent(4) = "Input File with name " & Chr(34) & "FCO OP review file_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
    End If
    
    'Escalations_Overview_ALL BIUs.xlsx
    Dim escOFileOpen As String
    escOFileOpen = ""
    escOFileOpen = Dir(ThisWorkbook.Path & "\" & "Escalations_Overview_ALL BIUs_" & fstMonthChk & "*.xls*")
    If escOFileOpen = "" Then
        fileNotPresent(5) = "Input File with name " & Chr(34) & "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"
        
    End If
    
    'Customer escalations (Weekly Review) Complaints.xlsx

    Dim compOFileOpen As String
    compOFileOpen = ""
    compOFileOpen = Dir(ThisWorkbook.Path & "\" & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & "*.xls*")
    If compOFileOpen = "" Then
        fileNotPresent(6) = "Input File with name " & Chr(34) & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" & Chr(34) & " doesn't exist!"

    End If
    
    '2015-05 Installation spend L2-report.xlsb
    Dim inscostFileOpen As String
    inscostFileOpen = ""
    inscostFileOpen = Dir(ThisWorkbook.Path & "\" & Sheet1.combYear.value & " " & "Installation spend L2-report" & "*.xls*")
        If inscostFileOpen = "" Then
        fileNotPresent(7) = "Input File with name " & Chr(34) & Sheet1.combYear.value & " " & "Installation spend L2-report" & ".xlsb" & Chr(34) & " doesn't exist!"
    
    End If
    
    'warranty cost file
    Dim warrantyCostFile1 As String
    Dim warrantyCostFile2 As String
    Dim valFind As String, found As String
    Dim found2 As String
    
    valFind = Replace(Sheet1.combYear.value, "-", "")
    warrantyCostFile1 = Dir(ThisWorkbook.Path & "\" & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb")
    warrantyCostFile2 = Dir(ThisWorkbook.Path & "\" & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb")
    found = InStr(1, warrantyCostFile1, valFind, vbTextCompare)
    found = InStr(1, warrantyCostFile2, valFind, vbTextCompare)
    
    If found = "0" Then
    fileNotPresent(8) = "Input File with name " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb" & Chr(34) & " or " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb" & Chr(34) & vbCrLf & " doesn't exist!"
    End If
    
    'For CQ Data File Validation
    CQDataFile = ""
    CQDataFile = Dir(ThisWorkbook.Path & "\" & "CQ_Data_SPM.xlsx")
    
    If CQDataFile = "" Then
    fileNotPresent(9) = "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "CQ_Data_SPM.xlsx" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End If
    
    'For service information Data File Validation
    serviceInfoDataFile = ""
    serviceInfoDataFile = Dir(ThisWorkbook.Path & "\" & "Service_Information_Quality_Completion.xlsx")
    
    If CQDataFile = "" Then
    fileNotPresent(10) = "Input File name format does not correspond to Selected Month and Year!" & vbCrLf & vbCrLf & "File with Name " & Chr(34) & "Service_Information_Quality_Completion.xlsx" & Chr(34) & " Not Found" & vbCrLf & vbCrLf & "Please select the appropriate date or change the input file!"
    End If

Dim flNt As Integer
Dim msg1
For flNt = 1 To 10
    If fileNotPresent(flNt) <> "" Then
        msg1 = MsgBox("Following Files are not Present, Do you want to Continue?" & vbCrLf & fileNotPresent(1) & vbCrLf & fileNotPresent(2) & fileNotPresent(3) & vbCrLf & fileNotPresent(4) & vbCrLf _
                & fileNotPresent(5) & vbCrLf & fileNotPresent(6) & vbCrLf & fileNotPresent(7) & vbCrLf & fileNotPresent(8) & vbCrLf _
                & fileNotPresent(9) & vbCrLf & fileNotPresent(10), vbYesNo)
        If msg1 = vbNo Then
        End
        End If
        Exit For
    End If
Next
End If

    
'validation for files present over shared drive
If Sheet1.rdbSharedDrive.value = True Then
    
    Dim fileNotFoundShared(10) As String
    
    SharedDrive_Path "KPI Summary.xlsx"
    
    If fileExists = False Then
        MsgBox "Output File with name " & Chr(34) & "KPI Summary.xlsx" & Chr(34) & " doesn't exist!"
        End
    End If
    
    'for service scorecard
    fstMonthChk = Format(Sheet1.combYear.value, "mmmyy")
    SharedDrive_Path ("Service Scorecard F 6.1_" & fstMonthChk & ".xlsm")
    
    If fileExists = False Then
    fileNotFoundShared(1) = Chr(34) & "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm" & Chr(34)
    End If
    
    'for innovation file
    innovationFileName = "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx"
    SharedDrive_Path innovationFileName
    
    If fileExists = False Then
    fileNotFoundShared(2) = Chr(34) & "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" & Chr(34)
    
    End If

    'checking Install Hrs file exists
    installFileOpen = "Install SPAN P95.xlsx"
    SharedDrive_Path installFileOpen
    
    If fileExists = False Then
        fileNotFoundShared(3) = Chr(34) & "Install SPAN P95.xlsx" & Chr(34)
        
    End If

    'Checking for FCO OP review file.xlsx
    fcoFileOpen = "FCO OP review file.xlsx"
    SharedDrive_Path fcoFileOpen
    
    If fileExists = False Then
        fileNotFoundShared(4) = Chr(34) & "FCO OP review file.xlsx" & Chr(34)
        
    End If
    
    'Escalations_Overview_ALL BIUs.xlsx
    escOFileOpen = "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx"
    SharedDrive_Path escOFileOpen
    
    If fileExists = False Then
        fileNotFoundShared(5) = Chr(34) & "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx" & Chr(34)
        
    End If
    
    'Customer escalations (Weekly Review) Complaints.xlsx
    compOFileOpen = "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx"
    SharedDrive_Path compOFileOpen
    
    If fileExists = False Then
        fileNotFoundShared(6) = Chr(34) & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" & Chr(34)
        
    End If
    
    '2015-05 Installation spend L2-report.xlsb
    inscostFileOpen = Sheet1.combYear.value & " " & "Installation spend L2-report" & ".xlsb"
    SharedDrive_Path inscostFileOpen
    
    If fileExists = False Then
        fileNotFoundShared(7) = Chr(34) & Sheet1.combYear.value & " " & "Installation spend L2-report" & ".xlsb" & Chr(34)
        
    End If
    
    'warranty cost file
    
    valFind = Replace(Sheet1.combYear.value, "-", "")
    warrantyCostFile1 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb"
    warrantyCostFile2 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb"
    SharedDrive_Path warrantyCostFile1
    SharedDrive_Path warrantyCostFile2
    
    If fileExists = False Then
    fileNotFoundShared(8) = Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb" & Chr(34) & " or " & vbCrLf & Chr(34) & "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb" & Chr(34)
    
    End If
    
    'For CQ Data File Validation
    CQDataFile = "CQ_Data_SPM.xlsx"
    SharedDrive_Path CQDataFile
    
    If fileExists = False Then
    fileNotFoundShared(9) = Chr(34) & "CQ_Data_SPM.xlsx" & Chr(34)
    
    End If
    
    'For service information Data File Validation
    serviceInfoDataFile = "Service_Information_Quality_Completion.xlsx"
    SharedDrive_Path "Service_Information_Quality_Completion.xlsx"
    
    If fileExists = False Then
    fileNotFoundShared(10) = Chr(34) & "Service_Information_Quality_Completion.xlsx" & Chr(34)
    
    End If

For flNt = 1 To 10
    If fileNotFoundShared(flNt) <> "" Then
        msg1 = MsgBox("Input File name format does not correspond to Selected Month and Year!, Do you want to Continue?" & vbCrLf & fileNotFoundShared(1) & vbCrLf & fileNotFoundShared(2) & fileNotFoundShared(3) & vbCrLf & fileNotFoundShared(4) & vbCrLf _
                & fileNotFoundShared(5) & vbCrLf & fileNotFoundShared(6) & vbCrLf & fileNotFoundShared(7) & vbCrLf & fileNotFoundShared(8) & vbCrLf _
                & fileNotFoundShared(9) & vbCrLf & fileNotFoundShared(10), vbYesNo)
        If msg1 = vbNo Then
        End
        End If
        Exit For
    End If
Next
    
End If
End Sub
    
