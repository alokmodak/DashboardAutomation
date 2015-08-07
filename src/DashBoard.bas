Attribute VB_Name = "DashBoard"
'**********************************************************************************************************************
'* Code for DashBoard Automation
'* 'Date           Who     What
'*
'**********************************************************************************************************************

Option Explicit
Public mnthNot(10) As String
Public mnthNt As Integer 'integer for array month not present
Public shtNotPresent(20) As String
Public shtNt As Integer
Public outputFileGlobal As String



Public Sub Generate_Dashboard_Output()

On Error Resume Next

outputFileGlobal = "KPI Summary.xlsx"
Dim NewWkbk As Workbook 'new workbook
Dim inputFlName As String 'input file name
Dim outputFlName As String 'output file name
Dim selYear As String
Dim fstAddress As String
Dim lstAddress As String
Dim celItem As Range
Dim valToFind As String
Dim printValue As String
Dim monthVal As String
Dim inputFstAdd As String
Dim inputLstAdd As String
Dim i As Integer
Dim valToPaste As String
Dim flag As Integer
Dim KPISheetName As String
Dim outputPath As String
Dim flNameCheckDate As String
Dim selectedDate As String
Dim j As Integer
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selectSheet As Integer 'flag for sheet not found

selectSheet = 0
'for insCs
Dim TestStr As String
Dim insCFl As String

'for install hrs
Dim inputFile As String
Dim outputFile As String
Dim insFindValue As String
Dim insFilterValue1 As String
Dim p As PivotTable
Dim pf As PivotField
Dim pfi As PivotItem
Dim pvtName As String
Dim startDate As String
Dim endDate As String
Dim insFilterValue2 As String
Dim insFilterValue3 As String, insFilterValue4 As String, insFilterValue5 As String
Dim installPasteValue As String
Dim inputItem As String
Dim myWorkBook As String
Dim installFileOpen As String
Dim productItem As Variant 'for loop for each product
Dim fstMonthChk As String
Dim installFlName As String

shtNt = 1 'sheet not present array
mnthNt = 1 ' month/year not present in input file array
fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")
yrSelectedFirst = Sheet1.combYear.Value

'declaring output path
If Sheet1.rdbLocalDrive.Value = True Then
outputPath = ThisWorkbook.Path & "\" & outputFileGlobal
outputFlName = outputFileGlobal
End If

If Sheet1.rdbSharedDrive.Value = True Then
SharedDrive_Path outputFileGlobal
outputPath = sharedDrivePath
outputFlName = outputFileGlobal
End If

Application.Workbooks.Open (outputPath), False
Application.Workbooks(outputFileGlobal).Windows(1).Visible = True

'Open service scorecard file and install file

inputItem = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*") 'input file path
'skipping if file not present
Dim serviceFileNotPresent As Boolean
serviceFileNotPresent = False
If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*") = "" Then
    serviceFileNotPresent = True
    End If
End If

installFlName = ThisWorkbook.Path & "\" & "Install SPAN P95.xlsx"
'skipping if file not present
Dim insFileNotPresent As Boolean
insFileNotPresent = False
If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "Install SPAN P95.xlsx") = "" Then
    insFileNotPresent = True
    End If
End If

installFileOpen = "Install SPAN P95.xlsx"
inputFlName = "Service Scorecard F 6.1_" & fstMonthChk & ".xlsm"
'inputFlName = Dir(ThisWorkbook.Path & "\" & "Service Scorecard F 6.1_" & fstMonthChk & "*.xls*") 'Input file name declared
    
    'for Shared drive define path
    If Sheet1.rdbSharedDrive.Value = True Then
        SharedDrive_Path inputFlName
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> inputFlName Then
            serviceFileNotPresent = True
        End If
        inputItem = sharedDrivePath
        SharedDrive_Path installFileOpen
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> installFileOpen Then
            insFileNotPresent = True
        End If
        installFlName = sharedDrivePath
    End If
    
'exit if both file are not present
If insFileNotPresent = True And serviceFileNotPresent = True Then
     Exit Sub
End If

    Application.Workbooks.Open (installFlName), False
    Application.Workbooks.Open (inputItem), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets("Data Analysis Pivot").Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find("Column Labels").Select
    
Application.Workbooks(myWorkBook).Windows(1).Visible = False
Application.Workbooks(installFileOpen).Windows(1).Visible = True

'Filtering servicescorecard data based on selection
Dim productGroup As String

productGroup = Sheet1.combProductGroup.Value

'starting for loop for each item in product group
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        productGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If productGroup = "" Then
            Exit For
            End If
    End If
    

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFlName).Activate
For i = 1 To Workbooks(outputFlName).Sheets.Count
    If Workbooks(outputFlName).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

Workbooks(inputFlName).Activate
Worksheets("Data Analysis Pivot").Activate
ActiveSheet.Cells(5, 2).Select
pvtName = ActiveCell.PivotTable.name

'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)
'Unhide page field pivot items
For Each pf In p.PageFields
If pf = "Product Group" Then
    pf.CurrentPage = productGroup
End If
Next pf

Application.Calculation = xlCalculationAutomatic  'Enabling automatic calculations
Worksheets("Service Scorecard").Activate
ActiveSheet.Cells(8, 2).Select
pvtName = ActiveCell.PivotTable.name

Set p = ActiveSheet.PivotTables(pvtName) 'for YTD values
For Each pf In p.PageFields
If pf = "Product Group" Then
    pf.CurrentPage = productGroup
End If
Next pf

selYear = Sheet1.combYear.Value

Application.ScreenUpdating = False
Application.DisplayAlerts = False


'get input file name
Workbooks(inputFlName).Activate
Worksheets("Data Analysis Pivot").Activate
ActiveSheet.Cells(5, 2).Select

'selecting month value for output file
monthVal = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

Do Until j = 0

Select Case j

Case 1
monthVal = "Jan"
Case 2
monthVal = "Feb"
Case 3
monthVal = "Mar"
Case 4
monthVal = "Apr"
Case 5
monthVal = "May"
Case 6
monthVal = "Jun"
Case 7
monthVal = "Jul"
Case 8
monthVal = "Aug"
Case 9
monthVal = "Sep"
Case 10
monthVal = "Oct"
Case 11
monthVal = "Nov"
Case 12
monthVal = "Dec"

End Select

'month cell distance value as i
Workbooks(outputFlName).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.Cells(2, 1).Select
i = 0
Do Until ActiveCell.Value = monthVal
ActiveCell.Offset(0, 1).Select
i = i + 1
If i = 200 Then
Exit Do
End If
Loop

Dim inputFindMonth As String
inputFindMonth = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & CStr(Format(j, "00"))
Sheet1.combYear.Value = inputFindMonth

'exit service scorecard file if not present
If serviceFileNotPresent = True Then
GoTo serviceFlNt
End If

'copy values
Workbooks(inputFlName).Activate
Worksheets("Data Analysis Pivot").Activate
ActiveSheet.UsedRange.Find("Row Labels").Select
ActiveCell.Offset(1, 0).Select
inputFstAdd = ActiveCell.Address
ActiveCell.End(xlDown).Select
inputLstAdd = ActiveCell.Address
ActiveSheet.Range(inputFstAdd, inputLstAdd).Select
Selection.Copy
    
'paste values
Workbooks(outputFlName).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.Cells(200, 27).Select
ActiveSheet.Paste
    
'copy month values
inputFindMonth = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & CStr(Format(j, "00"))
Sheet1.combYear.Value = inputFindMonth
Workbooks(inputFlName).Activate
Worksheets("Data Analysis Pivot").Activate
ActiveWorkbook.ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
    
    
If inputFindMonth <> Sheet1.combYear.Value Then
MsgBox inputFindMonth
End If
    
'messagebox for year and month not available

If ActiveCell.Value <> Sheet1.combYear.Value Then
mnthNot(mnthNt) = "The Month/Year - " & Sheet1.combYear.Value & vbCrLf & vbCrLf & "is not present in the input file- " & inputFlName & vbCrLf & vbCrLf & "For Product- " & Sheet1.combProductGroup.Value
mnthNt = mnthNt + 1
GoTo MonthNotAvailable
End If
    
ActiveCell.Offset(1, 0).Select
inputFstAdd = ActiveCell.Address
ActiveCell.End(xlDown).Select
inputLstAdd = ActiveCell.Address
ActiveSheet.Range(inputFstAdd, inputLstAdd).Select
Selection.Copy

'paste month values
Workbooks(outputFlName).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.Cells(200, 28).Select
ActiveSheet.Paste
ActiveSheet.Cells(200, 27).Select

flag = 200
For Each celItem In Range(ActiveCell.Address, ActiveCell.End(xlDown).Address)
ActiveSheet.Cells(flag, 27).Select
valToFind = ActiveCell.Value
valToPaste = ActiveCell.Offset(0, 1).Value

'case to find difference in the values
Select Case valToFind
    Case "Contract GM"
    valToFind = "iGM%"
    Case "CR IW"
    valToFind = "Call Rate (IW)"
    Case "CR OoW"
    valToFind = "Call Rate (OoW Contract)"
    Case "Remote Resolution"
    valToFind = "RRR"
    Case "FTF"
    valToFind = "First Time Fix Rate"
    Case "UPRR"
    valToFind = "Unused Part Return Rate"
End Select

'putting values in output file
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(monthVal).Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.Offset(0, -(i - 1)).Value = ""
If ActiveCell.Offset(0, -(i - 1)).Value = valToFind Then
ActiveCell.Value = valToPaste
Exit Do
Else
ActiveCell.Offset(1, 0).Select
End If
Loop
flag = flag + 1
Next

serviceFlNt:
If insFileNotPresent = True Then
    GoTo insFileNt
End If
    
'Install hrs process
inputFile = installFileOpen

Workbooks(inputFile).Activate
ActiveWorkbook.Sheets("Install SPAN").Activate
ActiveSheet.UsedRange.Find("Period").Select
insFindValue = Sheet1.combProductGroup.Value
pvtName = ActiveCell.PivotTable.name

Select Case insFindValue
    
Case "IXR-MOS Pulsera-Y"
insFilterValue1 = pulseraSysCode1
insFilterValue2 = pulseraSysCode2
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "IXR-MOS BV Vectra-N"
insFilterValue1 = vectraSysCode
insFilterValue2 = ""
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "IXR-MOS Endura-Y"
insFilterValue1 = enduraSysCode1
insFilterValue2 = enduraSysCode2
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "IXR-MOS Veradius-Y"
insFilterValue1 = veradiusSysCode1
insFilterValue2 = veradiusSysCode2
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "IXR-MOS Libra-N"
insFilterValue1 = ""
insFilterValue2 = ""
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "DXR-PrimaryDiagnost Digital-N"
insFilterValue1 = PDSysCode1
insFilterValue2 = ""
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

Case "DXR-MicroDose Mammography-Y"
insFilterValue1 = mamoSysCode1
insFilterValue2 = mamoSysCode2
insFilterValue3 = mamoSysCode3
insFilterValue4 = mamoSysCode4
insFilterValue5 = mamoSysCode5

Case "DXR-MobileDiagnost Opta-N"
insFilterValue1 = optaSysCode1
insFilterValue2 = ""
insFilterValue3 = ""
insFilterValue4 = ""
insFilterValue5 = ""

End Select

'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)

For Each pfi In p.PivotFields("System").PivotItems
pfi.Visible = True
Next pfi

'selecting values for product group
    For Each pfi In p.PivotFields("System").PivotItems
        If pfi = insFilterValue1 Then
            Debug.Print pfi, insFilterValue1, insFilterValue2, insFilterValue3, insFilterValue4, insFilterValue5
                pfi.Visible = True
        ElseIf pfi = insFilterValue2 Then
                pfi.Visible = True
        ElseIf pfi = insFilterValue3 Then
                pfi.Visible = True
        ElseIf pfi = insFilterValue4 Then
                pfi.Visible = True
        ElseIf pfi = insFilterValue5 Then
                pfi.Visible = True
        Else
                pfi.Visible = False
        End If
    Next pfi
              
For Each pf In p.PageFields
If pf = "Period" Then
    pf.CurrentPage = Sheet1.combYear.Value
End If
Next pf
              
ActiveSheet.UsedRange.Find("INHrs").Select
ActiveCell.Offset(1, 0).Select

installPasteValue = Application.Average(Range(ActiveCell.Address, ActiveCell.End(xlDown).Address))

'putting value in output file

Workbooks(outputFlName).Activate
ActiveSheet.Cells(2, 2).Select
i = 0
Do Until ActiveCell.Value = monthVal
    ActiveCell.Offset(0, 1).Select
    i = i + 1
        If i = 200 Then
            MsgBox "Month not found! Please try again!"
            End
        End If
Loop

Do Until ActiveCell.Offset(0, -i).Value = "Install Hours"
ActiveCell.Offset(1, 0).Select
Loop
        
If ActiveCell.Offset(0, -i).Value = "Install Hours" Then
    ActiveCell.Value = installPasteValue
    ActiveCell.Offset(1, 0).Select
End If

insFileNt:
j = j - 1
Loop
Sheet1.combYear.Value = yrSelectedFirst

MonthNotAvailable: 'if month is not available in input file
If insFileNotPresent = True Then
    GoTo insFileNt2
End If

'YTD for Install hrs
startDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & "01"
endDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-")))
    
Set p = ActiveSheet.PivotTables("PivotTable1")
   
For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

For Each pfi In p.PivotFields("Period").PivotItems
    If pfi < startDate Or pfi > endDate Then
       ' Debug.Print DateValue(pfi.Name), StartDate, EndDate
            pfi.Visible = False
    Else
            pfi.Visible = True
    End If
Next pfi

Workbooks(inputFile).Activate
ActiveWorkbook.Sheets("Install SPAN").Activate
ActiveSheet.UsedRange.Find("INHrs").Select
ActiveCell.Offset(1, 0).Select

Dim fstAdd As String
Dim lstAdd As String
fstAdd = ActiveCell.Address
lstAdd = ActiveCell.End(xlDown).Address

Dim YTDinstallPasteValue As String
YTDinstallPasteValue = Application.Average(Range(fstAdd, lstAdd))

Workbooks(outputFlName).Activate
ActiveSheet.UsedRange.Find(What:="Install Hours", lookat:=xlWhole).Select

Do Until ActiveCell.End(xlUp).Value = "YTD"

        ActiveCell.Offset(0, 1).Select
i = i + 1
If i = 200 Then
Exit Do
End If
Loop

If ActiveCell.End(xlUp).Value = "YTD" Then
        ActiveCell.Value = YTDinstallPasteValue
End If

If serviceFileNotPresent = True Then
GoTo serviceFlNt2
End If

insFileNt2:
'clear values fleched from input file
ActiveSheet.Cells(200, 27).Select
ActiveSheet.Range(ActiveCell.Address, ActiveCell.Offset(0, 1).End(xlDown).Address).Clear
ActiveSheet.Cells(2, 2).Select


'Fletching YTD values
i = 0
Do Until ActiveCell.Value = "YTD"
i = i + 1
If i = 100 Then
MsgBox "YTD Column is not available - Please check the output file"
Workbooks(outputFlName).Save
Exit Sub
End If

ActiveCell.Offset(0, 1).Select
Loop

Dim YTDFindValue As String, YTDPasteValue As String, firstCell As String, lastCell As String
Dim YTDFlag As Integer
YTDFlag = 0

'Calculating range
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -i).Select
firstCell = ActiveCell.Address
lastCell = ActiveCell.End(xlDown).Address
ActiveCell.Offset(0, i).Select
    
For Each celItem In Range(firstCell, lastCell)
    
    YTDFindValue = celItem
    Workbooks(inputFlName).Activate
    ActiveWorkbook.Sheets("Service Scorecard").Activate
    ActiveSheet.Cells(1, 1).Select
    
    Select Case YTDFindValue
    
    Case "Contract Revenue"
        ActiveSheet.UsedRange.Find("Customer").Select
        ActiveSheet.UsedRange.Find(What:=YTDFindValue, After:=ActiveCell, LookIn:=xlValues).Select
         i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
         
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
    
    Case "iGM%"
        ActiveSheet.UsedRange.Find("Customer").Select
        ActiveSheet.UsedRange.Find(What:="Contract Profitability - Gross Margin %", After:=ActiveCell, LookIn:=xlValues).Select
         i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
    
        YTDPasteValue = ActiveCell.Value
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "Contract Penetration"
        ActiveSheet.UsedRange.Find("Contract Penetration").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "IB Count Contract"
        ActiveSheet.UsedRange.Find("# Contracts").Select
            ActiveCell.Offset(0, 1).Select
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "Call Rate (IW)"
        ActiveSheet.UsedRange.Find("IW Call Rate (CM calls p/s)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "Call Rate (OoW Contract)"
        ActiveSheet.UsedRange.Find("OoW Call Rate (CM calls p/s)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "MTTR"
        ActiveSheet.UsedRange.Find("MTTR (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "RRR"
        ActiveSheet.UsedRange.Find("MTTR (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Offset(1, 0).Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "First Time Fix Rate"
        ActiveSheet.UsedRange.Find("MTTR (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Offset(2, 0).Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "CM Labor"
        ActiveSheet.UsedRange.Find("Corrective Maintenance Labor (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue * CStr(Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-"))))
        YTDPasteValue = ""
        
    
    Case "CM Material"
        ActiveSheet.UsedRange.Find("Corrective Maintenance Material Cost").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue * CStr(Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-"))))
        YTDPasteValue = ""
        
    Case "PM Labor"
        ActiveSheet.UsedRange.Find("Preventive Maintenance Labor (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Value
        YTDPasteValue = YTDPasteValue * CStr(Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-"))))
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
        
    Case "Unused Part Return Rate"
        ActiveSheet.UsedRange.Find("MTTR (hrs)").Select
        i = 0
            Do Until ActiveCell.End(xlUp).Value = "YTD"
            i = i + 1
               If i = 100 Then
               Exit Sub
               End If
            ActiveCell.Offset(0, 1).Select
            Loop
            
        YTDPasteValue = ActiveCell.Offset(3, 0).Value
        YTDPasteValue = YTDPasteValue
        Workbooks(outputFlName).Activate
        ActiveCell.Value = YTDPasteValue
        YTDPasteValue = ""
    
    Case Else
        Workbooks(outputFlName).Activate ' else case added if cell value has space.
        
    End Select
    ActiveCell.Offset(1, 0).Select
Next

serviceFlNt2:
sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

Next productItem 'for loop for each product end

Workbooks(outputFlName).Save
Workbooks(inputFlName).Close False
Workbooks(installFileOpen).Close False

'getting original date value back
Sheet1.combYear.Value = yrSelectedFirst
End Sub

Public Sub Calculate_Innovation()

On Error Resume Next
Dim inputFl As String
Dim outputFl As String
Dim patternProductGroup As String
Dim patternDate As String
Dim i As Integer
Dim patternValToPaste As String
Dim dapValToPaste As String
Dim remoteValToPaste As String
Dim patternFstAdd As String
Dim patternLstAdd As String
Dim j As Integer
Dim productItem As Variant
Dim KPISheetName As String
Dim selectSheet As Integer
Dim YTDValToFind As String
Dim actAdd As String
Dim cell As Integer
Dim fstMonthChk As String

fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")

outputFl = outputFileGlobal
inputFl = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "KPI dashboard_Innovation_" & fstMonthChk & "*.xl*")

'Skipping if inout file not present
If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "KPI dashboard_Innovation_" & fstMonthChk & "*.xl*") = "" Then
        Exit Sub
    End If
End If
    
    'for Shared drive define path
    If Sheet1.rdbSharedDrive.Value = True Then
        SharedDrive_Path ("KPI dashboard_Innovation_" & fstMonthChk & ".xlsx")
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> "KPI dashboard_Innovation_" & fstMonthChk & ".xlsx" Then
            Exit Sub
        End If
        inputFl = sharedDrivePath
    End If

Application.Workbooks.Open (inputFl)
inputFl = ActiveWorkbook.name

patternProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        patternProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If patternProductGroup = "" Then
            Exit For
            End If
    End If
'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value
i = 0
patternDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)
selectSheet = 0

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

Do Until j = 0 ' loop for each month
Select Case j

    Case 1
    patternDate = "Jan"
    Case 2
    patternDate = "Feb"
    Case 3
    patternDate = "Mar"
    Case 4
    patternDate = "Apr"
    Case 5
    patternDate = "May"
    Case 6
    patternDate = "Jun"
    Case 7
    patternDate = "Jul"
    Case 8
    patternDate = "Aug"
    Case 9
    patternDate = "Sep"
    Case 10
    patternDate = "Oct"
    Case 11
    patternDate = "Nov"
    Case 12
    patternDate = "Dec"
End Select

Select Case patternProductGroup

Case "IXR-MOS Endura-Y"
YTDValToFind = "Endura"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "IXR-MOS Pulsera-Y"
YTDValToFind = "Pulsera"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "IXR-MOS Veradius-Y"
YTDValToFind = "Veradius"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "IXR-MOS BV Vectra-N"
YTDValToFind = "BV Vectra"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "DXR-PrimaryDiagnost Digital-N"
YTDValToFind = "Primary Diagnost"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "DXR-MobileDiagnost Opta-N"
YTDValToFind = "Mobile Opta 1.0"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

Case "DXR-MicroDose Mammography-Y"
YTDValToFind = "Microdose"
patternValToPaste = ""
remoteValToPaste = ""
dapValToPaste = ""

End Select

'for month values

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find("Product").Select
actAdd = ActiveSheet.UsedRange.Find("Product").Address

For cell = 1 To 3
ActiveSheet.Range(actAdd).Select
ActiveCell.Offset(1, 0).Select
    i = 0
    Do Until ActiveCell.Value = YTDValToFind
    ActiveCell.Offset(1, 0).Select
    i = i + 1
    If i > 200 Then
    Exit Do
    End If
    Loop
    
    If ActiveCell.Offset(0, -1).Value = "Patterns" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = patternDate
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        patternValToPaste = ActiveCell.Value
        
    ElseIf ActiveCell.Offset(0, -1).Value = "Dap Capability" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = patternDate
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        dapValToPaste = ActiveCell.Value
        
    ElseIf ActiveCell.Offset(0, -1).Value = "Remote Capability" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = patternDate
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        remoteValToPaste = ActiveCell.Value
    End If

Next
    
'for pattern values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("# of Patterns").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(patternDate).Select
ActiveSheet.UsedRange.Find(What:=patternDate, After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = patternValToPaste

'for DAP values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("DAP capability").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(patternDate).Select
ActiveSheet.UsedRange.Find(What:=patternDate, After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = dapValToPaste

'for Remote values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("Remote Capability").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(patternDate).Select
ActiveSheet.UsedRange.Find(What:=patternDate, After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = remoteValToPaste

j = j - 1 'loop for each month
Loop

'for YTD value
patternValToPaste = ""
dapValToPaste = ""
remoteValToPaste = ""

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find("Product").Select
actAdd = ActiveSheet.UsedRange.Find("Product").Address

For cell = 1 To 3
ActiveSheet.Range(actAdd).Select
ActiveCell.Offset(1, 0).Select
    i = 0
    Do Until ActiveCell.Value = YTDValToFind
    ActiveCell.Offset(1, 0).Select
    actAdd = ActiveCell.Address
    i = i + 1
    If i > 200 Then
    Exit Do
    End If
    Loop
    
    If ActiveCell.Offset(0, -1).Value = "Patterns" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = "YTD"
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        patternValToPaste = ActiveCell.Value
        
    ElseIf ActiveCell.Offset(0, -1).Value = "Dap Capability" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = "YTD"
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        dapValToPaste = ActiveCell.Value
        
    ElseIf ActiveCell.Offset(0, -1).Value = "Remote Capability" Then
    actAdd = ActiveCell.Address
        i = 0
        Do Until ActiveCell.End(xlUp).Value = "YTD"
        ActiveCell.Offset(0, 1).Select
        i = i + 1
        If i > 200 Then
        Exit Do
        End If
        Loop
        remoteValToPaste = ActiveCell.Value
    End If

Next
    
'for pattern values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("# of Patterns").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find("YTD").Select
ActiveSheet.UsedRange.Find(What:="YTD", After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = patternValToPaste

'for DAP values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("DAP capability").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find("YTD").Select
ActiveSheet.UsedRange.Find(What:="YTD", After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = dapValToPaste

'for Remote values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find("Remote Capability").Select
patternFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find("YTD").Select
ActiveSheet.UsedRange.Find(What:="YTD", After:=ActiveCell, LookIn:=xlValues).Select
ActiveCell.Offset(patternFstAdd - 2, 0).Select
ActiveCell.Value = remoteValToPaste

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

patternValToPaste = ""
dapValToPaste = ""
remoteValToPaste = ""

Next productItem 'for all product groups

Workbooks(inputFl).Close False
End Sub

Public Sub Complaints_Escalations_Calculation()

On Error Resume Next
Dim inputFl As String
Dim escInputFl As String
Dim outputFl As String
Dim complaintsProductGroup As String
Dim cProductGroup As String
Dim complaintsDate As String
Dim i As Integer
Dim complaintsValToPaste As String
Dim complaintsFstAdd As String
Dim complaintsLstAdd As String
Dim j As Integer
Dim productItem As Variant
Dim p As PivotTable
Dim pf As PivotField
Dim pfi As PivotItem
Dim pvtName As String
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim p95ValToPaste As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim escValToPaste As String
Dim escp95ValToPaste As String
Dim fstMonthChk As String

fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")

yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal

 'for Shared drive define path
    If Sheet1.rdbSharedDrive.Value = True Then
        SharedDrive_Path ("Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx")
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx" Then
            Exit Sub
        End If
        inputFl = sharedDrivePath
        SharedDrive_Path ("Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx")
        escInputFl = sharedDrivePath
    Else
    escInputFl = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "Escalations_Overview_ALL BIUs_" & fstMonthChk & "*.xls*")
    inputFl = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & "*.xls*")
        If Dir(ThisWorkbook.Path & "\" & "Escalations_Overview_ALL BIUs_" & fstMonthChk & "*.xls*") = "" Or Dir(ThisWorkbook.Path & "\" & "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & "*.xls*") = "" Then
            Exit Sub
        End If
    End If

Application.Workbooks.Open (inputFl)
Application.Workbooks.Open (escInputFl)

escInputFl = "Escalations_Overview_ALL BIUs_" & fstMonthChk & ".xlsx"
inputFl = "Customer escalations (Weekly Review) Complaints_" & fstMonthChk & ".xlsx"

complaintsProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        complaintsProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If complaintsProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
complaintsDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

Do Until j = 0 'for all months
Select Case j

    Case 1
    complaintsDate = "Jan"
    selMonth = "01"
    Case 2
    complaintsDate = "Feb"
    selMonth = "02"
    Case 3
    complaintsDate = "Mar"
    selMonth = "03"
    Case 4
    complaintsDate = "Apr"
    selMonth = "04"
    Case 5
    complaintsDate = "May"
    selMonth = "05"
    Case 6
    complaintsDate = "Jun"
    selMonth = "06"
    Case 7
    complaintsDate = "Jul"
    selMonth = "07"
    Case 8
    complaintsDate = "Aug"
    selMonth = "08"
    Case 9
    complaintsDate = "Sep"
    selMonth = "09"
    Case 10
    complaintsDate = "Oct"
    selMonth = "10"
    Case 11
    complaintsDate = "Nov"
    selMonth = "11"
    Case 12
    complaintsDate = "Dec"
    selMonth = "12"
End Select

Select Case complaintsProductGroup

'for Endura
Case "IXR-MOS Endura-Y"
cProductGroup = "Endura"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate
ActiveSheet.UsedRange.Find("Period").Select

pvtName = ActiveCell.PivotTable.name
'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)
'Unhide page field pivot items

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
    If pf = "Product" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

For Each pf In p.PageFields
    If pf = "Product" Then
        pf.CurrentPage = cProductGroup
    End If
    
    If pf = "Period" Then
        pf.CurrentPage = Mid(yrSelectedFirst, 1, 4) & "-" & selMonth
    End If
Next pf

ActiveSheet.UsedRange.Find("#Open Complaints").Select
Dim compValToFind As Integer
Dim toMinusVal As Integer
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
Dim compValToFindp95 As Integer
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
ActiveCell.End(xlToRight).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

' for YTD values
'selecting values for Period
Dim startDate As String
Dim endDate As String
startDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & "01"
endDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-")))
    
For Each pfi In p.PivotFields("Period").PivotItems
    If pfi < startDate Or pfi > endDate Then
        Debug.Print pfi, startDate, endDate
            pfi.Visible = False
    Else
            pfi.Visible = True
    End If
Next pfi

ActiveSheet.Cells(2, 2).Select
ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste


'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
Dim escValToFind As String
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'fpr Pulsera
Case "IXR-MOS Pulsera-Y"
cProductGroup = "Pulsera"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate
ActiveSheet.UsedRange.Find("Period").Select

pvtName = ActiveCell.PivotTable.name
'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)
'Unhide page field pivot items

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
    If pf = "Product" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

For Each pf In p.PageFields
    If pf = "Product" Then
        pf.CurrentPage = cProductGroup
    End If
    
    If pf = "Period" Then
        pf.CurrentPage = Mid(yrSelectedFirst, 1, 4) & "-" & selMonth
    End If
Next pf

ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
ActiveCell.End(xlToRight).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

' for YTD values
'selecting values for Period
startDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & "01"
endDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-")))
    
For Each pfi In p.PivotFields("Period").PivotItems
    If pfi < startDate Or pfi > endDate Then
        Debug.Print pfi, startDate, endDate
            pfi.Visible = False
    Else
            pfi.Visible = True
    End If
Next pfi

ActiveSheet.Cells(2, 2).Select
ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste


'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'for Veradius
Case "IXR-MOS Veradius-Y"
cProductGroup = "Veradius"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate
ActiveSheet.UsedRange.Find("Period").Select

pvtName = ActiveCell.PivotTable.name
'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)
'Unhide page field pivot items

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
    If pf = "Product" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

For Each pf In p.PageFields
    If pf = "Product" Then
        pf.CurrentPage = cProductGroup
    End If
    
    If pf = "Period" Then
        pf.CurrentPage = Mid(yrSelectedFirst, 1, 4) & "-" & selMonth
    End If
Next pf

ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
ActiveCell.End(xlToRight).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

' for YTD values
'selecting values for Period
startDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & "01"
endDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-")))
    
For Each pfi In p.PivotFields("Period").PivotItems
    If pfi < startDate Or pfi > endDate Then
        Debug.Print pfi, startDate, endDate
            pfi.Visible = False
    Else
            pfi.Visible = True
    End If
Next pfi

ActiveSheet.Cells(2, 2).Select
ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste


'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'for BV Libra
Case "IXR-MOS Libra-N"
cProductGroup = "BV Libra"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate
ActiveSheet.UsedRange.Find("Period").Select

pvtName = ActiveCell.PivotTable.name
'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)
'Unhide page field pivot items

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
    If pf = "Product" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

For Each pf In p.PageFields
    If pf = "Product" Then
        pf.CurrentPage = cProductGroup
    End If
    
    If pf = "Period" Then
        pf.CurrentPage = Mid(yrSelectedFirst, 1, 4) & "-" & selMonth
    End If
Next pf

ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
ActiveCell.End(xlToRight).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("MoS Open Complaints").Activate

For Each pf In p.PageFields
    If pf = "Period" Then
        pf.CurrentPage = "(All)"
    End If
Next pf

' for YTD values
'selecting values for Period
startDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & "01"
endDate = Split(Sheet1.combYear.Value, "-")(LBound(Split(Sheet1.combYear.Value, "-"))) & "-" & Split(Sheet1.combYear.Value, "-")(UBound(Split(Sheet1.combYear.Value, "-")))
    
For Each pfi In p.PivotFields("Period").PivotItems
    If pfi < startDate Or pfi > endDate Then
        Debug.Print pfi, startDate, endDate
            pfi.Visible = False
    Else
            pfi.Visible = True
    End If
Next pfi

ActiveSheet.Cells(2, 2).Select
ActiveSheet.UsedRange.Find("#Open Complaints").Select
compValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value

ActiveSheet.UsedRange.Find("p95").Select
compValToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Sheet1.combYear.Value).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(compValToFindp95 - toMinusVal, 0).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste


'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'for BV Vectra
Case "IXR-MOS BV Vectra-N"
cProductGroup = "BV Vectra"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("CHU synop").Activate
ActiveSheet.Cells(5, 5).Select
ActiveSheet.UsedRange.Find(What:=cProductGroup, After:=ActiveCell).Select
Dim chuYearToFind As Integer
chuYearToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value
p95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value


'for Allura FC
Case "IXR-CV Allura FC-Y"

cProductGroup = "Allura FC"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("CHU synop").Activate
ActiveSheet.Cells(5, 5).Select
ActiveSheet.UsedRange.Find(What:=cProductGroup, After:=ActiveCell).Select
chuYearToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value
p95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'for Opta
Case "DXR-MobileDiagnost Opta-N"

cProductGroup = "Digital Radiography - Opta"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("CHU synop").Activate
ActiveSheet.Cells(5, 5).Select
ActiveSheet.UsedRange.Find(cProductGroup).Select
chuYearToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value
p95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

'for Primary Diagnost
Case "DXR-PrimaryDiagnost Digital-N"

cProductGroup = "Primary Diagnost DR"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("CHU synop").Activate
ActiveSheet.Cells(5, 5).Select
ActiveSheet.UsedRange.Find(cProductGroup).Select
chuYearToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value
p95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find(cProductGroup).Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Case "DXR-MicroDose Mammography-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("Mammo Open complaints").Activate
ActiveSheet.Cells(5, 5).Select
ActiveSheet.UsedRange.Find("#Open Complaints").Select
Dim chuYearToFindp95 As Integer
chuYearToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
complaintsValToPaste = ActiveCell.Value
ActiveSheet.UsedRange.Find("p95").Select
chuYearToFindp95 = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(chuYearToFind - toMinusVal, 0).Select
p95ValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = complaintsValToPaste

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = p95ValToPaste

'for escalations
Workbooks(escInputFl).Activate
ActiveWorkbook.Sheets("Open Esc_Product").Activate
ActiveSheet.Cells.EntireRow.Hidden = False
ActiveSheet.UsedRange.Find("Legacy DXR").Select
escValToFind = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(Mid(Sheet1.combYear.Value, 1, 4) & "-" & selMonth).Select
toMinusVal = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveCell.Offset(escValToFind - toMinusVal, 0).Select
escValToPaste = ActiveCell.Value
escp95ValToPaste = ActiveCell.Offset(2, 0).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escValToPaste

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(complaintsDate).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = escp95ValToPaste

'for Escalations YTD values
Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS #", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open viper/ one EMS p95 days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

Workbooks(outputFl).Activate
Worksheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Compliants", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

ActiveSheet.UsedRange.Find(What:="Open Compliants P95 Days", lookat:=xlWhole).Select
complaintsFstAdd = CInt(Mid(ActiveCell.Address, 4, 2))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(complaintsFstAdd - 2, 0).Select
ActiveCell.Value = ActiveCell.End(xlToRight).Value

End Select

j = j - 1 'loop for each month
Loop

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

Next productItem 'for all product groups

Workbooks(inputFl).Close False
Workbooks(escInputFl).Close False

End Sub

Public Sub FCO_Calculations()

On Error Resume Next
Dim inputFl As String
Dim outputFl As String
Dim fcoProductGroup As String
Dim cProductGroup As String
Dim fcoDate As String
Dim i As Integer
Dim j As Integer
Dim productItem As Variant
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim fstMonthChk As String

fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")

yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal

    If Sheet1.rdbSharedDrive.Value = True Then
        SharedDrive_Path ("FCO OP review file.xlsx")
        inputFl = sharedDrivePath
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> "FCO OP review file.xlsx" Then
            Exit Sub
        End If
    Else
    inputFl = Dir(ThisWorkbook.Path & "\" & "FCO OP review file.xlsx")
    If inputFl = "" Then
        Exit Sub
    End If
    inputFl = ThisWorkbook.Path & "\" & "FCO OP review file.xlsx"
    End If
    
Application.Workbooks.Open (inputFl)

inputFl = "FCO OP review file.xlsx"

fcoProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        fcoProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If fcoProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
fcoDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

'Do Until j = 0
Select Case j

    Case 1
    fcoDate = "Jan"
    selMonth = "01"
    Case 2
    fcoDate = "Feb"
    selMonth = "02"
    Case 3
    fcoDate = "Mar"
    selMonth = "03"
    Case 4
    fcoDate = "Apr"
    selMonth = "04"
    Case 5
    fcoDate = "May"
    selMonth = "05"
    Case 6
    fcoDate = "Jun"
    selMonth = "06"
    Case 7
    fcoDate = "Jul"
    selMonth = "07"
    Case 8
    fcoDate = "Aug"
    selMonth = "08"
    Case 9
    fcoDate = "Sep"
    selMonth = "09"
    Case 10
    fcoDate = "Oct"
    selMonth = "10"
    Case 11
    fcoDate = "Nov"
    selMonth = "11"
    Case 12
    fcoDate = "Dec"
    selMonth = "12"
End Select


Select Case fcoProductGroup

'for Endura
Case "IXR-MOS Endura-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find("MoS").Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
Dim fstAdd As String
Dim lstAdd As String
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'fpr Pulsera
Case "IXR-MOS Pulsera-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find("MoS").Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for Veradius
Case "IXR-MOS Veradius-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find("MoS").Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for BV Libra
Case "IXR-MOS Libra-N"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find("MoS").Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for BV Vectra
Case "IXR-MOS BV Vectra-N"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find(What:="BV Vectra", lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for Allura FC
Case "IXR-CV Allura FC-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find(What:="Allura FC", lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for Opta
Case "DXR-MobileDiagnost Opta-N"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find(What:="Opta DR/AR", lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

'for Primary Diagnost
Case "DXR-PrimaryDiagnost Digital-N"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find(What:="Primary Diagnost  DR/AR", lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

Case "DXR-MicroDose Mammography-Y"

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("FCO").Activate
ActiveSheet.UsedRange.Find(What:="Mammography", lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.End(xlDown).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="# Released FCO", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveSheet.Paste
Selection.Font.Size = 18
ActiveCell.Offset(1, 0).Select

End Select

'j = j - 1 'loop for each month
'Loop

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

Next productItem 'for all product groups

Workbooks(inputFl).Close False

End Sub

Public Sub Installation_Cost_Calculation()

On Error Resume Next
Dim inputFl As String
Dim outputFl As String
Dim insCProductGroup As String
Dim cProductGroup As String
Dim insCDate As String
Dim i As Integer
Dim j As Integer
Dim productItem As Variant
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim inputFlOpen As String
Dim p As PivotTable
Dim pf As PivotField
Dim pfi As PivotItem
Dim pvtName As String

yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal
inputFlOpen = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & Sheet1.combYear.Value & " " & "*Installation spend L2-report*" & "*.xls*")
'skipping if input file not present
If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & Sheet1.combYear.Value & " " & "*Installation spend L2-report*" & "*.xls*") = "" Then
        Exit Sub
    End If
End If

    If Sheet1.rdbSharedDrive.Value = True Then
        SharedDrive_Path (Sheet1.combYear.Value & " " & "Installation spend L2-report" & ".xlsb")
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> Sheet1.combYear.Value & " " & "Installation spend L2-report" & ".xlsb" Then
            Exit Sub
        End If
        inputFlOpen = sharedDrivePath
    End If

Application.Workbooks.Open (inputFlOpen)
inputFl = Split(inputFlOpen, "\")(UBound(Split(inputFlOpen, "\")))
Application.Workbooks(inputFl).Windows(1).Visible = False

insCProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        insCProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If insCProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
insCDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value
Dim magNameFlt As String

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1
magNameFlt = "BV Pulsera"

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1
magNameFlt = "BV Vectra"

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1
magNameFlt = "BV Endura"

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1
magNameFlt = "Veradius"

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1
magNameFlt = "Allura FC/FD"

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1
magNameFlt = "PrimaryDiagnost DR"

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1
magNameFlt = "MicroDose SI"

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1
magNameFlt = "MobileDiagnost Opta"

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

Do Until j = 0
Select Case j

    Case 1
    insCDate = "Jan"
    selMonth = "001"
    Case 2
    insCDate = "Feb"
    selMonth = "002"
    Case 3
    insCDate = "Mar"
    selMonth = "003"
    Case 4
    insCDate = "Apr"
    selMonth = "4"
    Case 5
    insCDate = "May"
    selMonth = "5"
    Case 6
    insCDate = "Jun"
    selMonth = "6"
    Case 7
    insCDate = "Jul"
    selMonth = "7"
    Case 8
    insCDate = "Aug"
    selMonth = "8"
    Case 9
    insCDate = "Sep"
    selMonth = "9"
    Case 10
    insCDate = "Oct"
    selMonth = "10"
    Case 11
    insCDate = "Nov"
    selMonth = "11"
    Case 12
    insCDate = "Dec"
    selMonth = "12"
End Select

Dim insCMonth As String

insCMonth = selMonth & "." & Mid(Sheet1.combYear.Value, 1, 4)

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.Cells(11, 2).Select

pvtName = ActiveCell.PivotTable.name

'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)

              
For Each pf In p.PageFields
If pf = "Fiscal year/period" Then
    pf.CurrentPage = insCMonth
End If
Next pf

ActiveSheet.Range("B:B").Find(What:=magNameFlt, lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
Do Until ActiveCell.End(xlUp).Value = "Tot Installation %"
ActiveCell.Offset(0, 1).Select
i = i + 1
If i = 100 Then
Exit Do
End If
Loop
   
Dim insCValToPaste As String
insCValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Installation Cost / ASP", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:=insCDate, lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = insCValToPaste

j = j - 1 'loop for each month
Loop

'for YTD values
Workbooks(inputFl).Activate
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.Cells(11, 2).Select

pvtName = ActiveCell.PivotTable.name

'filtering the data based on selection
Set p = ActiveSheet.PivotTables(pvtName)

              
For Each pf In p.PageFields
If pf = "Fiscal year/period" Then
    pf.CurrentPage = "(All)"
End If
Next pf

ActiveSheet.Range("B:B").Find(What:=magNameFlt, lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
Do Until ActiveCell.End(xlUp).Value = "Tot Installation %"
ActiveCell.Offset(0, 1).Select
i = i + 1
If i = 100 Then
Exit Do
End If
Loop
   
insCValToPaste = ActiveCell.Value

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Installation Cost / ASP", lookat:=xlWhole).Select
i = Mid(ActiveCell.Address, 4, 2)
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = insCValToPaste

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

Next productItem 'for all product groups

Workbooks(inputFl).Close False

End Sub


Public Sub Warranty_Cost_Calculations()

On Error Resume Next
Dim inputFlIGT As String
Dim inputFlDI As String
Dim outputFl As String
Dim warrantyCProductGroup As String
Dim cProductGroup As String
Dim warrantyCDate As String
Dim i As Integer
Dim j As Integer
Dim productItem As Variant
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim inputFlOpenIGT As String
Dim inputFlOpenDI As String
Dim IGTValToPaste As String
Dim DIValToPaste As String
Dim warrantyCostFile1 As String
Dim warrantyCostFile2 As String
Dim valFind As String

valFind = Replace(Sheet1.combYear.Value, "-", "")
yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal
inputFlOpenIGT = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*IGT.xls*")
inputFlOpenDI = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*DI.xls*")

If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*IGT.xls*") = "" Or Dir(ThisWorkbook.Path & "\" & "*Warranty Spend Analysis*" & "*DI.xls*") = "" Then
        Exit Sub
    End If
End If

    If Sheet1.rdbSharedDrive.Value = True Then
        warrantyCostFile1 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_IGT.xlsb"
        warrantyCostFile2 = "Level 4 Warranty Spend Analysis - " & valFind & " @ " & valFind - 1 & " BS Rate_DI.xlsb"
        SharedDrive_Path warrantyCostFile1
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> warrantyCostFile1 Then
            Exit Sub
        End If
        inputFlOpenIGT = sharedDrivePath
        SharedDrive_Path warrantyCostFile2
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> warrantyCostFile2 Then
            Exit Sub
        End If
        inputFlOpenDI = sharedDrivePath
    End If

Application.Workbooks.Open (inputFlOpenDI)
Application.Workbooks.Open (inputFlOpenIGT)

inputFlDI = Split(inputFlOpenDI, "\")(UBound(Split(inputFlOpenDI, "\")))
inputFlIGT = Split(inputFlOpenIGT, "\")(UBound(Split(inputFlOpenIGT, "\")))

Application.Workbooks(inputFlDI).Windows(1).Visible = False
Application.Workbooks(inputFlIGT).Windows(1).Visible = False

warrantyCProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        warrantyCProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If warrantyCProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
warrantyCDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value
Dim magNameFlt As String

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
selectSheet = 1
magNameFlt = "BV Pulsera"

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
selectSheet = 1
magNameFlt = "BV Vectra"

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
selectSheet = 1
magNameFlt = "BV Endura"

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
selectSheet = 1
magNameFlt = "Veradius"

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
selectSheet = 1
magNameFlt = "Allura FC/FD"

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
selectSheet = 1

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
selectSheet = 1
magNameFlt = "PrimaryDiagnost DR"

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
selectSheet = 1

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
selectSheet = 1
magNameFlt = "MobileDiagnost Opta"

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    shtNotPresent(shtNt) = "Sheet with name " & KPISheetName & " Does not exists in KPI Summary.xlsx" & vbCrLf & vbCrLf
    shtNt = shtNt + 1
    GoTo sheetNameNotPresent
End If

Dim YTDnum As Integer
YTDnum = 1 'array for YTD values

Do Until j = 0
Select Case j

    Case 1
    warrantyCDate = "Jan"
    selMonth = "001"
    Case 2
    warrantyCDate = "Feb"
    selMonth = "002"
    Case 3
    warrantyCDate = "Mar"
    selMonth = "003"
    Case 4
    warrantyCDate = "Apr"
    selMonth = "004"
    Case 5
    warrantyCDate = "May"
    selMonth = "005"
    Case 6
    warrantyCDate = "Jun"
    selMonth = "006"
    Case 7
    warrantyCDate = "Jul"
    selMonth = "007"
    Case 8
    warrantyCDate = "Aug"
    selMonth = "008"
    Case 9
    warrantyCDate = "Sep"
    selMonth = "009"
    Case 10
    warrantyCDate = "Oct"
    selMonth = "010"
    Case 11
    warrantyCDate = "Nov"
    selMonth = "011"
    Case 12
    warrantyCDate = "Dec"
    selMonth = "012"
End Select

Dim warrantyCMonth As String

warrantyCMonth = selMonth & "." & Mid(Sheet1.combYear.Value, 1, 4)

Select Case warrantyCProductGroup

Case "DXR-MicroDose Mammography-Y"
Dim mcSysCode1 As Double, mcSysCode2 As Double, mcSysCode3 As Double, mcSysCode4 As Double, mcSysCode5 As Double

Workbooks(inputFlDI).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = mamoSysCode1 Then
    mcSysCode1 = ActiveCell.Value * 12 / mamographyASP1
    ElseIf ActiveCell.End(xlToLeft).Value = mamoSysCode2 Then
    mcSysCode2 = ActiveCell.Value * 12 / mamographyASP2
    ElseIf ActiveCell.End(xlToLeft).Value = mamoSysCode3 Then
    mcSysCode3 = ActiveCell.Value * 12 / mamographyASP3
    ElseIf ActiveCell.End(xlToLeft).Value = mamoSysCode4 Then
    mcSysCode4 = ActiveCell.Value * 12 / mamographyASP4
    ElseIf ActiveCell.End(xlToLeft).Value = mamoSysCode5 Then
    mcSysCode5 = ActiveCell.Value * 12 / mamographyASP5
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

Dim ytdDIvalToPaste As Double
DIValToPaste = Application.WorksheetFunction.Average(mcSysCode1, mcSysCode2, mcSysCode3, mcSysCode4, mcSysCode5)

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = DIValToPaste

'For YTD
Dim YTDDI(12) As String
Dim ytdDIvalToPaste1 As Double

If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Dim YTDRangeDI As Range
Set YTDRangeDI = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(YTDRangeDI)
End If

Case "DXR-PrimaryDiagnost Digital-N"

Workbooks(inputFlDI).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = PDSysCode1 Then
    mcSysCode1 = ActiveCell.Value
    ElseIf ActiveCell.End(xlToLeft).Value = PDSysCode2 Then
    mcSysCode2 = ActiveCell.Value
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

DIValToPaste = (mcSysCode1 + mcSysCode2) * 12 / PDASP

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = DIValToPaste

'For YTD

If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = Application.WorksheetFunction.Average(ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address))
End If


Case "IXR-MOS Endura-Y"
Dim ytdIGTvalToPaste As Double

Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, LookIn:=xlValues, After:=ActiveCell).Select

i = 0
Do Until ActiveCell.Value = ""

If ActiveCell.End(xlToLeft).Value = enduraSysCode1 Then
    mcSysCode1 = ActiveCell.Value
    ElseIf ActiveCell.End(xlToLeft).Value = enduraSysCode2 Then
    mcSysCode2 = ActiveCell.Value
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = (mcSysCode1 + mcSysCode2) * 12 / enduraASP

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
Dim YTDIGT(12) As String
Dim ytdIGTvalToPaste1 As Double
YTDIGT(YTDnum) = ytdIGTvalToPaste
YTDnum = YTDnum + 1

If j = 1 Then

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Dim myrange As Range
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If

Case "IXR-MOS Pulsera-Y"
Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = pulseraSysCode1 Then
    mcSysCode1 = ActiveCell.Value
    ElseIf ActiveCell.End(xlToLeft).Value = pulseraSysCode2 Then
    mcSysCode2 = ActiveCell.Value
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = (mcSysCode1 + mcSysCode2) * 12 / pulseraASP

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If

Case "IXR-MOS Veradius-Y"
Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = veradiusSysCode1 Then
    mcSysCode1 = ActiveCell.Value
    ElseIf ActiveCell.End(xlToLeft).Value = veradiusSysCode2 Then
    mcSysCode2 = ActiveCell.Value
    ElseIf ActiveCell.End(xlToLeft).Value = veradiusSysCode3 Then
    mcSysCode3 = ActiveCell.Value
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = (mcSysCode1 + mcSysCode2 + mcSysCode3) * 12 / veradiusASP

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If

Case "IXR-MOS BV Vectra-N"
Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = vectraSysCode Then
    mcSysCode1 = ActiveCell.Value * 12 * 100 / vectraASP 'Chaned to 40K from 45K as discussed
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = mcSysCode1
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
YTDnum = 1
YTDIGT(YTDnum) = ytdIGTvalToPaste1

ytdIGTvalToPaste1 = ytdIGTvalToPaste + YTDIGT(YTDnum)
YTDnum = YTDnum + 1

If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If

Case "IXR-CV Allura FC-Y"
Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, After:=ActiveCell, LookIn:=xlValues).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = "723003" Then
    mcSysCode1 = ActiveCell.Value * 12 / alluraASP
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = mcSysCode1

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If


Case "DXR-MobileDiagnost Opta-N"
Workbooks(inputFlIGT).Activate
ActiveWorkbook.Sheets("Product Level Data Sheet").Activate
ActiveSheet.UsedRange.Find(What:="Product Level Spend / Unit Per Month - Total", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(What:=warrantyCMonth, LookIn:=xlValues, After:=ActiveCell).Select

i = 0
Do Until ActiveCell.Value = ""

    If ActiveCell.End(xlToLeft).Value = optaSysCode1 Then
    mcSysCode1 = ActiveCell.Value * 12 / optaASP
    End If
    
ActiveCell.Offset(1, 0).Select
i = i + 1
If i = 30 Then
Exit Do
End If
Loop

ytdIGTvalToPaste = mcSysCode1

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(warrantyCDate).Select
ActiveCell.Offset(i - 2, 0).Select
ActiveCell.Value = ytdIGTvalToPaste

'For YTD
If j = 1 Then
Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="1st Year Warranty Cost / ASP", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Set myrange = ActiveSheet.Range(ActiveCell.Offset(0, 1).Address, ActiveCell.End(xlToRight).Address)
ActiveCell.Value = Application.WorksheetFunction.Average(myrange)
End If

End Select

j = j - 1 'loop for each month
Loop

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

Next productItem 'for all product groups

Workbooks(inputFlDI).Close False
Workbooks(inputFlIGT).Close False
End Sub


'Service information Calculations
Public Sub Service_Information()

On Error Resume Next
Dim inputFl As String
Dim outputFl As String
Dim InfoProductGroup As String
Dim cProductGroup As String
Dim InfoDate As String
Dim i As Integer
Dim j As Integer
Dim productItem As Variant
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim fstMonthChk As String

fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")

yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal
inputFl = ThisWorkbook.Path & "\" & "Service_Information_Quality_Completion.xlsx"

If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "Service_Information_Quality_Completion.xlsx") = "" Then
        Exit Sub
    End If
End If

If Sheet1.rdbSharedDrive.Value = True Then
    SharedDrive_Path "Service_Information_Quality_Completion.xlsx"
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> "Service_Information_Quality_Completion.xlsx" Then
            Exit Sub
        End If
    inputFl = sharedDrivePath
End If

Application.Workbooks.Open (inputFl)
inputFl = "Service_Information_Quality_Completion.xlsx"

InfoProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        InfoProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If InfoProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
InfoDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
InfoProductGroup = "Pulsera"

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
InfoProductGroup = "BV Vectra"

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
InfoProductGroup = "Endura"

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
InfoProductGroup = "Veradius"

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
InfoProductGroup = "Allura FC"

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
InfoProductGroup = "Libra"

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
InfoProductGroup = "PrimaryDiagnost Digital"

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
InfoProductGroup = "MicroDose Mammography"

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
InfoProductGroup = "MobileDiagnost Opta"

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

'Do Until j = 0
Select Case j

    Case 1
    InfoDate = "Jan"
    selMonth = "01"
    Case 2
    InfoDate = "Feb"
    selMonth = "02"
    Case 3
    InfoDate = "Mar"
    selMonth = "03"
    Case 4
    InfoDate = "Apr"
    selMonth = "04"
    Case 5
    InfoDate = "May"
    selMonth = "05"
    Case 6
    InfoDate = "Jun"
    selMonth = "06"
    Case 7
    InfoDate = "Jul"
    selMonth = "07"
    Case 8
    InfoDate = "Aug"
    selMonth = "08"
    Case 9
    InfoDate = "Sep"
    selMonth = "09"
    Case 10
    InfoDate = "Oct"
    selMonth = "10"
    Case 11
    InfoDate = "Nov"
    selMonth = "11"
    Case 12
    InfoDate = "Dec"
    selMonth = "12"
End Select

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.UsedRange.Find(What:=InfoProductGroup, lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
Dim fstAdd As String
Dim lstAdd As String
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.Offset(1, 0).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Service Information Quality", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Application.ActiveCell.PasteSpecial xlPasteAll

'j = j - 1 'loop for each month
'Loop

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

ActiveSheet.Cells(11, 11).Select

Next productItem 'for all product groups

Workbooks(inputFl).Close False
Workbooks(outputFl).Save

End Sub



Public Sub CQ_Calculations()

On Error Resume Next
Dim inputFl As String
Dim outputFl As String
Dim cqProductGroup As String
Dim cProductGroup As String
Dim cqDate As String
Dim i As Integer
Dim j As Integer
Dim productItem As Variant
Dim yrSelectedFirst As String 'Month and year selected at first
Dim selMonth As String
Dim KPISheetName As String
Dim selectSheet As Integer
Dim fstMonthChk As String

fstMonthChk = Format(Sheet1.combYear.Value, "mmmyy")

yrSelectedFirst = Sheet1.combYear.Value

outputFl = outputFileGlobal
inputFl = ThisWorkbook.Path & "\" & "CQ_Data_SPM.xlsx"

'Skipping if input file not present
If Sheet1.rdbLocalDrive.Value = True Then
    If Dir(ThisWorkbook.Path & "\" & "CQ_Data_SPM.xlsx") = "" Then
        Application.Workbooks(outputFl).Windows(1).Visible = True
        Exit Sub
    End If
End If

If Sheet1.rdbSharedDrive.Value = True Then
    SharedDrive_Path "CQ_Data_SPM.xlsx"
        'if file is not present in shared drive then exit
        If Split(sharedDrivePath, "\")(UBound(Split(sharedDrivePath, "\"))) <> "CQ_Data_SPM.xlsx" Then
        Application.Workbooks(outputFl).Windows(1).Visible = True
        
            Exit Sub
        End If
    inputFl = sharedDrivePath
End If

Application.Workbooks.Open (inputFl)
inputFl = "CQ_Data_SPM.xlsx"

cqProductGroup = Sheet1.combProductGroup.Value

'loop for all the product groups
For Each productItem In Sheet1.combProductGroup.List
    If Sheet1.chkAllGroups.Value = True Then
        Sheet1.combProductGroup.Value = productItem
        cqProductGroup = Sheet1.combProductGroup.Value
            
            'exit for for end of list
            If cqProductGroup = "" Then
            Exit For
            End If
    End If

selectSheet = 0

i = 0
cqDate = Mid(Sheet1.combYear.Value, 6, 2)
j = Mid(Sheet1.combYear.Value, 6, 2)

'Case select for sheet tab
KPISheetName = Sheet1.combProductGroup.Value

Select Case KPISheetName

Case "IXR-MOS Pulsera-Y"
KPISheetName = "Pulsera"
cqProductGroup = "BV Family"

Case "IXR-MOS BV Vectra-N"
KPISheetName = "BV Vectra"
cqProductGroup = "BV Vectra"

Case "IXR-MOS Endura-Y"
KPISheetName = "Endura"
cqProductGroup = "BV Family"

Case "IXR-MOS Veradius-Y"
KPISheetName = "Veradius"
cqProductGroup = "BV Family"

Case "IXR-CV Allura FC-Y"
KPISheetName = "Allura FC"
cqProductGroup = "Allura FC"

Case "IXR-MOS Libra-N"
KPISheetName = "Libra"
cqProductGroup = "Libra"

Case "DXR-PrimaryDiagnost Digital-N"
KPISheetName = "PrimaryDiagnost Digital"
cqProductGroup = "PrimaryDiagnost Digital"

Case "DXR-MicroDose Mammography-Y"
KPISheetName = "MicroDose Mammography"
cqProductGroup = "MicroDose Mammography"

Case "DXR-MobileDiagnost Opta-N"
KPISheetName = "MobileDiagnost Opta"
cqProductGroup = "MobileDiagnost Opta"

End Select

'checking whether sheet exists in the output file
Dim exists As Boolean
exists = False
Workbooks(outputFl).Activate
For i = 1 To Workbooks(outputFl).Sheets.Count
    If Workbooks(outputFl).Sheets(i).name = KPISheetName Then
        exists = True
    End If
Next i

If Not exists Then
    GoTo sheetNameNotPresent
End If

'Do Until j = 0
Select Case j

    Case 1
    cqDate = "Jan"
    selMonth = "01"
    Case 2
    cqDate = "Feb"
    selMonth = "02"
    Case 3
    cqDate = "Mar"
    selMonth = "03"
    Case 4
    cqDate = "Apr"
    selMonth = "04"
    Case 5
    cqDate = "May"
    selMonth = "05"
    Case 6
    cqDate = "Jun"
    selMonth = "06"
    Case 7
    cqDate = "Jul"
    selMonth = "07"
    Case 8
    cqDate = "Aug"
    selMonth = "08"
    Case 9
    cqDate = "Sep"
    selMonth = "09"
    Case 10
    cqDate = "Oct"
    selMonth = "10"
    Case 11
    cqDate = "Nov"
    selMonth = "11"
    Case 12
    cqDate = "Dec"
    selMonth = "12"
End Select

Workbooks(inputFl).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.UsedRange.Find(What:=cqProductGroup, lookat:=xlWhole).Select

Do Until ActiveCell.Value = "YTD"
ActiveCell.Offset(0, 1).Select
Loop

ActiveCell.Offset(1, 0).Select
Dim fstAdd As String
Dim lstAdd As String
fstAdd = ActiveCell.Address
ActiveCell.Offset(0, j).Select
ActiveCell.Offset(4, 0).Select
lstAdd = ActiveCell.Address

ActiveSheet.Range(fstAdd, lstAdd).Select
Selection.Copy

Workbooks(outputFl).Activate
ActiveWorkbook.Sheets(KPISheetName).Activate
ActiveSheet.UsedRange.Find(What:="Open Service Interest CQ - PR", lookat:=xlWhole).Select
i = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(i - 2, 0).Select
Application.ActiveCell.PasteSpecial xlPasteAll

'j = j - 1 'loop for each month
'Loop

sheetNameNotPresent:
'exit loop if all groups option is not selected
If Sheet1.chkAllGroups.Value = False Then
    Exit For
End If

ActiveSheet.Cells(11, 11).Select

Next productItem 'for all product groups

Workbooks(inputFl).Close False
Application.Workbooks(outputFl).Windows(1).Visible = True
Workbooks(outputFl).Activate
Dim Sheet
For Each Sheet In ActiveWorkbook.Sheets
Sheet.Activate
ActiveSheet.UsedRange.Find(What:="YTD", lookat:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
ActiveCell.End(xlToRight).Select
Dim forNA As Integer
For forNA = 1 To 48
    If ActiveCell.Value = "" Then
        ActiveCell.Value = "NA"
    End If
    ActiveCell.Offset(1, 0).Select
Next
ActiveSheet.Cells.Select
Selection.Font.Size = 18
ActiveSheet.Cells(11, 11).Select
Next

Workbooks(outputFl).Save

If Sheet1.chkAllGroups.Value = True Then
ActiveWorkbook.Sheets(2).Activate
End If

Application.FileDialog(msoFileDialogFolderPicker).Filters.Clear
'loop for month not present in input
For mnthNt = 1 To 10
    If mnthNot(mnthNt) <> "" Then
        MsgBox mnthNot(mnthNt)
        mnthNot(mnthNt) = ""
    End If
Next

'Message box for sheet not present
For i = 1 To 20
    If shtNotPresent(shtNt) <> "" Then
        MsgBox shtNotPresent(1) & vbCrLf & shtNotPresent(2) & vbCrLf & shtNotPresent(3) & vbCrLf & shtNotPresent(4) & vbCrLf & shtNotPresent(5) & vbCrLf & shtNotPresent(6) & vbCrLf & shtNotPresent(7) & vbCrLf & shtNotPresent(8) & vbCrLf & shtNotPresent(9) & vbCrLf & shtNotPresent(10) & vbCrLf & shtNotPresent(11) & "No work is done for above!"
        shtNotPresent(20) = ""
    End If
Next

End Sub

