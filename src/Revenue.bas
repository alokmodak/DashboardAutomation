Attribute VB_Name = "Revenue"
Option Explicit
Public revenueSelCountry As String

Public Sub Revenue_Graph_Creation()

Dim NCNotPresent(20) As String
Dim ncNt As Integer
Dim inputFileNameContracts As String
Dim outputFile As String
Dim inputRevenue As String
Dim fstAddForPivot As String
Dim lstAddForPivot As String
Dim duration As String
Dim monthsForTable As String
Dim monthCellForTable As Integer
Dim topCelVal As Integer
Dim fstVal As String
Dim lstVal As String
Dim cell As Variant
Dim i As Integer
Dim j As Integer
Dim zcswVal As Boolean
Dim countFstAddress As String
Dim countLstAddress As String

Dim pvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim pvtFld As PivotField
Dim lastRow
Dim lastColumn
Dim rngDataForPivot As String
Dim pvtItem As PivotItem
Dim strtMonth As String

'On Error Resume Next
ncNt = 1
'Copy Data from SAP file
strtMonth = Format(Now() - 31, "mmmyyyy")
inputRevenue = "EPV_SAPBW_" & strtMonth & ".xlsx"
SharedDrive_Path inputRevenue
Application.Workbooks.Open (sharedDrivePath)
inputFileNameContracts = inputRevenue
Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate

'verify selected system code values are present in SAP data
Dim findSysCode As Integer
For findSysCode = 0 To Sheet1.lstBx6NC.ListCount - 1
    If Sheet1.lstBx6NC.Selected(findSysCode) = True Then
        If Not ActiveSheet.UsedRange.Find(what:=Sheet1.lstBx6NC.List(findSysCode), LookAt:=xlWhole) = True Then
            If Sheet1.chkAllGroups.value = True Then
                NCNotPresent(ncNt) = "The System Code " & Sheet1.lstBx6NC.List(findSysCode) & " Not Available in SAP data!"
                ncNt = ncNt + 1
            Else
                MsgBox "The System Code " & Sheet1.lstBx6NC.List(findSysCode) & " Not Available in SAP data!"
                End
            End If
        End If
    End If
Next

outputFile = Left(sharedDrivePath, InStrRev(sharedDrivePath, "\") - 1) & "\" & "ContractDynamics_Waterfall_" & Format(Now, "mmmyy") & ".xlsx"
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
If Dir(outputFile) = "" Then
    Application.Workbooks.Add
    ActiveWorkbook.SaveAs fileName:=outputFile, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outputFile = ActiveWorkbook.name
Else
    Application.Workbooks.Open (outputFile)
    outputFile = ActiveWorkbook.name
End If

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole, after:=ActiveCell).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.End(xlDown).Select
ActiveCell.End(xlToRight).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

'Paste Copied data in new workbook
Workbooks(outputFile).Activate
Sheets.Add
With ActiveSheet.Range("A:A")
    .PasteSpecial xlPasteValues
End With
ActiveSheet.name = "Data"

'Creating PivotTable
'Application.Workbooks(inputFileNameContracts).Close False

'determine the worksheet which contains the source data
Set wsData = Worksheets("Data")

'A Pivot Cache represents the memory cache for a PivotTable report. Each Pivot Table report has one cache only. Create a new PivotTable cache, and then create a new PivotTable report based on the cache.

'determine source data range (dynamic):
'last row in column no. 1:
lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
'last column in row no. 1:
lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column

Set rngData = wsData.Cells(1, 1).Resize(lastRow, lastColumn)
rngDataForPivot = rngData.Address
'for creating a Pivot Cache (version excel 2003), use the PivotCaches.Create Method. When version is not specified, default version of the PivotTable will be xlPivotTableVersion12:

Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!" & rngDataForPivot, Version:=xlPivotTableVersion14)
'create a PivotTable report based on a Pivot Cache, using the PivotCache.CreatePivotTable method. TableDestination is mandatory to specify in this method.

'create PivotTable in a new worksheet:
Sheets.Add
ActiveSheet.name = "Pivot"
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R1C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion14)

'change style of the new PivotTable:
pvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
pvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
pvtTbl.ManualUpdate = True

Dim pvtTblName As String
pvtTblName = pvtTbl.name
'Add row, column and page fields in a Pivot Table using the AddFields method:
    ActiveWorkbook.Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Company Code")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Company Code").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 3
    End With
    Range("A5").Select
    ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Reference Equipment") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With ActiveSheet.PivotTables(pvtTblName)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .PivotItems("#").Visible = False
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 5
    End With
    Range("B6").Select
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract Start Date (Header)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .PivotItems("#").Visible = False
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)")
        .PivotItems("#").Visible = False
    End With
    Range("C7").Select
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type")
        .Orientation = xlRowField
        .Position = 6
    End With
    Range("D5").Select
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type")
        For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type").PivotItems
            If pvtItem.name = "#" Then
            .PivotItems("#").Visible = False
            ElseIf pvtItem.name = "MV" Then
            .PivotItems("MV").Visible = False
            ElseIf pvtItem.name = "ZPO" Then
            .PivotItems("ZPO").Visible = False
            ElseIf pvtItem.name = "ZSO" Then
            .PivotItems("ZSO").Visible = False
            Else
            pvtItem.Visible = True
            End If
        Next
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").ClearAllFilters

Dim filterSelectedValues As Integer
    Dim secondLoop As Integer
    secondLoop = 1
    
    ActiveWorkbook.Sheets("Pivot").Activate
'        For filterSelectedValues = 0 To Sheet1.lstBx6NC.ListCount - 1
'            If Sheet1.lstBx6NC.Selected(filterSelectedValues) Then
'                For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'                    "[C,S] System Code Material (Material no of  R Eq)").PivotItems
'                    If Sheet1.lstBx6NC.List(filterSelectedValues) <> pvtItem Then
'                        If secondLoop < 2 Then
'                            pvtItem.Visible = False
'                        End If
'                    Else
'                        pvtItem.Visible = True
'                    End If
'                Next
'                secondLoop = secondLoop + 1 'secondloop value is added to avoid visible = false for all selected values
'            End If
'        Next
'turn on automatic update / calculation in the Pivot Table
pvtTbl.ManualUpdate = False

'Copy Pivot table values to new sheet
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.End(xlToRight).Select
ActiveCell.End(xlDown).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

ActiveWorkbook.Sheets.Add
With ActiveSheet.Cells(2, 36)
    .PasteSpecial xlPasteValues
End With

ActiveSheet.name = "Contracts-Chart"
ActiveWorkbook.Sheets("Pivot").delete

ActiveWorkbook.Sheets("Contracts-Chart").Activate
ActiveSheet.Cells(2, 38).Select
Dim fstTableAdd As String
fstTableAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select

monthsForTable = DateAdd("m", -24, Date)

ActiveCell.Offset(0, 1).Select
For monthCellForTable = 2 To 37
    ActiveCell.value = monthsForTable
    ActiveCell.NumberFormat = "[$-409]mmm-yy;@"
        If monthCellForTable > 1 Then
            ActiveCell.Offset(0, 3).Select
            ActiveCell.Offset(0, -1).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Joined"
            ActiveCell.Offset(0, -2).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Dropped"
        End If
    monthsForTable = DateAdd("m", 1, monthsForTable)
Next


ActiveSheet.Range(fstTableAdd).Select

ActiveCell.Offset(1, 0).Select
fstAddForPivot = ActiveCell.Address

countFstAddress = ActiveCell.Address 'first cell for total count

Range(Mid(ActiveCell.Address, 2, 2) & Rows.Count).End(xlUp).Select
lstAddForPivot = ActiveCell.Address

countLstAddress = ActiveCell.Address 'Last cell for total count

ActiveSheet.Range(fstAddForPivot).Select

topCelVal = 1

'Loop for each row individually to calculate values
For Each cell In Range(fstAddForPivot, lstAddForPivot)
If ActiveCell.value <> "" Then
            'leave row values blank if start or end date is not available
            If ActiveCell.Offset(0, 1).value = "" Then
                ActiveCell.Offset(0, 1).value = ActiveCell.Offset(-1, 1).value
            End If
            If ActiveCell.Offset(0, 2).value = "" Then
                ActiveCell.Offset(0, 2).value = ActiveCell.Offset(-2, 2).value
            End If
            duration = DateDiff("m", Replace(ActiveCell.Offset(0, 1).value, ".", "/"), Replace(ActiveCell.Offset(0, 2).value, ".", "/"))
            i = 1
            Do Until ActiveCell.Offset(i, 0).value <> "" Or i > 20
            'exit loop for last cell
                If ActiveCell.Offset(i, 3).value = "" Then
                Exit Do
                End If
            If ActiveCell.Offset(i, 1).value = "" Then
                ActiveCell.Offset(i, 1).value = ActiveCell.Offset(-1, 1).value
            End If
            If ActiveCell.Offset(i, 2).value = "" Then
                ActiveCell.Offset(i, 2).value = ActiveCell.Offset(-2, 2).value
            End If
            duration = duration + DateDiff("m", Replace(ActiveCell.Offset(i, 1).value, ".", "/"), Replace(ActiveCell.Offset(i, 2).value, ".", "/"))
            i = i + 1
            Loop
        
            monthCellForTable = 4
            For i = 1 To 36
            
        Dim k As Integer
        k = 0
        Do
        'exit for last cell
        If ActiveCell.Offset(k, 3).value = "" Then
            Exit Do
        End If
                fstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 1).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 1).value, ".", "/", 4)), 1)
                lstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 2).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 2).value, ".", "/", 4)) + 1, 0)
                
                If fstVal <= CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) And CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) <= lstVal Then
                    ActiveCell.Offset(0, monthCellForTable).value = "Yes"
                Else
                    'condition not to overwrite Yes values
                    If ActiveCell.Offset(0, monthCellForTable).value = "" Then
                        ActiveCell.Offset(0, monthCellForTable).value = "No"
                    End If
                End If
        k = k + 1
        Loop Until ActiveCell.Offset(k, 0).value <> "" Or k > 20

    If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "1To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "MoreThan5Years"
            End If
    
        'condition for After warranty
        If ActiveCell.Offset(0, 3).value = "ZCSW" Then
        j = 1
        zcswVal = True
        Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
        'condition for last row exit loop
            If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                If ActiveCell.Offset(1, 3).value = "" Then
                    Exit Do
            End If
            zcswVal = False
        End If
        j = j + 1
        Loop
        If zcswVal = True Then
            ActiveCell.Offset(0, monthCellForTable - 2).value = "Warranty"
        End If
    End If

End If
End If

    If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "1To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "MoreThan5Years"
            End If
    
            If ActiveCell.Offset(0, 3).value = "ZCSW" Then
            j = 1
            zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
            'condition for last row exit loop
                If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                    If ActiveCell.Offset(1, 3).value = "" Then
                        Exit Do
                    End If
                zcswVal = False
                End If
                j = j + 1
            Loop
            If zcswVal = True Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "Warranty"
            End If
            End If
    End If
End If

If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
  If ActiveCell.Offset(0, monthCellForTable - 3).value = "No" Then
   If duration <= 12 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "0To1Year"
   ElseIf 13 >= duration <= 36 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "1To3Years"
   ElseIf 37 >= duration <= 60 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "3To5Years"
   ElseIf duration >= 61 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "MoreThan5Years"
   End If
'condition for After warranty
If ActiveCell.Offset(0, 3).value = "ZCSW" Then
   j = 1
zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                    'condition for last row exit loop
                    If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                        If ActiveCell.Offset(1, 3).value = "" Then
                            Exit Do
                        End If
                        zcswVal = False
                    End If
                j = j + 1
                Loop
                If zcswVal = True Then
                    ActiveCell.Offset(0, monthCellForTable - 1).value = "Warranty"
                End If
        End If

    End If
End If
            If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
                
                If ActiveCell.Offset(0, monthCellForTable - 3).value = "No" Then
                    If duration <= 12 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "0To1Year"
                    ElseIf 13 >= duration <= 36 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "1To3Years"
                    ElseIf 37 >= duration <= 60 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "3To5Years"
                    ElseIf duration >= 61 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "MoreThan5Years"
                    End If
                    
                    'condition for After warranty
                    If ActiveCell.Offset(0, 3).value = "ZCSW" Then
                        j = 1
                        zcswVal = True
                            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                                'condition for last row exit loop
                                If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                                    If ActiveCell.Offset(1, 3).value = "" Then
                                        Exit Do
                                    End If
                                    zcswVal = False
                                End If
                            j = j + 1
                            Loop
                            If zcswVal = True Then
                                ActiveCell.Offset(0, monthCellForTable - 1).value = "Warranty"
                            End If
                    End If
                End If
            End If

            monthCellForTable = monthCellForTable + 3
    Next
End If
topCelVal = topCelVal + 1
ActiveCell.Offset(1, 0).Select
Next

'Filling country code in the table
ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", LookAt:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
Dim rowCount As Integer
Dim lstRowCnt As Long
Dim celAdd As String
celAdd = Mid(ActiveCell.Address, 2, 2)
rowCount = 0
lstRowCnt = ActiveSheet.Range(celAdd & Rows.Count).End(xlUp).Row
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Offset(1, 0).value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.value = ActiveCell.Offset(-1, 0).value
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Next
ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", LookAt:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Offset(1, -1).value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, -1).value = ActiveCell.Offset(-1, -1).value
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Next

Dim fstPivoAdd As String
Dim lstPivoAdd As String
Dim rngData2 As String
Dim rngDataDest As String
Dim pivoWs As Worksheet
Dim tempA As String
Dim tempB As String


rngDataDest = "Pivot!" & "R40:C1"
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole).Select
fstPivoAdd = ActiveCell.Address
tempA = Application.ConvertFormula(Formula:=fstPivoAdd, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlR1C1)
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(lstRowCnt, 0).Select
lstPivoAdd = ActiveCell.Address
tempB = Application.ConvertFormula(Formula:=lstPivoAdd, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlR1C1)
    ActiveSheet.Range(fstPivoAdd, lstPivoAdd).Select
    rngData2 = "Contracts-Chart!" & tempA & ":" & tempB
    Sheets.Add
    ActiveSheet.name = "Pivot"
    Set pivoWs = ActiveSheet
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rngData2, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="Pivot!R30C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion14
        
ActiveSheet.Cells(30, 1).Select
Dim pvtName As String
Dim posVal As Integer
pvtName = ActiveCell.PivotTable.name
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = False
        .RowGrand = False
        .InGridDropZones = True
        .DisplayFieldCaptions = False
        .RowAxisLayout xlTabularRow
    End With
    Range("A7").Select
    With ActiveSheet.PivotTables("PivotTable1")
        .DisplayFieldCaptions = True
        .DisplayContextTooltips = False
        .ShowDrillIndicators = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Company Code")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Company Code").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    Range("B5").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Reference Equipment").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    Range("C8").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 5
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract End Date (Header)").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type")
        .Orientation = xlRowField
        .Position = 6
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Type").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
        
Dim filterValue As String
Dim filterValueJoined As String
Dim filterValueDropped As String
Dim filterPos As Integer
filterPos = 7
monthsForTable = DateAdd("m", -24, Date)
For monthCellForTable = 7 To 42
    filterValue = Format(monthsForTable, "mmm-yy")
    filterValueJoined = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Joined"
    filterValueDropped = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Dropped"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValue)
        .Orientation = xlRowField
        .Position = filterPos
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValue).Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValueDropped)
        .Orientation = xlRowField
        .Position = filterPos + 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValueDropped).Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValueJoined)
        .Orientation = xlRowField
        .Position = filterPos + 2
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields(filterValueJoined).Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    monthsForTable = DateAdd("m", 1, monthsForTable)
    filterPos = filterPos + 3
Next

ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", LookAt:=xlWhole).Select
ActiveSheet.Cells(Rows.Count, 6).End(xlUp).Select
'Calculating total numbers for dropped and up
Dim lstCelNum As String
Dim lstCelNumForChart As String
lstCelNumForChart = ActiveCell.Address
ActiveCell.Offset(10, 0).Select
lstCelNum = ActiveCell.Address
ActiveCell.value = "ZCSS"
ActiveCell.Offset(2, 0).value = "0To1Year"
ActiveCell.Offset(3, 0).value = "1To3Years"
ActiveCell.Offset(4, 0).value = "3To5Years"
ActiveCell.Offset(5, 0).value = "MoreThan5Years"
ActiveCell.Offset(6, 0).value = "Warranty"
ActiveCell.Offset(7, 0).value = "EOL"
ActiveCell.Offset(8, 0).value = "ZCSP"
ActiveCell.Offset(9, 0).value = "ZCSW"
ActiveCell.Offset(1, 0).value = "Blanks"

ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", LookAt:=xlWhole).Select
Dim fstCelCount As String
fstCelCount = ActiveCell.Offset(1, 0).Address
Dim celNumber As Integer
celNumber = Split(ActiveCell.Address, "$")(UBound(Split(ActiveCell.Address, "$")))
ActiveCell.Offset(0, 1).Select
ActiveSheet.Range(Selection, ActiveCell.End(xlToRight).Address).Select
Selection.Copy
ActiveSheet.Range(lstCelNum).Select
ActiveCell.Offset(-1, 1).Select
ActiveCell.PasteSpecial xlPasteAll

Dim lstCelCount As Integer
lstCelCount = ActiveSheet.Range(fstCelCount, lstCelNum).Count
Dim formulaZCSS As String

Dim fstcntF As String, fstcntG As String, fstcntH As String, fstcntI As String
Dim lstcntF As String, lstcntG As String, lstcntH As String, lstcntI As String

ActiveSheet.Range(fstCelCount).Select
fstcntF = ActiveCell.Address
fstcntG = ActiveCell.Offset(0, 1).Address(False, False)
fstcntH = ActiveCell.Offset(0, 2).Address(False, False)
fstcntI = ActiveCell.Offset(0, 3).Address(False, False)

ActiveSheet.Range(lstCelNumForChart).Select
lstcntF = ActiveCell.Address
lstcntG = ActiveCell.Offset(0, 1).Address(False, False)
lstcntH = ActiveCell.Offset(0, 2).Address(False, False)
lstcntI = ActiveCell.Offset(0, 3).Address(False, False)

Dim blanksAdd As String
Dim lstBanksAdd As String
Dim fstAddB As String
Dim lstAddB As String
Dim fstAddB2 As String
Dim lstaddB2 As String
Dim fstAddB3 As String
Dim lstAddB3 As String
Dim fstAddB4 As String
Dim lstAddB4 As String

ActiveSheet.Range(lstCelNum).Select
ActiveSheet.UsedRange.Find(what:="Blanks", LookAt:=xlWhole).Select
ActiveCell.Offset(0, 2).Select
fstAddB = ActiveCell.Offset(1, 0).Address(False, False)
lstAddB = ActiveCell.Offset(6, 0).Address(False, False)
fstAddB2 = ActiveCell.Offset(-1, -1).Address(False, False)
lstaddB2 = ActiveCell.Offset(8, -1).Address(False, False)
fstAddB3 = ActiveCell.Offset(1, 1).Address(False, False)
lstAddB3 = ActiveCell.Offset(6, 1).Address(False, False)
fstAddB4 = ActiveCell.Offset(-1, 2).Address(False, False)
lstAddB4 = ActiveCell.Offset(8, 2).Address(False, False)

ActiveSheet.Range(lstCelNum).Select
ActiveCell.Offset(0, 1).Formula = "=COUNTIFS(" & fstcntF & ":" & lstcntF & "," & Chr(34) & "ZCSS" & Chr(34) & "," & fstcntG & ":" & lstcntG & "," & Chr(34) & "Yes" & Chr(34) & ")"
ActiveCell.Offset(8, 1).Formula = "=COUNTIFS(" & fstcntF & ":" & lstcntF & "," & Chr(34) & "ZCSP" & Chr(34) & "," & fstcntG & ":" & lstcntG & "," & Chr(34) & "Yes" & Chr(34) & ")"
ActiveCell.Offset(9, 1).Formula = "=COUNTIFS(" & fstcntF & ":" & lstcntF & "," & Chr(34) & "ZCSW" & Chr(34) & "," & fstcntG & ":" & lstcntG & "," & Chr(34) & "Yes" & Chr(34) & ")"
ActiveCell.Offset(1, 2).Formula = "=SUM(" & fstAddB2 & ":" & lstaddB2 & ")-SUM(" & fstAddB & ":" & lstAddB & ")"
ActiveCell.Offset(1, 3).Formula = "=SUM(" & fstAddB4 & ":" & lstAddB4 & ")-SUM(" & fstAddB3 & ":" & lstAddB3 & ")"
ActiveCell.Offset(2, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "0To1Year" & Chr(34) & ")"
ActiveCell.Offset(2, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "0To1Year" & Chr(34) & ")"
ActiveCell.Offset(3, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "1To3Years" & Chr(34) & ")"
ActiveCell.Offset(3, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "1To3Years" & Chr(34) & ")"
ActiveCell.Offset(4, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "3To5Years" & Chr(34) & ")"
ActiveCell.Offset(4, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "3To5Years" & Chr(34) & ")"
ActiveCell.Offset(5, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "MoreThan5Years" & Chr(34) & ")"
ActiveCell.Offset(5, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "MoreThan5Years" & Chr(34) & ")"
ActiveCell.Offset(6, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "Warranty" & Chr(34) & ")"
ActiveCell.Offset(6, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "Warranty" & Chr(34) & ")"
ActiveCell.Offset(7, 2).Formula = "=COUNTIF(" & fstcntH & ":" & lstcntH & "," & Chr(34) & "EOL" & Chr(34) & ")"
ActiveCell.Offset(7, 3).Formula = "=COUNTIF(" & fstcntI & ":" & lstcntI & "," & Chr(34) & "EOL" & Chr(34) & ")"

ActiveCell.Offset(0, 1).Select
ActiveSheet.Range(Selection, ActiveCell.Offset(9, 2).Address).Select
Selection.Copy

Dim formulaCopy As Integer
For formulaCopy = 1 To 36
ActiveCell.Offset(0, 3).Select
ActiveCell.PasteSpecial xlPasteFormulas
Next

'Creating chart
Dim lstChartAdd As String
Dim fstChartAdd As String
Dim chartRange As String

ActiveSheet.Range(lstCelNum).Select
ActiveCell.Offset(-1, 1).Select
fstChartAdd = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(11, 0).Select
lstChartAdd = ActiveCell.Address

    Range(fstChartAdd, lstChartAdd).Select
    chartRange = Selection.Address
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnStacked
    ActiveChart.SetSourceData Source:=Range("Pivot!" & chartRange)
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 18
    ActiveChart.ClearToMatchStyle
    Selection.Format.Fill.Visible = msoFalse
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.ChartGroups(1).GapWidth = 0
    
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    
    ActiveChart.SetElement (msoElementDataLabelCenter)
    ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendLeft)
    With ActiveChart.Parent
         .Height = 325 ' resize
         .Width = 900  ' resize
         .Top = 10    ' reposition
         .Left = 160   ' reposition
     End With
    ActiveChart.SeriesCollection(9).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(10).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.6000000238
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Legend.LegendEntries(9).Select
    Selection.delete
    
Dim c As Chart
Dim s As Series
Dim iPoint As Long
Dim nPoint As Long
Dim colorCounter As Integer
colorCounter = 1

Set c = ActiveChart
For i = 3 To 8
Set s = c.SeriesCollection(i)
s.Select
With Selection.Format.Fill
.Visible = msoTrue
If i = 8 Then
.ForeColor.RGB = RGB(220, 100, 100)
ElseIf i = 7 Then
.ForeColor.RGB = RGB(200, 100, 100)
ElseIf i = 6 Then
.ForeColor.RGB = RGB(170, 70, 70)
ElseIf i = 5 Then
.ForeColor.RGB = RGB(150, 50, 50)
ElseIf i = 4 Then
.ForeColor.RGB = RGB(120, 30, 30)
ElseIf i = 3 Then
.ForeColor.RGB = RGB(100, 20, 20)
End If
.BackColor.ObjectThemeColor = msoThemeColorAccent2
.BackColor.TintAndShade = 0
.BackColor.Brightness = 0.4
End With

nPoint = s.Points.Count
For iPoint = 1 To nPoint
If InStr(1, s.XValues(iPoint), "Joined") Then
s.Points(iPoint).Select
With Selection.Format.Fill
.Visible = msoTrue
.ForeColor.ObjectThemeColor = msoThemeColorAccent3
.ForeColor.TintAndShade = 0
If colorCounter = 4 Then
.ForeColor.Brightness = -0.25
ElseIf colorCounter = 3 Then
.ForeColor.Brightness = -0.4
ElseIf colorCounter = 2 Then
.ForeColor.Brightness = -0.6
ElseIf colorCounter = 1 Then
.ForeColor.Brightness = -0.8
End If
.Transparency = 0
.Solid
End With
End If
colorCounter = i
If colorCounter > 4 Then
colorCounter = colorCounter - 1
End If
Next iPoint
Next i

ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.SeriesCollection(9).Select
    ActiveChart.SeriesCollection(9).ApplyDataLabels
    ActiveChart.SeriesCollection(10).Select
    ActiveChart.SeriesCollection(10).ApplyDataLabels

ActiveChart.SeriesCollection(9).Points(37).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(10).Select
    ActiveChart.SeriesCollection(10).Points(37).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(1).Points(37).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
        .Solid
    End With
    
Application.CutCopyMode = False
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "[C,S] System Code Material (Material no of  R Eq)").Slicers.Add ActiveSheet, _
        , "[C,S] System Code Material (Material no of  R Eq)", _
        "[C,S] System Code Material (Material no of  R Eq)", 120, 153.75, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "[C,S] Company Code").Slicers.Add ActiveSheet, , "[C,S] Company Code", _
        "[C,S] Company Code", 157.5, 191.25, 144, 198.75
    ActiveSheet.Shapes.Range(Array("[C,S] Company Code")).Select
    ActiveSheet.Shapes.Range(Array("[C,S] Company Code")).Top = 10
    ActiveSheet.Shapes.Range(Array("[C,S] Company Code")).Left = 10
    ActiveSheet.Shapes.Range(Array( _
        "[C,S] System Code Material (Material no of  R Eq)")).Select
    ActiveSheet.Shapes.Range(Array( _
        "[C,S] System Code Material (Material no of  R Eq)")).Top = 10
    ActiveSheet.Shapes.Range(Array( _
        "[C,S] System Code Material (Material no of  R Eq)")).Left = 30
            
'deleting old chart
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    If ws.name = "Contracts-Data" Or ws.name = "ContractDynamics-WaterFall" Then
        ws.delete
    End If
Next
ActiveWorkbook.Sheets("Contracts-Chart").Activate
ActiveSheet.name = "Contracts-Data"
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.name = "ContractDynamics-WaterFall"
ActiveSheet.Cells(1, 1).Select

Sheet1.lstBx6NC.MultiSelect = fmMultiSelectSingle
Sheet1.lstBx6NC.value = ""
Sheet1.lstBx6NC.MultiSelect = fmMultiSelectMulti
Sheet1.comb6NC2.value = ""
ActiveWorkbook.Sheets("Data").delete
Application.Workbooks(outputFile).Save
Application.Calculation = xlCalculationAutomatic

End Sub

