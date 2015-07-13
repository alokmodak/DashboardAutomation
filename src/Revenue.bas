Attribute VB_Name = "Revenue"
Option Explicit

Public Sub Revenue_Graph_Creation()

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

Dim PvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim pvtFld As PivotField
Dim lastRow
Dim lastColumn
Dim rngDataForPivot As String
Dim pvtItem As PivotItem


'Copy Data from SAP file
inputRevenue = "EPV_2014YTD2015.xlsx"
SharedDrive_Path inputRevenue
Application.Workbooks.Open (sharedDrivePath)
inputFileNameContracts = inputRevenue
outputFile = Left(sharedDrivePath, InStrRev(sharedDrivePath, "\") - 1) & "\" & "ContractDynamics_Waterfall_Jul15.xlsx"
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
Application.Workbooks.Add
ActiveWorkbook.SaveAs fileName:=outputFile, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
outputFile = ActiveWorkbook.name

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
ActiveWorkbook.Sheets(1).Activate
With ActiveSheet.Range("A:A")
    .PasteSpecial xlPasteValues
End With
ActiveSheet.name = "Data"

'Creating PivotTable
Application.Workbooks(inputFileNameContracts).Close False

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
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R1C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion14)

'change style of the new PivotTable:
PvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
PvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
PvtTbl.ManualUpdate = True

Dim pvtTblName As String
pvtTblName = PvtTbl.name
'Add row, column and page fields in a Pivot Table using the AddFields method:
    ActiveWorkbook.Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Company Code")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 1
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
        .Position = 2
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 3
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
        .Position = 4
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
    For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").PivotItems
    
        If pvtItem = "718094" Or pvtItem = "718095" Then
            pvtItem.Visible = True
        Else
            pvtItem.Visible = False
        End If
    Next
'turn on automatic update / calculation in the Pivot Table
PvtTbl.ManualUpdate = False

'Copy Pivot table values to new sheet
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", LookAt:=xlWhole).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.End(xlDown).Select
ActiveCell.End(xlToLeft).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

ActiveWorkbook.Sheets.Add
With ActiveSheet.Cells(2, 36)
    .PasteSpecial xlPasteValues
End With
ActiveSheet.name = "Endura"

ActiveWorkbook.Sheets("Endura").Activate
ActiveSheet.Cells(2, 36).Select
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
            Do Until ActiveCell.Offset(i, 0).value <> ""
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
        Loop Until ActiveCell.Offset(k, 0).value <> ""

    If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "2To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "MoreThan5Years"
            End If
    
        'condition for After warranty
        If ActiveCell.Offset(0, 3).value = "ZCSW" Then
        j = 1
        zcswVal = True
        Do Until ActiveCell.Offset(j, 0) <> ""
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
            ActiveCell.Offset(0, monthCellForTable - 2).value = "AfterWarranty"
        End If
    End If

End If
End If

    If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "2To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).value = "MoreThan5Years"
            End If
    
            If ActiveCell.Offset(0, 3).value = "ZCSW" Then
            j = 1
            zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> ""
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
                ActiveCell.Offset(0, monthCellForTable - 2).value = "AfterWarranty"
            End If
            End If
    End If
End If


If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
  If ActiveCell.Offset(0, monthCellForTable - 3).value = "No" Then
   If duration <= 12 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "0To1Year"
   ElseIf 13 >= duration <= 36 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "2To3Years"
   ElseIf 37 >= duration <= 60 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "3To5Years"
   ElseIf duration >= 61 Then
     ActiveCell.Offset(0, monthCellForTable - 1).value = "MoreThan5Years"
   End If
'condition for After warranty
If ActiveCell.Offset(0, 3).value = "ZCSW" Then
   j = 1
zcswVal = True
Do Until ActiveCell.Offset(j, 0) <> ""
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
                    ActiveCell.Offset(0, monthCellForTable - 1).value = "AfterWarranty"
                End If
        End If

    End If
End If
            If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
                
                If ActiveCell.Offset(0, monthCellForTable - 3).value = "No" Then
                    If duration <= 12 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "0To1Year"
                    ElseIf 13 >= duration <= 36 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "2To3Years"
                    ElseIf 37 >= duration <= 60 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "3To5Years"
                    ElseIf duration >= 61 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).value = "MoreThan5Years"
                    End If
                    
                    'condition for After warranty
                    If ActiveCell.Offset(0, 3).value = "ZCSW" Then
                        j = 1
                        zcswVal = True
                            Do Until ActiveCell.Offset(j, 0) <> ""
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
                                ActiveCell.Offset(0, monthCellForTable - 1).value = "AfterWarranty"
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

'Calculating total numbers for dropped and up

ActiveCell.Offset(3, 3).Select
ActiveCell.value = "Ends"
ActiveCell.Offset(2, 0).value = "0To1Year"
ActiveCell.Offset(3, 0).value = "2To3Years"
ActiveCell.Offset(4, 0).value = "3To5Years"
ActiveCell.Offset(5, 0).value = "MoreThan5Years"
ActiveCell.Offset(6, 0).value = "AfterWarranty"
ActiveCell.Offset(7, 0).value = "EOL"
ActiveCell.Offset(1, 0).value = "Blanks"

Dim fstTotalCel As String
fstTotalCel = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
Dim totalVal As Integer
'loop for counting totals
For i = 1 To 36
        countLstAddress = ActiveCell.Offset(-3, 0).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 2), 0).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "Yes" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.value = totalVal
            ActiveCell.Offset(-1, 0).value = Format(ActiveCell.Offset(-(topCelVal + 3), 0).value, "mmmyy")
            ActiveCell.Offset(0, 3).Select
            ActiveCell.Offset(-1, -1).value = ActiveCell.Offset(-(topCelVal + 3), -1).value
            ActiveCell.Offset(-1, -2).value = ActiveCell.Offset(-(topCelVal + 3), -2).value
Next

ActiveSheet.Range(fstTotalCel).Select
ActiveCell.Offset(2, 0).Select
Dim fstYear As String
fstYear = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-5, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 4), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "0To1Year" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstYear).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-5, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 4), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "0To1Year" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstYear).Select
ActiveCell.Offset(1, 0).Select
Dim fst1To2Year As String
fst1To2Year = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-6, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 5), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "2To3Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst1To2Year).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-6, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 5), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "2To3Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst1To2Year).Select
ActiveCell.Offset(1, 0).Select
Dim fst2To3Year As String
fst2To3Year = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-7, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 6), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "3To5Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst2To3Year).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-7, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 6), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "3To5Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst2To3Year).Select
ActiveCell.Offset(1, 0).Select
Dim fst3To5Year As String
fst3To5Year = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-8, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 7), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "MoreThan5Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst3To5Year).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-8, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 7), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "MoreThan5Years" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fst3To5Year).Select
ActiveCell.Offset(1, 0).Select
Dim fstMoreThan5Year As String
fstMoreThan5Year = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-9, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 8), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "AfterWarranty" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstMoreThan5Year).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-9, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 8), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "AfterWarranty" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstMoreThan5Year).Select
ActiveCell.Offset(1, 0).Select
Dim fstAfterWarranty As String
fstAfterWarranty = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-10, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 9), 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "EOL" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 1).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstAfterWarranty).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(-10, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 9), 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "EOL" Then
             totalVal = totalVal + 1
            End If
        Next
        ActiveCell.Offset(0, 2).value = totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstTotalCel).Select
ActiveCell.Offset(1, 0).Select
Dim fstBlanks As String
fstBlanks = ActiveCell.Address

For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(0, 1).Address
        countFstAddress = ActiveCell.Offset(6, 1).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            totalVal = totalVal + cell.value
        Next
        ActiveCell.Offset(0, 1).value = ActiveCell.Offset(-1, 0).value - totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

ActiveSheet.Range(fstBlanks).Select
For i = 1 To 37
    If i > 1 Then
        countLstAddress = ActiveCell.Offset(6, 2).Address
        countFstAddress = ActiveCell.Offset(1, 2).Address
        totalVal = 0
        For Each cell In Range(countFstAddress, countLstAddress)
            totalVal = totalVal + cell.value
        Next
        ActiveCell.Offset(0, 2).value = ActiveCell.Offset(-1, 3).value - totalVal
    End If
    If i <= 1 Then
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 3).Select
    End If
Next

'Creating chart
Dim lstChartAdd As String
Dim fstChartAdd As String
Dim chartRange As String
ActiveCell.Offset(0, -1).Select
lstChartAdd = ActiveCell.End(xlDown).Address
ActiveSheet.Range(fstTotalCel).Select
ActiveCell.Offset(-1, 0).Select
fstChartAdd = ActiveCell.Address
chartRange = Range(fstChartAdd, lstChartAdd).Address

    Range(fstChartAdd, lstChartAdd).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnStacked
    ActiveChart.SetSourceData Source:=Range("Endura!" & chartRange)
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
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Endura"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Endura"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 6).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 6).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 18
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendLeft)
    With ActiveChart.Parent
         .Height = 325 ' resize
         .Width = 1500  ' resize
         .Top = 10    ' reposition
         .Left = 10   ' reposition
     End With

ActiveSheet.Cells(1, 1).Select
ActiveWorkbook.Sheets("Pivot").delete
ActiveWorkbook.Sheets("Data").delete
End Sub
