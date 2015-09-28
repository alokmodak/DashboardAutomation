Attribute VB_Name = "Revenue"
Option Explicit
Public revenueSelCountry As String
Public revenueOutputGlobal As String
Public marketInputFile As String
Public lastmonthVal As String

Public Sub Revenue_Graph_Creation()

Dim NCNotPresent(20) As String
Dim ncNt As Integer
Dim inputFileNameContracts As String
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

On Error Resume Next

Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
If Application.FileDialog(msoFileDialogFilePicker).Show <> -1 Then
MsgBox "No File is Selected!"
End
End If

inputRevenue = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = ActiveWorkbook.name

'Copy Data from SAP file
strtMonth = Format(Now() - 31, "mmmyyyy")
marketInputFile = "Market_Groups_Markets_Country.xlsx"
marketInputFile = Replace(inputRevenue, inputFileNameContracts, marketInputFile)
Application.Workbooks.Open (marketInputFile), False

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "ContractDynamics_Waterfall_" & Format(Now, "mmmyy") & ".xlsm"
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
If Dir(revenueOutputGlobal) = "" Then
    Application.Workbooks.Add
    ActiveWorkbook.SaveAs fileName:=revenueOutputGlobal, FileFormat:=xlOpenXMLWorkbookMacroEnabled, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    revenueOutputGlobal = ActiveWorkbook.name
Else
    Application.Workbooks.Open (revenueOutputGlobal), False
    revenueOutputGlobal = ActiveWorkbook.name
End If

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, After:=ActiveCell).Select

'Putting names in blank cells
Do Until ActiveCell.Offset(1, 0).Value = "" And ActiveCell.Offset(0, 1).Value = ""
    If ActiveCell.Value = "" Then
        ActiveCell.Value = ActiveCell.Offset(0, -1).Value & " " & "A"
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 1).Select
    End If
    
    If ActiveCell.Value = "EUR" Then
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
    End If
Loop

ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, After:=ActiveCell).Select
fstAddForPivot = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

'Paste Copied data in new workbook
Workbooks(revenueOutputGlobal).Activate
Sheets.Add
With ActiveSheet.Range("A:A")
    .PasteSpecial xlPasteValues
End With
ActiveSheet.name = "Data"

'Adding 6NC Names column
marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
Dim marketFSTAdd As String
Dim marketLSTAdd As String

marketFSTAdd = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Country", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteValues
Dim marketRNG As Range
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "System Code (6NC)"

Dim lstPasteRNG As String
Dim fstPasteRNG As String
Dim lookForVal As String
Dim rngStringMarket As String

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

'Adding Market column

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="Country Code", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Country", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "Market"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

'Adding EOL
Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet2").Activate
ActiveSheet.UsedRange.Find(what:="EOL System code", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
ActiveCell.Offset(0, 2).Select
ActiveCell.End(xlDown).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Country", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "EOL Status"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=IF(IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "3" & "," & "FALSE)<" & Chr(61) & "YEAR(TODAY()),)," & Chr(34) & "EOL" & Chr(34) & "," & Chr(34) & "Not EOL" & Chr(34) & ")"
'=IF(VLOOKUP(B31,$AS$30:$AU$64,3,FALSE)<YEAR(TODAY()),"EOL","Not EOL")
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

Calculating_Data_Downloaded_Date

ActiveWorkbook.Sheets("Data").Activate
Set wsData = Worksheets("Data")

'A Pivot Cache represents the memory cache for a PivotTable report. Each Pivot Table report has one cache only. Create a new PivotTable cache, and then create a new PivotTable report based on the cache.
'determine source data range (dynamic):
'last row in column no. 1:
lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
'last column in row no. 1:
lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column

Set rngData = wsData.Cells(1, 1).Resize(lastRow, lastColumn)
rngDataForPivot = rngData.Address
rngData.Select
'for creating a Pivot Cache (version excel 2003), use the PivotCaches.Create Method. When version is not specified, default version of the PivotTable will be xlPivotTableVersion12:

Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!" & rngDataForPivot, Version:=xlPivotTableVersion15)
'create a PivotTable report based on a Pivot Cache, using the PivotCache.CreatePivotTable method. TableDestination is mandatory to specify in this method.

'create PivotTable in a new worksheet:
Sheets.Add
ActiveSheet.name = "Pivot"
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R1C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)

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
        "EOL Status")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "EOL Status").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "System Code (6NC)")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "System Code (6NC)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Market")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Market").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Country A")
        .Orientation = xlRowField
        .Position = 5
    End With

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Country A").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 6
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
        .Position = 7
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 8
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
        .Position = 9
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
'turn on automatic update / calculation in the Pivot Table
pvtTbl.ManualUpdate = False

'Copy Pivot table values to new sheet
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.UsedRange.Find(what:="EOL Status", lookat:=xlWhole).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.SpecialCells(xlCellTypeLastCell).Select
ActiveCell.End(xlUp).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

ActiveWorkbook.Sheets.Add
With ActiveSheet.Cells(2, 35)
    .PasteSpecial xlPasteValues
End With

ActiveSheet.name = "Contracts-Chart"
ActiveWorkbook.Sheets("Pivot").Delete


ActiveWorkbook.Sheets("Contracts-Chart").Activate
'FOr EOL Blank cells
ActiveSheet.Cells(2, 35).Select
Do Until ActiveCell.Offset(1, 8).Value = ""
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value = "" Then
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
    End If
    If ActiveCell.Value = "#NA" Then
        ActiveCell.Value = "Not Available"
    End If
Loop

ActiveSheet.Cells(2, 40).Select
Dim fstTableAdd As String
fstTableAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select

monthsForTable = DateAdd("m", -24, Date)

ActiveCell.Offset(0, 1).Select
For monthCellForTable = 2 To 37
    ActiveCell.Value = monthsForTable
    ActiveCell.NumberFormat = "[$-409]mmm-yy;@"
        If monthCellForTable > 1 Then
            ActiveCell.Offset(0, 3).Select
            ActiveCell.Offset(0, -1).Value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Joined"
            ActiveCell.Offset(0, -2).Value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Dropped"
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
If ActiveCell.Value <> "" Then
            'leave row values blank if start or end date is not available
            If ActiveCell.Offset(0, 1).Value = "" Then
                ActiveCell.Offset(0, 1).Value = ActiveCell.Offset(-1, 1).Value
            End If
            If ActiveCell.Offset(0, 2).Value = "" Then
                ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(-2, 2).Value
            End If
            duration = DateDiff("m", Replace(ActiveCell.Offset(0, 1).Value, ".", "/"), Replace(ActiveCell.Offset(0, 2).Value, ".", "/"))
            i = 1
            Do Until ActiveCell.Offset(i, 0).Value <> "" Or i > 20
            'exit loop for last cell
                If ActiveCell.Offset(i, 3).Value = "" Then
                Exit Do
                End If
            If ActiveCell.Offset(i, 1).Value = "" Then
                ActiveCell.Offset(i, 1).Value = ActiveCell.Offset(-1, 1).Value
            End If
            If ActiveCell.Offset(i, 2).Value = "" Then
                ActiveCell.Offset(i, 2).Value = ActiveCell.Offset(-2, 2).Value
            End If
            duration = duration + DateDiff("m", Replace(ActiveCell.Offset(i, 1).Value, ".", "/"), Replace(ActiveCell.Offset(i, 2).Value, ".", "/"))
            i = i + 1
            Loop
        
            monthCellForTable = 4
            For i = 1 To 36
            
        Dim k As Integer
        k = 0
        Do
        'exit for last cell
        If ActiveCell.Offset(k, 3).Value = "" Then
            Exit Do
        End If
                fstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 1).Value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 1).Value, ".", "/", 4)), 1)
                lstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 2).Value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 2).Value, ".", "/", 4)) + 1, 0)
                
                If fstVal <= CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).Value) And CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).Value) <= lstVal Then
                    ActiveCell.Offset(0, monthCellForTable).Value = "Yes"
                Else
                    'condition not to overwrite Yes values
                    If ActiveCell.Offset(0, monthCellForTable).Value = "" Then
                        ActiveCell.Offset(0, monthCellForTable).Value = "No"
                    End If
                End If
        k = k + 1
        Loop Until ActiveCell.Offset(k, 0).Value <> "" Or k > 20

    If i = 2 And ActiveCell.Offset(0, monthCellForTable).Value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).Value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "1To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "MoreThan5Years"
            End If
            If ActiveCell.Offset(0, -5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "EOL"
            End If
            
        'condition for After warranty
        If ActiveCell.Offset(0, 3).Value = "ZCSW" Then
        j = 1
        zcswVal = True
        Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
        'condition for last row exit loop
            If ActiveCell.Offset(j, 3).Value <> "ZCSW" Then
                If ActiveCell.Offset(1, 3).Value = "" Then
                    Exit Do
            End If
            zcswVal = False
        End If
        j = j + 1
        Loop
        If zcswVal = True Then
            ActiveCell.Offset(0, monthCellForTable - 2).Value = "Warranty"
        End If
    End If

End If
End If

    If i > 2 And ActiveCell.Offset(0, monthCellForTable).Value = "No" Then
        If ActiveCell.Offset(0, monthCellForTable - 3).Value = "Yes" Then
            If duration <= 12 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "0To1Year"
            ElseIf 13 >= duration <= 36 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "1To3Years"
            ElseIf 37 >= duration <= 60 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "3To5Years"
            ElseIf duration >= 61 Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "MoreThan5Years"
            End If
            If ActiveCell.Offset(0, -5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "EOL"
            End If
            
            If ActiveCell.Offset(0, 3).Value = "ZCSW" Then
            j = 1
            zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
            'condition for last row exit loop
                If ActiveCell.Offset(j, 3).Value <> "ZCSW" Then
                    If ActiveCell.Offset(1, 3).Value = "" Then
                        Exit Do
                    End If
                zcswVal = False
                End If
                j = j + 1
            Loop
            If zcswVal = True Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "Warranty"
            End If
            End If
    End If
End If

If i = 2 And ActiveCell.Offset(0, monthCellForTable).Value = "Yes" Then
  If ActiveCell.Offset(0, monthCellForTable - 3).Value = "No" Then
   If duration <= 12 Then
     ActiveCell.Offset(0, monthCellForTable - 1).Value = "0To1Year"
   ElseIf 13 >= duration <= 36 Then
     ActiveCell.Offset(0, monthCellForTable - 1).Value = "1To3Years"
   ElseIf 37 >= duration <= 60 Then
     ActiveCell.Offset(0, monthCellForTable - 1).Value = "3To5Years"
   ElseIf duration >= 61 Then
     ActiveCell.Offset(0, monthCellForTable - 1).Value = "MoreThan5Years"
   End If
   If ActiveCell.Offset(0, -5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 1).Value = "EOL"
            End If
            
'condition for After warranty
If ActiveCell.Offset(0, 3).Value = "ZCSW" Then
   j = 1
zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                    'condition for last row exit loop
                    If ActiveCell.Offset(j, 3).Value <> "ZCSW" Then
                        If ActiveCell.Offset(1, 3).Value = "" Then
                            Exit Do
                        End If
                        zcswVal = False
                    End If
                j = j + 1
                Loop
                If zcswVal = True Then
                    ActiveCell.Offset(0, monthCellForTable - 1).Value = "Warranty"
                End If
        End If

    End If
End If
            If i > 2 And ActiveCell.Offset(0, monthCellForTable).Value = "Yes" Then
                
                If ActiveCell.Offset(0, monthCellForTable - 3).Value = "No" Then
                    If duration <= 12 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).Value = "0To1Year"
                    ElseIf 13 >= duration <= 36 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).Value = "1To3Years"
                    ElseIf 37 >= duration <= 60 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).Value = "3To5Years"
                    ElseIf duration >= 61 Then
                        ActiveCell.Offset(0, monthCellForTable - 1).Value = "MoreThan5Years"
                    End If
                    If ActiveCell.Offset(0, -5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 1).Value = "EOL"
            End If
            
                    'condition for After warranty
                    If ActiveCell.Offset(0, 3).Value = "ZCSW" Then
                        j = 1
                        zcswVal = True
                            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                                'condition for last row exit loop
                                If ActiveCell.Offset(j, 3).Value <> "ZCSW" Then
                                    If ActiveCell.Offset(1, 3).Value = "" Then
                                        Exit Do
                                    End If
                                    zcswVal = False
                                End If
                            j = j + 1
                            Loop
                            If zcswVal = True Then
                                ActiveCell.Offset(0, monthCellForTable - 1).Value = "Warranty"
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
ActiveSheet.UsedRange.Find(what:="Country A", lookat:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
Dim rowCount As Integer
Dim lstRowCnt As Long
Dim celAdd As String
celAdd = Mid(ActiveCell.Offset(0, 4).Address, 2, 2)
rowCount = 0
lstRowCnt = ActiveSheet.Range(celAdd & Rows.Count).End(xlUp).Row
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Offset(1, 0).Value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Next
ActiveSheet.UsedRange.Find(what:="Country A", lookat:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Offset(1, -1).Value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(-1, -1).Value
    Else
        ActiveCell.Offset(1, 0).Select
    End If
    
Next
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Value = "" Then
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
        ActiveCell.Offset(1, 0).Select
    Else
        ActiveCell.Offset(1, 0).Select
    End If
    
Next
ActiveSheet.UsedRange.Find(what:="Market", lookat:=xlWhole).Select
ActiveCell.Offset(1, 0).Select
For rowCount = 0 To lstRowCnt - 4
    If ActiveCell.Offset(1, 0).Value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
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
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
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
        rngData2, Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Pivot!R30C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15
        
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
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Country A")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Country A").Subtotals = Array(False, _
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

ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", lookat:=xlWhole).Select
ActiveSheet.Cells(Rows.Count, 6).End(xlUp).Select
'Calculating total numbers for dropped and up
Dim lstCelNum As String
Dim lstCelNumForChart As String
lstCelNumForChart = ActiveCell.Address
ActiveCell.Offset(10, 0).Select
lstCelNum = ActiveCell.Address
ActiveCell.Value = "ZCSS"
ActiveCell.Offset(2, 0).Value = "0To1Year"
ActiveCell.Offset(3, 0).Value = "1To3Years"
ActiveCell.Offset(4, 0).Value = "3To5Years"
ActiveCell.Offset(5, 0).Value = "MoreThan5Years"
ActiveCell.Offset(6, 0).Value = "Warranty"
ActiveCell.Offset(7, 0).Value = "EOL"
ActiveCell.Offset(8, 0).Value = "ZCSP"
ActiveCell.Offset(9, 0).Value = "ZCSW"
ActiveCell.Offset(1, 0).Value = "Blanks"

ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", lookat:=xlWhole).Select
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
ActiveSheet.UsedRange.Find(what:="Blanks", lookat:=xlWhole).Select
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
ActiveCell.Offset(10, 0).Select
lstChartAdd = ActiveCell.Address

    Range(fstChartAdd, lstChartAdd).Select
    chartRange = Selection.Address
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnStacked
    ActiveChart.SetSourceData Source:=Range("Pivot!" & chartRange)
    ActiveChart.seriesCollection(2).Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 279
    ActiveChart.ClearToMatchStyle
    Selection.Format.Fill.Visible = msoFalse
    ActiveChart.seriesCollection(1).Select
    ActiveChart.ChartGroups(1).GapWidth = 0
    
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.seriesCollection(1).Select
    ActiveChart.seriesCollection(1).ApplyDataLabels
    
    ActiveChart.SetElement (msoElementDataLabelCenter)
    ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendLeft)
    With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 750  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
    ActiveChart.seriesCollection(9).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent5
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
    ActiveChart.seriesCollection(10).Select
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
    Selection.Delete
    
Dim c As Chart
Dim s As Series
Dim iPoint As Long
Dim nPoint As Long
Dim colorCounter As Integer
colorCounter = 1

Set c = ActiveChart
For i = 3 To 8
Set s = c.seriesCollection(i)
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
.BackColor.Brightness = 0.4000006
End With

nPoint = s.Points.Count
For iPoint = 1 To nPoint
If InStr(1, s.XValues(iPoint), "Joined") Then
s.Points(iPoint).Select
With Selection.Format.Fill
.Visible = msoTrue
.ForeColor.ObjectThemeColor = msoThemeColorAccent6
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
    ActiveChart.seriesCollection(9).Select
    ActiveChart.seriesCollection(9).ApplyDataLabels
    ActiveChart.seriesCollection(10).Select
    ActiveChart.seriesCollection(10).ApplyDataLabels
    ActiveChart.ChartTitle.Text = "Contracts Joins & Drops"
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.seriesCollection(1).Select
    ActiveChart.seriesCollection(1).Points(73).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
        .Solid
    End With
    ActiveChart.seriesCollection(10).Select
    ActiveChart.seriesCollection(9).Select
    ActiveChart.seriesCollection(9).Points(73).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    ActiveChart.seriesCollection(10).Select
    ActiveChart.seriesCollection(10).Points(73).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.8000000119
        .Transparency = 0
        .Solid
    End With
                
    ActiveChart.FullSeriesCollection(10).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.7
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.5
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(9).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.6000000238
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    ActiveSheet.ChartObjects(1).name = "ContractsDynamics"
'deleting old chart
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    If ws.name = "Contracts-Data" Or ws.name = "ContractDynamics-WaterFall" Or ws.name = "Data" Then
        ws.Delete
    End If
Next
ActiveWorkbook.Sheets("Contracts-Chart").Activate
ActiveSheet.name = "Contracts-Data"
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.name = "ContractDynamics-WaterFall"
Range("A1:AG29").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
ActiveSheet.Cells(1, 1).Select


Application.CutCopyMode = False
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "[C,S] System Code Material (Material no of  R Eq)").Slicers.Add ActiveSheet, _
        , "[C,S] System Code Material (Material no of  R Eq)", _
        "[C,S] System Code Material (Material no of  R Eq)", 5, 5, 144, 198
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Country A").Slicers.Add ActiveSheet, , "Country", _
        "Country", 210, 150, 144, 198
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC)", _
        "System Code (6NC)", 5, 150, 144, 198
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Market").Slicers.Add ActiveSheet, , "Market", "Market", 210, 5, 144, _
        198

Range(fstChartAdd, lstChartAdd).Select
Dim fstAddForSecondChart As String
Dim lstAddForSecondChart As String
Dim fstSourceDataAdd As String
Dim lstSourceDataAdd As String
Dim seventhAdd As String

ActiveCell.Offset(1, 0).Select
fstAddForSecondChart = ActiveCell.Offset(-1, 0).Address
fstSourceDataAdd = ActiveCell.Offset(-1, 1).Address
seventhAdd = ActiveCell.Offset(9, 7).Address
lstAddForSecondChart = ActiveCell.End(xlDown).Address
ActiveCell.Offset(-1, 1).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(10, 0).Select
lstSourceDataAdd = ActiveCell.Address

'Adding ScrollBars in to the sheet
ActiveSheet.OLEObjects.Add(ClassType:="Forms.ScrollBar.1", Link:=False, _
        DisplayAsIcon:=False, Left:=1100, Top:=100, _
        Width:=27.1739130434783, Height:=196.739130434783).Select
        ActiveSheet.OLEObjects.Add(ClassType:="Forms.ScrollBar.1", Link:=False, _
        DisplayAsIcon:=False, Left:=550, Top:=400, Width:= _
        318.478260869565, Height:=25).Select
        ActiveSheet.OLEObjects.Add(ClassType:="Forms.ScrollBar.1", Link:=False, _
        DisplayAsIcon:=False, Left:=2100, Top:=400, Width:= _
        318.478260869565, Height:=25).Select
        ActiveSheet.OLEObjects.Add(ClassType:="Forms.ScrollBar.1", Link:=False, _
        DisplayAsIcon:=False, Left:=2800, Top:=100, Width:= _
        27, Height:=196).Select
        
ActiveSheet.ChartObjects("Chart 1").Activate
Dim maxZoomAxes As Integer
maxZoomAxes = ActiveChart.Axes(xlValue).MaximumScale
ActiveSheet.Shapes("ScrollBar1").OLEFormat.Object.Object.Max = maxZoomAxes
ActiveSheet.Shapes("ScrollBar2").OLEFormat.Object.Object.Max = 92
ActiveSheet.Shapes("ScrollBar3").OLEFormat.Object.Object.Max = 35


'Writting code into the new created chart book
    Dim myWrkb As Workbook
    Dim strMacro As String
    Dim prrfModule As VBComponent
    Dim inxNum As Integer
    Dim cdName As String
    cdName = ActiveWorkbook.Sheets("ContractDynamics-WaterFall").CodeName
    inxNum = ActiveWorkbook.Sheets("ContractDynamics-WaterFall").Index
    Set myWrkb = ActiveWorkbook
    
    Set prrfModule = myWrkb.VBProject.VBComponents(cdName)    '  ..VBComponents.Add(1)
    strMacro = "Private Sub ScrollBar1_Change()" & vbCrLf & _
                "Dim chartName As String" & vbCrLf & _
                "ActiveSheet.ChartObjects(" & Chr(34) & "ContractsDynamics" & Chr(34) & ").Activate" & vbCrLf & _
                "    ActiveChart.PlotArea.Select" & vbCrLf & _
                "    ActiveChart.Axes(xlValue).MinimumScale = ScrollBar1.Value" & vbCrLf & _
                "End Sub"
    strMacro = strMacro & vbCrLf & vbCrLf & _
                "Private Sub ScrollBar2_Change()" & vbCrLf & _
                "On Error Resume Next" & vbCrLf & _
                "Dim i As Integer" & vbCrLf & _
                "ActiveSheet.ChartObjects(" & Chr(34) & "ContractsDynamics" & Chr(34) & ").Activate" & vbCrLf & _
                "If ScrollBar2.Value = 0 Then" & vbCrLf & _
                "For i = 1 To 108" & vbCrLf & _
                "ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = False" & vbCrLf & _
                "ActiveChart.FullSeriesCollection(i).HasDataLabels = False" & vbCrLf & _
                "Next" & vbCrLf & _
                "Else" & vbCrLf & _
                    "For i = 1 To 108" & vbCrLf & _
                    "    If i = ScrollBar2.Value Or i = ScrollBar2.Value + 1 Or i = ScrollBar2.Value + 2 Or i = ScrollBar2.Value + 3 Or i = ScrollBar2.Value + 4 Or i = ScrollBar2.Value + 5 Or i = ScrollBar2.Value + 6 Then" & vbCrLf & _
                    "        ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = False" & vbCrLf & _
                    "       ActiveChart.FullSeriesCollection(i).ApplyDataLabels" & vbCrLf & _
                    "    Else" & vbCrLf & _
                    "        ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = True" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    "Next" & vbCrLf & _
                "End If" & vbCrLf & _
                "ActiveSheet.ChartObjects(" & Chr(34) & "ContractsDynamics" & Chr(34) & ").Activate" & vbCrLf & _
                "End Sub"

    strMacro = strMacro & vbCrLf & vbCrLf & _
                "Private Sub ScrollBar3_Change()" & vbCrLf & _
                "On Error Resume Next" & vbCrLf & _
                "Dim i As Integer" & vbCrLf & _
                "ActiveSheet.ChartObjects(" & Chr(34) & "JoinsAndDropsAll" & Chr(34) & ").Activate" & vbCrLf & _
                "For i = 1 To 12" & vbCrLf & _
                "ActiveChart.FullSeriesCollection(i).ApplyDataLabels" & vbCrLf & _
                "Next" & vbCrLf & _
                "If ScrollBar3.Value = 0 Then" & vbCrLf & _
                "ActiveChart.FullSeriesCollection(13).IsFiltered = false" & vbCrLf & _
                "For i = 1 To 35" & vbCrLf & _
                "ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = False" & vbCrLf & _
                "ActiveChart.FullSeriesCollection(i).HasDataLabels = False" & vbCrLf & _
                "Next" & vbCrLf & _
                "Else" & vbCrLf & _
                "ActiveChart.FullSeriesCollection(13).IsFiltered = True" & vbCrLf & _
                    "For i = 1 To 35" & vbCrLf & _
                    "    If i = ScrollBar3.Value Or i = ScrollBar3.Value + 1 Or i = ScrollBar3.Value + 2 Or i = ScrollBar3.Value + 3 Then" & vbCrLf & _
                    "        ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = False" & vbCrLf & _
                    "       ActiveChart.FullSeriesCollection(i).ApplyDataLabels" & vbCrLf & _
                    "    Else" & vbCrLf & _
                    "        ActiveChart.ChartGroups(1).FullCategoryCollection(i).IsFiltered = True" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    "Next" & vbCrLf & _
                "End If"

strMacro = strMacro & vbCrLf & _
                "ActiveSheet.ChartObjects(" & Chr(34) & "JoinsAndDropsAll" & Chr(34) & ").Activate" & vbCrLf & _
                "End Sub" & vbCrLf & vbCrLf & _
            "Private Sub ScrollBar4_Change()" & vbCrLf & _
            "ActiveSheet.ChartObjects(" & Chr(34) & "JoinsAndDropsAll" & Chr(34) & ").Activate" & vbCrLf & _
            "ActiveChart.PlotArea.Select" & vbCrLf & _
            "ActiveChart.Axes(xlValue).MinimumScale = -ScrollBar4.Value" & vbCrLf & _
            "ActiveChart.Axes(xlValue).MaximumScale = ScrollBar4.Value" & vbCrLf & _
           "End Sub"

    prrfModule.CodeModule.AddFromString strMacro
        
    Range("D26:F26").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("D26:F26").Select
    ActiveCell.FormulaR1C1 = "Scroll Right To Zoom"
    Range("D26:F26").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("I6").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Scroll Down To zoom"
    Range("I6").Select
        
    'Creating tables for Sceifics of drops and joins
    Range(fstChartAdd).Select
    ActiveCell.Offset(-1, 0).Select
    For i = 1 To 11
        ActiveCell.Offset(i, -1).Value = i
    Next i
    
    Range(fstChartAdd).Select
    Dim fstSceficsAdd As String
    fstSceficsAdd = ActiveCell.Offset(0, 1).Address
    Dim tableRef As String
    tableRef = ActiveCell.Offset(1, -1).Address(False, True)
    Dim tableRefForHlook As String
    tableRefForHlook = ActiveCell.Offset(3, -1).Address(False, True)
    
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    
    For i = 1 To 35
        ActiveCell.Offset(-1, i).Value = i * 3
    Next i
    
    Dim multiplyCellAdd As String
    multiplyCellAdd = ActiveCell.Offset(-1, 1).Address(False, False)
    Dim negativeAdd As String
    negativeAdd = ActiveCell.Offset(10, 1).Address(False, False)
    Dim fstNewTableAdd As String
    fstNewTableAdd = ActiveCell.Offset(0, 1).Address(True, False)
    
    ActiveCell.Offset(0, 1).Select
    
    ActiveCell.Formula = "=OFFSET(" & fstSceficsAdd & ",0," & multiplyCellAdd & ")"
                                    '=OFFSET($G$11166,0,G11178)
    ActiveCell.Offset(1, 0).Formula = "=HLOOKUP(" & fstNewTableAdd & "," & fstChartAdd & ":" & lstChartAdd & "," & tableRef & ",False)"
                                    '=HLOOKUP($G$11187,$F$11174:$DJ$11185,E11175,FALSE)
    ActiveCell.Offset(2, 0).Formula = "=HLOOKUP(" & fstNewTableAdd & "," & fstChartAdd & ":" & lstChartAdd & "," & "10" & ",False)"
    ActiveCell.Offset(3, 0).Formula = "=HLOOKUP(" & fstNewTableAdd & "," & fstChartAdd & ":" & lstChartAdd & "," & "11" & ",False)"
    
    ActiveCell.Offset(4, 0).Formula = "=HLOOKUP(CONCATENATE(LEFT(" & fstNewTableAdd & ",3),RIGHT(" & fstNewTableAdd & ",2)," & Chr(34) & "-Joined" & Chr(34) & ")," & fstChartAdd & ":" & lstChartAdd & "," & tableRefForHlook & ",False)"
                                    '=HLOOKUP(CONCATENATE(LEFT(G$11179,3),RIGHT(G$11179,2),"-Joined"),$G$11166:$DJ$11176,$E11169,FALSE)
    ActiveCell.Offset(10, 0).Formula = "=HLOOKUP(CONCATENATE(LEFT(" & fstNewTableAdd & ",3),RIGHT(" & fstNewTableAdd & ",2)," & Chr(34) & "-Dropped" & Chr(34) & ")," & fstChartAdd & ":" & lstChartAdd & "," & tableRefForHlook & ",False)"
                                    '=HLOOKUP(CONCATENATE(LEFT(G$11179,3),RIGHT(G$11179,2),"-Dropped"),$G$11166:$DJ$11176,$E11169,FALSE)
    ActiveCell.Offset(16, 0).Formula = "=" & negativeAdd & "*-1"
                                    '=G11189*-1
                     
    For i = 1 To 22
        If ActiveCell.Value = "" Then
            ActiveCell.Offset(-1, 0).Copy
            ActiveCell.PasteSpecial xlPasteFormulas
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
                 
    Range(fstNewTableAdd).Select
    Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Copy
    
    Do Until ActiveCell.Offset(-1, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    Range(fstNewTableAdd).Select
ActiveCell.Offset(1, -1).Value = "ZCSS"
ActiveCell.Offset(2, -1).Value = "ZCSP"
ActiveCell.Offset(3, -1).Value = "ZCSW"
ActiveCell.Offset(4, -1).Value = "Join - 0To1Year"
ActiveCell.Offset(5, -1).Value = "Join - 1To3Years"
ActiveCell.Offset(6, -1).Value = "Join - 3To5Years"
ActiveCell.Offset(7, -1).Value = "Join -MoreThan5Years"
ActiveCell.Offset(8, -1).Value = "Join -Warranty"
ActiveCell.Offset(9, -1).Value = "Join -EOL"
ActiveCell.Offset(10, -1).Value = "Drop - 0To1Year"
ActiveCell.Offset(11, -1).Value = "Drop - 1To3Years"
ActiveCell.Offset(12, -1).Value = "Drop - 3To5Years"
ActiveCell.Offset(13, -1).Value = "Drop -MoreThan5Years"
ActiveCell.Offset(14, -1).Value = "Drop -Warranty"
ActiveCell.Offset(15, -1).Value = "Drop -EOL"
ActiveCell.Offset(16, -1).Value = "Drop - 0To1Year"
ActiveCell.Offset(17, -1).Value = "Drop - 1To3Years"
ActiveCell.Offset(18, -1).Value = "Drop - 3To5Years"
ActiveCell.Offset(19, -1).Value = "Drop -MoreThan5Years"
ActiveCell.Offset(20, -1).Value = "Drop -Warranty"
ActiveCell.Offset(21, -1).Value = "Drop -EOL"

ActiveCell.Offset(23, 0).Formula = "=" & ActiveCell.Address(True, False)
ActiveCell.Offset(24, 0).Formula = "=SUM(" & ActiveCell.Offset(4, 0).Address(False, False) & ":" & ActiveCell.Offset(9, 0).Address(False, False) & ")"
                                    '=SUM(G11183:G11188)
ActiveCell.Offset(25, 0).Formula = "=SUM(" & ActiveCell.Offset(10, 0).Address(False, False) & ":" & ActiveCell.Offset(15, 0).Address(False, False) & ")" & "*-1"
                                    '=SUM(G11189:G11194)*-1
ActiveCell.Offset(26, 0).Formula = "=SUM(" & ActiveCell.Offset(1, 0).Address(False, False) & ":" & ActiveCell.Offset(3, 0).Address(False, False) & ")"

ActiveCell.Offset(27, 0).Formula = "=" & ActiveCell.Offset(25, 0).Address(False, False) & "*-1"
ActiveCell.Offset(28, 0).Formula = "=" & ActiveCell.Offset(26, 0).Address(False, False) & "*-1"

ActiveCell.Offset(24, -1).Value = "Join Total"
ActiveCell.Offset(25, -1).Value = "Drop Total"
ActiveCell.Offset(26, -1).Value = "IB Total"
ActiveCell.Offset(27, -1).Value = "Drop Total Label"
ActiveCell.Offset(28, -1).Value = "IB Dummy"

ActiveCell.Offset(23, 0).Select
Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Copy

Do Until ActiveCell.Offset(-2, 1).Value = ""
    ActiveCell.Offset(0, 1).Select
    ActiveCell.PasteSpecial xlPasteFormulas
Loop

ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, -1).Select

Dim fstChartJoinDrop As String
Dim lstChartJoinDrop As String

fstChartJoinDrop = ActiveCell.Address
ActiveCell.Offset(3, 0).Select
ActiveCell.End(xlToRight).Select
lstChartJoinDrop = ActiveCell.Address

ActiveSheet.Shapes.AddChart2(279, xlColumnStacked).Select
    ActiveChart.SetSourceData Source:=Range( _
        "'ContractDynamics-WaterFall'!" & fstChartJoinDrop & ":" & lstChartJoinDrop)
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlLow
    With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 750  ' resize
         .Top = 10    ' reposition
         .Left = 1200   ' reposition
     End With
ActiveChart.ChartArea.Select
    
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).AxisGroup = 2
    ActiveChart.Axes(xlValue, xlSecondary).Select
    ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = -8000
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveSheet.ChartObjects(2).name = "JoinsAndDrops"
    ActiveSheet.Shapes("ScrollBar4").OLEFormat.Object.Object.Max = ActiveChart.Axes(xlValue).MaximumScale
    ActiveChart.ChartTitle.Text = "Contracts Joins & Drops Total"
    
Range(fstNewTableAdd).Select
Dim fstSecondChartAdd As String
Dim lstSecondChartAdd As String

fstSecondChartAdd = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlToRight).Select
ActiveCell.End(xlDown).Select
lstSecondChartAdd = ActiveCell.Address

Range(fstChartJoinDrop).Select
Dim ibTotalAdd As String
Dim fstToAddAddress As String
Dim lstToAddAddress As String

ibTotalAdd = ActiveCell.Offset(3, 0).Address
ActiveCell.Offset(3, 0).Select
fstToAddAddress = ActiveCell.Offset(0, 1).Address
ActiveCell.End(xlToRight).Select
lstToAddAddress = ActiveCell.Address

ActiveSheet.Shapes.AddChart2(279, xlColumnStacked).Select
ActiveChart.SetSourceData Source:=Range( _
        "'ContractDynamics-WaterFall'!" & fstSecondChartAdd & ":" & lstSecondChartAdd)
ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlLow
    With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 750  ' resize
         .Top = 10    ' reposition
         .Left = 2000   ' reposition
     End With
     
ActiveChart.ChartTitle.Text = "Contracts Joins & Drops Saggrigated"
ActiveChart.seriesCollection(1).Delete
ActiveChart.seriesCollection(1).Delete
ActiveChart.seriesCollection(1).Delete
ActiveChart.seriesCollection(7).Delete
ActiveChart.seriesCollection(7).Delete
ActiveChart.seriesCollection(7).Delete
ActiveChart.seriesCollection(7).Delete
ActiveChart.seriesCollection(7).Delete
ActiveChart.seriesCollection(7).Delete

ActiveChart.ChartArea.Select
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(13).name = _
        "='ContractDynamics-WaterFall'!" & ibTotalAdd
    ActiveChart.FullSeriesCollection(13).Values = _
        "='ContractDynamics-WaterFall'!" & fstToAddAddress & ":" & lstToAddAddress
    ActiveChart.FullSeriesCollection(13).Select
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(13).ChartType = xlLine
    ActiveChart.FullSeriesCollection(13).AxisGroup = 2
    ActiveChart.Axes(xlValue, xlSecondary).Select
    ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = -8000
    ActiveChart.FullSeriesCollection(13).Select
    ActiveChart.FullSeriesCollection(13).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).Select
    Cells(2, 2).Value = ActiveChart.Axes(xlValue).MaximumScale
    Cells(3, 2).Value = ActiveChart.Axes(xlValue).MinimumScale
    
ActiveSheet.ChartObjects(3).name = "JoinsAndDropsAll"
Creating_Trend_Drops_Joins 'calling function for trends
ActiveSheet.Cells(1, 1).Select
Application.Workbooks(marketInputFile).Close False
Application.Workbooks(inputRevenue).Close False
Application.Workbooks(revenueOutputGlobal).Save
Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub Market_Revenue_Chart_Generation()

Dim NCNotPresent(20) As String
Dim ncNt As Integer
Dim inputFileNameContracts As String
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
Dim strtMonth As String

On Error Resume Next

Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
If Application.FileDialog(msoFileDialogFilePicker).Show <> -1 Then
MsgBox "No File is Selected!"
End
End If

inputRevenue = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = ActiveWorkbook.name

'Copy Data from SAP file
strtMonth = Format(Now() - 31, "mmmyyyy")
marketInputFile = "Market_Groups_Markets_Country.xlsx"
marketInputFile = Replace(inputRevenue, inputFileNameContracts, marketInputFile)
Application.Workbooks.Open (marketInputFile), False

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "MarketDynamics_Waterfall_" & Format(Now, "mmmyy") & ".xlsx"
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
If Dir(revenueOutputGlobal) = "" Then
    Application.Workbooks.Add
    ActiveWorkbook.SaveAs fileName:=revenueOutputGlobal, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    revenueOutputGlobal = ActiveWorkbook.name
Else
    Application.Workbooks.Open (revenueOutputGlobal), False
    revenueOutputGlobal = ActiveWorkbook.name
End If

'Paste Copied data in new workbook
Workbooks(revenueOutputGlobal).Activate

'delete Data Sheet if present
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    If ws.name = "Market_Dynamics" Or ws.name = "Contract_Dynamics" Or ws.name = "Data" Then
        ws.Delete
    End If
Next

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, After:=ActiveCell).Select

'Putting names in blank cells
Do Until ActiveCell.Offset(1, 0).Value = "" And ActiveCell.Offset(0, 1).Value = ""
    If ActiveCell.Value = "" Then
        ActiveCell.Value = ActiveCell.Offset(0, -1).Value & " " & "A"
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 1).Select
    End If
    
    If ActiveCell.Value = "EUR" Then
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
    End If
Loop

ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, After:=ActiveCell).Select
fstAddForPivot = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

'Paste Copied data in new workbook
Workbooks(revenueOutputGlobal).Activate

Sheets.Add
With ActiveSheet.Range("A:A")
    .PasteSpecial xlPasteValues
End With

ActiveSheet.name = "Data"

Dim pvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim pvtFld As PivotField
Dim lastRow
Dim lastColumn
Dim rngDataForPivot As String
Dim pvtItem As PivotItem

marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="Country Code", lookat:=xlWhole).Select
Dim marketFSTAdd As String
Dim marketLSTAdd As String

marketFSTAdd = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

'Adding Market column

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Country", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Dim marketRNG As Range
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] Company Code", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "Market"

Dim lstPasteRNG As String
Dim fstPasteRNG As String
Dim lookForVal As String
Dim rngStringMarket As String

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Application.Calculation = xlCalculationAutomatic ' setting to automatic calculations
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

'adding Fiscal Year/Period Column
Application.Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.Value = "Fiscal Year/Period"

ActiveCell.Offset(1, 1).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, -1).Select
lookForVal = ActiveCell.Offset(0, 1).Address(False, False)

ActiveCell.Formula = "=MID(" & lookForVal & ", 5, 4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & lookForVal & ", 2, 2)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)

'Adding 6NC Names column
marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Country", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "System Code (6NC)"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Select

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

Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!" & rngDataForPivot, Version:=xlPivotTableVersion15)
'create a PivotTable report based on a Pivot Cache, using the PivotCache.CreatePivotTable method. TableDestination is mandatory to specify in this method.

'create PivotTable in a new worksheet:
Sheets.Add
ActiveSheet.name = "Pivot"

Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R40C1", TableName:="marketPivotTable", DefaultVersion:=xlPivotTableVersion15)

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
    Cells(40, 1).Select
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields("Market")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields("Market")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Fiscal Year/Period")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("marketPivotTable").AddDataField ActiveSheet. _
        PivotTables("marketPivotTable").PivotFields("�� �Total Contract Revenue"), _
        "Count of �� �Total Contract Revenue", xlCount
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Count of �� �Total Contract Revenue")
        .Caption = "Sum of �� �Total Contract Revenue"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Country A")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("marketPivotTable")
        .RowGrand = False
    End With
    
    ActiveSheet.PivotTables("marketPivotTable").PivotFields("Market").AutoSort _
        xlDescending, "Sum of �� �Total Contract Revenue"

pvtTbl.ManualUpdate = False
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    Dim fstPercentAdd As String
    Dim lstPercentAdd As String
    Dim marketCelAdd As String
    Dim lstMarketAdd As String
        
    fstPercentAdd = ActiveCell.Offset(2, 0).Address(False, True)
    marketCelAdd = ActiveCell.Offset(1, 2).Address
    ActiveCell.End(xlDown).Select
    lstPercentAdd = ActiveCell.Address
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Formula = "=" & fstPercentAdd
    Selection.Copy
    Dim periodRowCount As Integer
    For periodRowCount = 1 To Range(fstPercentAdd, lstPercentAdd).Count - 1
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    Next
    
    Dim numCounter As Integer
    numCounter = 1
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.EntireColumn.Insert xlToRight
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.Value = ""
        ActiveCell.Offset(0, -1).Value = numCounter
        ActiveCell.Offset(1, 0).Select
        numCounter = numCounter + 1
    Loop
    
    lstPercentAdd = ActiveCell.Address
    ActiveSheet.Range(fstPercentAdd).Select
    ActiveCell.End(xlToRight).Select
    lstMarketAdd = ActiveCell.Address(False, False)
    
    Dim constAddMarketCelAdd As String
    constAddMarketCelAdd = ActiveSheet.Range(marketCelAdd).Address(False, False)
    ActiveSheet.Range(lstPercentAdd).Select
    ActiveCell.Offset(1, 1).Select
    ActiveCell.Formula = "=" & constAddMarketCelAdd
    ActiveCell.Copy
    Dim marketColCount As Integer
    Do Until ActiveCell.Offset(-2, 1).Value = ""
    ActiveCell.Offset(0, 1).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    Loop

    Dim fstForPercentCal As String
    Dim lstForPercentCal As String
    
    Dim percentCal As Integer
    Dim perCounter As Integer
    Dim perCounterAdd As String
    Dim lstperCounterAdd As String
    Dim lstCelAdd As String
    Dim lstCelForCal As String
    
    ActiveSheet.Range(lstPercentAdd).Select
    ActiveCell.Offset(2, 1).Select
    fstForPercentCal = ActiveCell.Offset(-1, 0).Address(True, False)
    ActiveSheet.Range(marketCelAdd).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstForPercentCal = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 1).Select
    lstperCounterAdd = ActiveCell.Address(False, True)
    ActiveCell.End(xlToRight).Select
    perCounterAdd = ActiveCell.Address(False, True)
    
    ActiveSheet.Range(lstPercentAdd).Select
    ActiveCell.Offset(2, 1).Select
    lstCelAdd = ActiveCell.Address
    ActiveCell.Formula = "=IFERROR(HLOOKUP(" & fstForPercentCal & "," & marketCelAdd & ":" & lstForPercentCal & "," & fstPercentAdd & ",FALSE)/SUM(" & lstperCounterAdd & ":" & perCounterAdd & "),)"
    ActiveCell.Copy
    ActiveCell.SpecialCells(xlCellTypeLastCell).Select
    lstperCounterAdd = ActiveCell.Address
    ActiveSheet.Range(lstCelAdd, lstperCounterAdd).PasteSpecial xlPasteAll
    
    Application.CutCopyMode = False
    Selection.NumberFormat = "0%"
    Application.CutCopyMode = True
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    Dim fstReferenceTable As String
    Dim lstReferenceTable As String
    Dim rngToPaste As String
    
    ActiveCell.Offset(1, 0).Select
    fstReferenceTable = ActiveCell.Address(False, False)
    
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstReferenceTable = ActiveCell.Address(False, False)
    ActiveSheet.Range(fstReferenceTable, lstReferenceTable).Select
    Selection.Copy
    
    Range(fstReferenceTable).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(0, 3).Select
    ActiveCell.PasteSpecial xlPasteAll
    rngToPaste = Selection.Address
    
    ActiveCell.Formula = "=" & fstReferenceTable
    ActiveCell.Copy
    
    Range(rngToPaste).Select
    Selection.PasteSpecial xlPasteFormulas
    
    Range(fstReferenceTable).Select

    ActiveSheet.PivotTables("marketPivotTable").PivotSelect "", xlDataAndLabel, _
        True
    
    Dim rngForChart As String
    Dim fstAddForChartMarket As String
    Dim lstAddForChartMarket As String
    
    Range(rngToPaste).Select
    fstAddForChartMarket = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddForChartMarket = ActiveCell.Offset(-1, 0).Address
    rngForChart = Range(fstAddForChartMarket, lstAddForChartMarket).Address
    Range(rngForChart).Select
    ActiveSheet.Shapes.AddChart2(276, xlAreaStacked).Select
    ActiveChart.SetSourceData Source:=Range("Pivot!" & rngForChart)
     With ActiveChart.Parent
         .Height = 420 ' resize
         .Width = 900  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.ChartStyle = 279
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.Legend.Position = xlLegendPositionRight
    
     Dim fstRNGData As String
     Dim seriesCollection As Integer
     seriesCollection = 1
     Dim fstSeriesDataAdd As String
     Dim lstSeriesDataAdd As String
     
     ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
     ActiveCell.End(xlDown).Select
     ActiveCell.End(xlDown).Select
     ActiveCell.Offset(0, 1).Select
     
     For seriesCollection = 1 To 5
        fstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(-1, 0).Select
        lstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlUp).Select
        ActiveCell.Offset(1, 1).Select
        fstRNGData = "=Pivot!" & fstSeriesDataAdd & ":" & lstSeriesDataAdd
        ActiveSheet.ChartObjects("Chart 1").Activate
        ActiveChart.FullSeriesCollection(seriesCollection).Select
        ActiveChart.SetElement (msoElementDataLabelCallout)
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.ShowCategoryName = False
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        ActiveChart.seriesCollection(seriesCollection).DataLabels.Format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, fstRNGData, 0
        Selection.ShowRange = True
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.Format.Fill.Visible = msoFalse
        Selection.Format.Line.Visible = msoFalse
        Selection.NumberFormat = "#,##0.00"
        Selection.NumberFormat = "#,##0"
        Selection.NumberFormat = "0"
        ActiveChart.ChartTitle.Text = "Market Revenue Dynamics"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Market Revenue Dynamics"
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
        ActiveCell.Offset(0, 1).Select
     Next
    
    Range("B40").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable") _
        , "Market").Slicers.Add ActiveSheet, , "Market 1", "Market", 477, 297, 144, _
        198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable") _
        , "Country A").Slicers.Add ActiveSheet, , "Country 1", "Country", 514.5, 334.5, _
        144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC)", _
        "System Code (6NC)", 525.75, 413.25, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "[C,S] System Code Material (Material no of  R Eq)").Slicers.Add ActiveSheet, _
        , "[C,S] System Code Material (Material no of  R Eq)", _
        "[C,S] System Code Material (Material no of  R Eq)", 563.25, 450.75, 144, _
        198.75
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(5).DataLabels.Select
    ActiveSheet.Shapes.Range(Array( _
        "[C,S] System Code Material (Material no of  R Eq)")).Select
    
    ActiveSheet.Shapes.Range(Array("Market 1")).Select
    With ActiveSheet.Shapes.Range(Array("Market 1"))
        .Top = 10
        .Left = 5
    End With
    With ActiveSheet.Shapes.Range(Array("Country 1"))
        .Top = 210
        .Left = 5
    End With
    With ActiveSheet.Shapes.Range(Array("[C,S] System Code Material (Material no of  R Eq)"))
        .Top = 210
        .Left = 150
    End With
    With ActiveSheet.Shapes.Range(Array("System Code (6NC)"))
        .Top = 10
        .Left = 150
    End With

Range("A1:W30").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
ActiveSheet.Cells(1, 1).Select
ActiveSheet.name = "Market_Dynamics"

Sheets.Add
ActiveSheet.name = "Pivot"

Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R40C50", TableName:="contractsPivotTable", DefaultVersion:=xlPivotTableVersion15)

'change style of the new PivotTable:
pvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
pvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
pvtTbl.ManualUpdate = True

pvtTblName = pvtTbl.name
'Add row, column and page fields in a Pivot Table using the AddFields method:
    ActiveWorkbook.Sheets("Pivot").Select
    Cells(40, 50).Select
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields("[C] Contract Material Line Item")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "Fiscal Year/Period")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("contractsPivotTable").AddDataField ActiveSheet. _
        PivotTables("contractsPivotTable").PivotFields("�� �Total Contract Revenue"), _
        "Count of �� �Total Contract Revenue", xlCount
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "Count of �� �Total Contract Revenue")
        .Caption = "Sum of �� �Total Contract Revenue"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "Country A")
        .Orientation = xlPageField
        .Position = 1
    End With

pvtTbl.ManualUpdate = False

With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "Fiscal Year/Period")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "[C] Contract Material Line Item")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "[C] Contract Material Line Item")
        .PivotItems("#").Visible = False
    End With
    ActiveSheet.PivotTables("contractsPivotTable").PivotFields( _
        "[C] Contract Material Line Item").AutoSort xlDescending, _
        "Sum of �� �Total Contract Revenue"
    
    ActiveSheet.PivotTables("contractsPivotTable").RowGrand = False 'row grand total not visible
    
    ActiveWorkbook.Sheets("Pivot").Activate
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    Dim fstContractAdd As String
    Dim lstContractAdd As String
    
    fstContractAdd = ActiveCell.Offset(1, 0).Address(False, False)
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Formula = "=" & fstContractAdd
    Selection.NumberFormat = "0"
    ActiveCell.Copy
    
    fstContractAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(13, 0).Select
    lstContractAdd = ActiveCell.Address

    Range(fstContractAdd, lstContractAdd).PasteSpecial (xlPasteAll)
    Dim GT As Integer
    GT = 2
    Do Until ActiveCell.Offset(0, 1).Value = ""
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Offset(-1, 0).Value = GT
     GT = GT + 1
    Loop
    ActiveCell.End(xlToLeft).Select
    
    'Calculating Grand Total
    Dim refNumberForSum As String
    refNumberForSum = ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(2, 0).Value = "Grand Total"
    Dim fstGrandTotal As String
    Dim lstGrandTotal As String
    
    fstGrandTotal = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToRight).Select
    lstGrandTotal = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToLeft).Select
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    Dim fstSumOfTableAdd As String
    Dim lstSumOfTableAdd As String
    
    fstSumOfTableAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstSumOfTableAdd = ActiveCell.Address
    
    Range(fstGrandTotal).Select
    ActiveCell.Offset(0, 1).Select
    
    ActiveCell.Formula = "=IFERROR(VLOOKUP(" & fstGrandTotal & "," & fstSumOfTableAdd & ":" & lstSumOfTableAdd & "," & refNumberForSum & ",FALSE),)" ' For grand total for sum of revenue
    ActiveCell.Copy
    Do Until ActiveCell.Offset(-2, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    Range(fstGrandTotal).Select
    'for count table address reference
    Dim countTableReferenceAdd As String
    countTableReferenceAdd = ActiveCell.Offset(4, 0).Address(False, False)
    
    'Calculating Others
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.Value = "Others"
    
    Dim OthersAdd As String
    OthersAdd = ActiveCell.Offset(1, 1).Address(False, False)
    Dim fstOthersAdd As String
    Dim lstOthersAdd As String
    
    fstOthersAdd = ActiveCell.Offset(-10, 1).Address(False, False)
    lstOthersAdd = ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=IFERROR(" & OthersAdd & "-SUM(" & fstOthersAdd & ":" & lstOthersAdd & "),)" 'others for Sum of revenue
    ActiveCell.NumberFormat = "0"
    ActiveCell.Copy
    Do Until ActiveCell.Offset(-1, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteAll
    Loop
    ActiveCell.End(xlToLeft).Select
    ActiveCell.End(xlUp).Select
    
    'Calculating % table
    Dim fstAddForPercentContracts As String
    Dim lstAddForPercentContracts As String
    
    fstAddForPercentContracts = ActiveCell.Address(False, False)
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddForPercentContracts = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveSheet.Range(fstAddForPercentContracts, lstAddForPercentContracts).Select
    Selection.Copy
    
    Range(fstAddForPercentContracts).Select
    Dim fstAddForLoc As String
    Dim lstAddForLoc As String
    
    fstAddForLoc = ActiveCell.Offset(1, 1).Address(False, False)
    ActiveCell.End(xlDown).Select
    lstAddForLoc = ActiveCell.Offset(0, 1).Address(True, False)
    ActiveCell.Offset(3, 0).Select
    ActiveCell.PasteSpecial xlPasteAll
    Dim rngAddForPercentContracts As String
    rngAddForPercentContracts = Selection.Address
    ActiveCell.Formula = "=" & fstAddForPercentContracts
    ActiveCell.Copy
    Do Until ActiveCell.Offset(1, 0).Value = ""
        ActiveCell.Offset(1, 0).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveCell.End(xlUp).Select
    Do Until ActiveCell.Offset(0, 1).Value = ""
     ActiveCell.Offset(0, 1).Select
     ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(1, 1).Select
    ActiveCell.Formula = "=IFERROR(" & fstAddForLoc & "/" & lstAddForLoc & ",)"
    ActiveCell.Copy
    Dim fstAddForPercentCal As String
    fstAddForPercentCal = ActiveCell.Address
    Dim lstAddForPercentCal As String
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddForPercentCal = ActiveCell.Address
    Range(fstAddForPercentCal, lstAddForPercentCal).PasteSpecial xlPasteFormulas
    
    Selection.PasteSpecial xlPasteAll
    Selection.NumberFormat = "0%"
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveSheet.PivotTables("contractsPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Copy

    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    
    ActiveCell.Offset(-3, 3).Select
    ActiveCell.PasteSpecial xlPasteAll
    
    ActiveCell.PivotTable.name = "countContractsPivotTable"
    With ActiveSheet.PivotTables("countContractsPivotTable").PivotFields( _
        "Sum of �� �Total Contract Revenue")
        .Caption = "Count of �� �Total Contract Revenue"
        .Function = xlCount
    End With
    
    ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
    fstContractAdd = ActiveCell.Offset(1, 0).Address(False, False)
    
    ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Formula = "=" & fstContractAdd
    Selection.NumberFormat = "0"
    ActiveCell.Copy
    
    fstContractAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    Dim fstCountOthersAdd As String
    Dim lstCountOthersAdd As String
    
    fstCountOthersAdd = ActiveCell.Address 'For Others calculation table first address
    
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    
    lstCountOthersAdd = ActiveCell.Address 'For Others calculation table Last address
    
    ActiveCell.Offset(13, 0).Select
    lstContractAdd = ActiveCell.Address
    Range(fstContractAdd, lstContractAdd).PasteSpecial (xlPasteAll)
    
    Range(fstContractAdd).Select
    
    Dim countOthersAdd As String
    countOthersAdd = ActiveCell.Offset(1, 0).Address(False, True)
    GT = 2
    Do Until ActiveCell.Offset(0, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Offset(-1, 0).Value = GT
        GT = GT + 1
    Loop
    ActiveCell.End(xlToLeft).Select
    Dim referenceAddForOthersCount As String
    referenceAddForOthersCount = ActiveCell.Offset(-1, 1).Address(True, False)
    ActiveCell.Offset(1, 1).Select
    
    ActiveCell.Formula = "=IFERROR(VLOOKUP(" & countOthersAdd & "," & fstCountOthersAdd & ":" & lstCountOthersAdd & "," & referenceAddForOthersCount & ",FALSE),)" ' for Count table
    ActiveCell.Copy
    
    Dim fstAddToPasteCountTable As String
    Dim lstAddToPasteCountTable As String
    
    fstAddToPasteCountTable = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddToPasteCountTable = ActiveCell.Address
    
    Range(fstAddToPasteCountTable, lstAddToPasteCountTable).PasteSpecial xlPasteFormulas
    Range(fstContractAdd).Select
        
'Calculating Grand Total for Count
    Dim refNumberForCount As String
    refNumberForCount = ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(2, 0).Value = "Grand Total"
    Dim fstGrandTotalCount As String
    Dim lstGrandTotalCount As String
    
    fstGrandTotalCount = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToRight).Select
    lstGrandTotalCount = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToLeft).Select
    
    ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
    Dim fstSumOfTableAddCount As String
    Dim lstSumOfTableAddCount As String
    
    fstSumOfTableAddCount = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstSumOfTableAddCount = ActiveCell.Address
    
    Range(fstGrandTotalCount).Select
    ActiveCell.Offset(0, 1).Select
    
    ActiveCell.Formula = "=IFERROR(VLOOKUP(" & fstGrandTotalCount & "," & fstSumOfTableAddCount & ":" & lstSumOfTableAddCount & "," & refNumberForCount & ",FALSE),)" ' For grand total for Count of revenue
    ActiveCell.Copy
    Do Until ActiveCell.Offset(-2, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    Range(fstGrandTotalCount).Select
    
    'Calculating Others
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.Value = "Others"
    
    Dim OthersAddCount As String
    OthersAddCount = ActiveCell.Offset(1, 1).Address(False, False)
    Dim fstOthersAddcount As String
    Dim lstOthersAddcount As String
    
    fstOthersAddcount = ActiveCell.Offset(-10, 1).Address(False, False)
    lstOthersAddcount = ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=IFERROR(" & OthersAddCount & "-SUM(" & fstOthersAddcount & ":" & lstOthersAddcount & "),)" 'others for Count of revenue
    ActiveCell.NumberFormat = "0"
    ActiveCell.Copy
    Do Until ActiveCell.Offset(-1, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteAll
    Loop
    ActiveCell.End(xlToLeft).Select
    ActiveCell.End(xlUp).Select
            
    ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
    'ActiveCell.End(xlDown).Select
    'ActiveSheet.Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address).Copy
    ActiveSheet.Range(fstContractAdd).Select
    'ActiveCell.End(xlDown).Select
    'ActiveCell.Offset(1, 0).Select
    'ActiveCell.PasteSpecial xlPasteAll
        
    'ActiveCell.End(xlUp).Select
    
    fstAddForPercentContracts = ActiveCell.Address(False, False)
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddForPercentContracts = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveSheet.Range(fstAddForPercentContracts, lstAddForPercentContracts).Select
    Selection.Copy
    
    Range(fstAddForPercentContracts).Select
    
    lstAddForLoc = ActiveCell.Offset(1, 1).Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(20, 0).Select
    ActiveCell.PasteSpecial xlPasteAll
    rngAddForPercentContracts = Selection.Address
    ActiveCell.Formula = "=" & fstAddForPercentContracts
    ActiveCell.Copy
    Do Until ActiveCell.Offset(1, 0).Value = ""
        ActiveCell.Offset(1, 0).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveCell.End(xlUp).Select
    Dim fstRowAdd As String
    fstRowAdd = ActiveCell.Address
    Do Until ActiveCell.Offset(0, 1).Value = ""
     ActiveCell.Offset(0, 1).Select
     ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveSheet.Range(fstRowAdd).Select
    ActiveCell.Offset(1, 1).Select
    ActiveCell.Formula = "=IFERROR(" & "(" & fstAddForLoc & "/" & lstAddForLoc & ")" & "*12" & ",)"
    ActiveCell.Copy
    fstAddForPercentCal = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(-1, 0).Select
    lstAddForPercentCal = ActiveCell.Address
    Range(fstAddForPercentCal, lstAddForPercentCal).PasteSpecial xlPasteFormulas
    
    Selection.PasteSpecial xlPasteAll
    Selection.NumberFormat = "0"
    
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Formula = "=" & countTableReferenceAdd
    ActiveCell.Copy
    Range(ActiveCell.Offset(1, 0).Address, ActiveCell.End(xlDown).Address).PasteSpecial xlPasteFormulas
    
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    Dim fstChartAdd As String
    Dim lstChartAdd As String
    
    fstChartAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(-1, 0).Select
    lstChartAdd = ActiveCell.Address
    Range(fstChartAdd, lstChartAdd).Select
    
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlAreaStacked
    ActiveChart.SetSourceData Source:=Range("Pivot!" & fstChartAdd & ":" & lstChartAdd)
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.ChartStyle = 279
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Contracts Dynamics"
     With ActiveChart.Parent
         .Height = 420 ' resize
         .Width = 900  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
     
     'Adding Labels
     
     ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
     ActiveCell.End(xlDown).Select
     ActiveCell.Offset(19, 1).Select
     
     For seriesCollection = 1 To 5
        fstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToRight).Select
        lstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToLeft).Select
        ActiveCell.Offset(0, 1).Select
        fstRNGData = "=Pivot!" & fstSeriesDataAdd & ":" & lstSeriesDataAdd
        
        ActiveSheet.ChartObjects("Chart 1").Activate
        ActiveChart.FullSeriesCollection(seriesCollection).Select
        ActiveChart.SetElement (msoElementDataLabelCallout)
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.ShowCategoryName = False
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        ActiveChart.seriesCollection(seriesCollection).DataLabels.Format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, fstRNGData, 0
        Selection.ShowRange = True
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.Format.Fill.Visible = msoFalse
        Selection.Format.Line.Visible = msoFalse
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With

        ActiveCell.Offset(1, 0).Select
     Next
     
     'Adding Slicers
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("contractsPivotTable") _
        , "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC) 1", _
        "System Code (6NC)", 10, 5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("contractsPivotTable") _
        , "[C,S] System Code Material (Material no of  R Eq)").Slicers.Add ActiveSheet _
        , , "[C,S] System Code Material (Material no of  R Eq) 1", _
        "[C,S] System Code Material (Material no of  R Eq)", 10, 150, 144, _
        198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("contractsPivotTable") _
        , "Market").Slicers.Add ActiveSheet, , "Market 2", "Market", 210, 5, _
        144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("contractsPivotTable") _
        , "Country A").Slicers.Add ActiveSheet, , "Country 2", "Country", 210, 150 _
        , 144, 198.75
    
    Range("A1:Z60").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1")).Select
    ActiveWorkbook.SlicerCaches("Slicer_System_Code__6NC1").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("countContractsPivotTable"))
    ActiveSheet.Shapes.Range(Array( _
        "[C,S] System Code Material (Material no of  R Eq) 1")).Select
    ActiveWorkbook.SlicerCaches( _
        "Slicer__C_S__System_Code_Material__Material_no_of__R_Eq1").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("countContractsPivotTable"))
    ActiveSheet.Shapes.Range(Array("Market 2")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Market1").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("countContractsPivotTable"))
    ActiveSheet.Shapes.Range(Array("Country 2")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Country1").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("countContractsPivotTable"))
    
    'Creating Line chart for Contracts Dynamics
    Range(fstRowAdd).Select
    Dim fstAddForLineChart As String
    Dim lstAddForLineChart As String
    
    fstAddForLineChart = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(11, 0).Select
    lstAddForLineChart = ActiveCell.Address
    
    Range(fstAddForLineChart, lstAddForLineChart).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(fstAddForLineChart & ":" & lstAddForLineChart), , xlYes). _
        name = "Table2"
    Range("Table2[#All]").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), _
        "Row Labels").Slicers.Add ActiveSheet, , "Row Labels", "Row Labels", 500, _
        30, 144, 198.75
    ActiveSheet.Shapes.Range(Array("Row Labels")).Select
        
        ActiveSheet.Shapes.AddChart.Select
        Dim lineChart As String
        lineChart = ActiveChart.name
    ActiveChart.ChartType = xlLineMarkers
    ActiveChart.SetSourceData Source:=Range("Pivot!" & fstAddForLineChart & ":" & lstAddForLineChart)
    ActiveChart.Axes(xlValue).Select
    ActiveChart.ChartStyle = 279
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Contract Type - Trending"
    
    Dim trendLine As Integer
    For trendLine = 1 To 11
    ActiveChart.FullSeriesCollection(trendLine).Select
    ActiveChart.FullSeriesCollection(trendLine).Trendlines.Add
    ActiveChart.FullSeriesCollection(trendLine).ApplyDataLabels
    Next
    ActiveChart.PlotBy = xlRows
     With ActiveChart.Parent
         .Height = 420 ' resize
         .Width = 900  ' resize
         .Top = 450    ' reposition
         .Left = 300   ' reposition
     End With
     
     'Adding Labels
     
    Range(fstAddForLineChart).Select
    ActiveCell.Offset(1, 0).Select
    
    'filtering Slicer for top four values
    Dim slicerSelected As Integer
    For slicerSelected = 1 To 11
        If slicerSelected < 5 Then
            ActiveWorkbook.SlicerCaches("Slicer_Row_Labels").SlicerItems(ActiveCell.Value).selected = True
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveWorkbook.SlicerCaches("Slicer_Row_Labels").SlicerItems(ActiveCell.Value).selected = False
            ActiveCell.Offset(1, 0).Select
        End If
    Next
     
     ActiveSheet.UsedRange.Find(what:="Count of �� �Total Contract Revenue", lookat:=xlWhole).Select
     ActiveCell.End(xlDown).Select
     ActiveCell.Offset(19, 1).Select
     
     For seriesCollection = 1 To 4
        fstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToRight).Select
        lstSeriesDataAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToLeft).Select
        ActiveCell.Offset(0, 1).Select
        fstRNGData = "=Pivot!" & fstSeriesDataAdd & ":" & lstSeriesDataAdd
        
        ActiveSheet.ChartObjects("Chart 6").Activate
        ActiveChart.FullSeriesCollection(seriesCollection).Select
        ActiveChart.SetElement (msoElementDataLabelCallout)
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.ShowCategoryName = False
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        ActiveChart.seriesCollection(seriesCollection).DataLabels.Format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, fstRNGData, 0
        Selection.ShowRange = True
        ActiveChart.FullSeriesCollection(seriesCollection).DataLabels.Select
        Selection.Format.Fill.Visible = msoFalse
        Selection.Format.Line.Visible = msoFalse
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        Selection.ShowRange = False
        ActiveCell.Offset(1, 0).Select
     Next

    ActiveWorkbook.Sheets("Pivot").Activate
    ActiveSheet.name = "Contract_Dynamics"
    ActiveWindow.Zoom = 80
    
'Creating Chart For Forcast Revenue
For Each ws In ActiveWorkbook.Sheets
    If ws.name = "Revenue_Forcast" Then
        ws.Delete
    End If
Next
    ActiveWorkbook.Sheets("Contract_Dynamics").Activate
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveSheet.PivotTables("contractsPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.name = "Revenue_Forcast"
    Range("B31").Select
    ActiveSheet.Paste
    ActiveCell.PivotTable.name = "forcastPivotTable"
    ActiveSheet.PivotTables("forcastPivotTable").PivotFields("Market").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("forcastPivotTable").PivotFields( _
        "Sum of �� �Total Contract Revenue").Orientation = xlHidden
    With ActiveSheet.PivotTables("forcastPivotTable").PivotFields("Fiscal Year/Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("forcastPivotTable").PivotFields( _
        "[C] Contract Material Line Item").Orientation = xlHidden
    ActiveSheet.PivotTables("forcastPivotTable").AddDataField ActiveSheet. _
        PivotTables("forcastPivotTable").PivotFields("�� �Total Contract Revenue"), _
        "Count of �� �Total Contract Revenue", xlCount
    With ActiveSheet.PivotTables("forcastPivotTable").PivotFields( _
        "Count of �� �Total Contract Revenue")
        .Caption = "Sum of �� �Total Contract Revenue"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("forcastPivotTable").AddDataField ActiveSheet. _
        PivotTables("forcastPivotTable").PivotFields("�� swo cost" & Chr(10) & "settled to" & Chr(10) & "contract") _
        , "Count of �� swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("forcastPivotTable").PivotFields( _
        "Count of �� swo cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of �� swo cost"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("forcastPivotTable").DataPivotField
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select
    Dim fstAdd As String
    Dim lstAdd As String
    Dim fstAddForTable As String
    Dim lstAddForTable As String
    
    fstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlToRight).Select
    
    Dim fstSecondYearAdd As String
    Dim lstSecondYearAdd As String
    
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Offset(0, -1).Address
    lstAddForTable = ActiveCell.Offset(5, -1).Address
    ActiveCell.End(xlToLeft).Select
    
    ActiveCell.Offset(3, 0).Select
    fstAddForTable = ActiveCell.Address
    
    ActiveCell.Formula = "=" & fstAdd
    ActiveCell.Copy
    Range(fstAddForTable, lstAddForTable).Select
    Selection.PasteSpecial xlPasteFormulas
    Range(fstAddForTable).Select
    
    ActiveCell.Offset(3, 0).Select
    fstAddForTable = ActiveCell.Address
    
    ActiveCell.Value = "Cumm. Revenue"
    ActiveCell.Offset(1, 0).Value = "Cumm. Cost"
    ActiveCell.Offset(2, 0).Value = "iGM%"
    
    ActiveCell.Offset(0, 1).Formula = "=" & ActiveCell.Offset(-2, 1).Address(False, False)
    ActiveCell.Offset(1, 1).Formula = "=" & ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.Offset(2, 1).Formula = "=" & "(" & ActiveCell.Offset(0, 1).Address(False, False) & "-" & ActiveCell.Offset(1, 1).Address(False, False) & ")" & "/" & ActiveCell.Offset(0, 1).Address(False, False)
    ActiveCell.Offset(2, 2).Formula = "=" & "(" & ActiveCell.Offset(0, 1).Address(False, False) & "-" & ActiveCell.Offset(1, 1).Address(False, False) & ")" & "/" & ActiveCell.Offset(0, 1).Address(False, False)
    ActiveCell.Offset(2, 1).NumberFormat = "0%"
    ActiveCell.Offset(2, 2).NumberFormat = "0%"
    
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Formula = "=" & ActiveCell.Offset(0, -1).Address(False, False) & "+" & ActiveCell.Offset(-2, 0).Address(False, False)
    ActiveCell.Offset(1, 0).Formula = "=" & ActiveCell.Offset(1, -1).Address(False, False) & "+" & ActiveCell.Offset(-1, 0).Address(False, False)
    
    Range(ActiveCell.Address, ActiveCell.Offset(2, 0).Address).Copy
    Do Until ActiveCell.Offset(-1, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteAll
    Loop
    
ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select
    
    fstAdd = ActiveCell.Address(False, False)
    
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    lstAddForTable = ActiveCell.Offset(5, 0).Address
    ActiveCell.End(xlToLeft).Select
    
    ActiveCell.Offset(3, 0).Select
    fstAddForTable = ActiveCell.Address
    
    ActiveCell.Formula = "=" & fstAdd
    ActiveCell.Copy
    Range(fstAddForTable, lstAddForTable).Select
    Selection.PasteSpecial xlPasteFormulas
    Range(fstAddForTable).Select
    
    ActiveCell.Offset(3, 0).Select
    fstAddForTable = ActiveCell.Address
    
    Dim yearVal As String
    yearVal = Mid(ActiveCell.Offset(-3, 1).Value, 1, 4)
    ActiveCell.Value = "Cumm. Revenue - " & Mid(ActiveCell.Offset(-3, 1).Value, 1, 4)
    ActiveCell.Offset(1, 0).Value = "Cumm. Cost - " & Mid(ActiveCell.Offset(-3, 1).Value, 1, 4)
    ActiveCell.Offset(2, 0).Value = "iGM% - " & Mid(ActiveCell.Offset(-3, 1).Value, 1, 4)
    
    ActiveCell.Offset(0, 1).Formula = "=" & ActiveCell.Offset(-2, 1).Address(False, False)
    ActiveCell.Offset(1, 1).Formula = "=" & ActiveCell.Offset(-1, 1).Address(False, False)
    ActiveCell.Offset(2, 1).Formula = "=" & "(" & ActiveCell.Offset(0, 1).Address(False, False) & "-" & ActiveCell.Offset(1, 1).Address(False, False) & ")" & "/" & ActiveCell.Offset(0, 1).Address(False, False)
    ActiveCell.Offset(2, 2).Formula = "=" & "(" & ActiveCell.Offset(0, 2).Address(False, False) & "-" & ActiveCell.Offset(1, 2).Address(False, False) & ")" & "/" & ActiveCell.Offset(0, 2).Address(False, False)
    ActiveCell.Offset(2, 1).NumberFormat = "0%"
    ActiveCell.Offset(2, 2).NumberFormat = "0%"
    
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Formula = "=" & ActiveCell.Offset(0, -1).Address(False, False) & "+" & ActiveCell.Offset(-2, 0).Address(False, False)
    ActiveCell.Offset(1, 0).Formula = "=" & ActiveCell.Offset(1, -1).Address(False, False) & "+" & ActiveCell.Offset(-1, 0).Address(False, False)
    
    Range(ActiveCell.Address, ActiveCell.Offset(2, 0).Address).Copy
    Do Until ActiveCell.Offset(-1, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteAll
    Loop
    
    ActiveCell.End(xlUp).Select
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(0, 2).Select
    Do Until Mid(ActiveCell.Value, 1, 4) <> Mid(ActiveCell.Offset(0, -1), 1, 4)
    ActiveCell.Offset(0, 1).Select
    Loop
        
    fstSecondYearAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstSecondYearAdd = ActiveCell.Address
    
    Range(fstSecondYearAdd, lstSecondYearAdd).Cut
    ActiveCell.End(xlToLeft).Select
    ActiveCell.End(xlUp).Select
    ActiveCell.Offset(8, 1).Select
    ActiveSheet.Paste
    
    ActiveCell.Offset(0, -1).Value = "Values"
    ActiveCell.Offset(1, -1).Value = "Sum of �� �Total Contract Revenue -" & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Offset(2, -1).Value = "Sum of �� swo cost - " & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Offset(3, -1).Value = "Cumm. Revenue - " & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Offset(4, -1).Value = "Cumm. Cost - " & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Offset(5, -1).Value = "iGM% - " & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    
    Dim nxtYearFormulaAdd As String
    Dim nxtYearFormulaAddCost As String
    nxtYearFormulaAddCost = ActiveCell.Offset(2, 0).Address(False, False)
    nxtYearFormulaAdd = ActiveCell.Offset(1, 0).Address(False, False)
    ActiveCell.Offset(3, 0).Formula = "=" & nxtYearFormulaAdd
    ActiveCell.Offset(4, 0).Formula = "=" & nxtYearFormulaAddCost
    
    ActiveCell.End(xlToRight).Select
    Dim lstAddForChart As String
    Dim fstAddForChart As String
    Dim addCompare As String
    Dim refAddForForcast As String
    refAddForForcast = ActiveCell.Address
    addCompare = ActiveCell.Offset(0, -1).Address(False, False)
    Dim formulaString As String
    ActiveCell.Formula = "=MID(" & addCompare & ",1,4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & addCompare & ",6,2)+1"
    ActiveCell.Copy
    Do Until Mid(ActiveCell.Offset(0, -1).Value, 6, 2) >= 11
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    ActiveCell.End(xlToLeft).Select
    'converting value to date for forecast
    Do Until ActiveCell.Offset(0, 1).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Offset(-1, 0).Value = CDate(ActiveCell.Value & "-" & 1)
    Loop
    
    Dim forecastFstAdd As String
    Dim forecastLstAdd As String
    Dim currentFstAdd As String
    Dim currentLstAdd As String
    Dim forecastAdd As String
    Dim forecastRNGX As String
    Dim forecastRNGY As String
    
    forecastAdd = ActiveCell.Offset(-1, 0).Address(True, False)
    Range(refAddForForcast).Select
    
    currentLstAdd = ActiveCell.Offset(-1, 0).Address(True, False)
    forecastLstAdd = ActiveCell.Offset(1, 0).Address(False, False)
    ActiveCell.End(xlToLeft).Select
    forecastFstAdd = ActiveCell.Offset(1, 1).Address(False, True)
    currentFstAdd = ActiveCell.Offset(-1, 1).Address
    
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(1, 0).Select
    
    If ActiveCell.Value = "" Then
    ActiveCell.Formula = "=FORECAST(" & forecastAdd & "," & forecastFstAdd & ":" & forecastLstAdd & "," & currentFstAdd & ":" & currentLstAdd & ")"
                        '=FORECAST(H$46,$C48:G48,$C$46:G$46)
    ActiveCell.Copy
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    ActiveCell.NumberFormat = "0%"
    
    End If
    
    Dim forecastTenMonths As Integer
    
    lstAddForChart = ActiveCell.Address
    
    ActiveCell.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToLeft).Select
    
    'Adding Chart for Revenue Forcast
    fstAddForChart = ActiveCell.Address
    
    Range(fstAddForChart, lstAddForChart).Select
    Dim rngChart As String
    rngChart = Selection.Address
    
    ActiveSheet.Shapes.AddChart2(322, xlColumnClustered).Select

    ActiveChart.FullSeriesCollection(1).ChartType = xlLine
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).Select
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(3).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(4).ChartType = xlLine
    ActiveChart.FullSeriesCollection(4).Select
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.FullSeriesCollection(5).ChartType = xlLine
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(5).ChartType = xlLineMarkersStacked
    ActiveChart.FullSeriesCollection(5).AxisGroup = 2
    
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(3).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(2).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(2).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(2).Points(3).Select
    ActiveChart.ChartArea.Select
    With ActiveChart.Parent
         .Height = 320 ' resize
         .Width = 500  ' resize
         .Top = 10    ' reposition
         .Left = 150   ' reposition
    End With
    
    Dim fstAddForLastYear As String
    Dim lstAddForLastYear As String
    Dim fstAddForiGMLast As String
    Dim lstAddForiGMLast As String
    Dim seriesAddRevenue As String
    Dim seriesAddiGM As String
    Dim rngRevenueCumm As String
    Dim rngiGMCumm As String
    
    ActiveSheet.UsedRange.Find(what:="Cumm. Revenue - " & yearVal, lookat:=xlWhole).Select
    seriesAddRevenue = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    
    fstAddForLastYear = ActiveCell.Address
    lstAddForLastYear = ActiveCell.End(xlToRight).Address
    
    rngRevenueCumm = Range(fstAddForLastYear, lstAddForLastYear).Address
    
    ActiveSheet.UsedRange.Find(what:="iGM% - " & yearVal, lookat:=xlWhole).Select
    seriesAddiGM = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    
    fstAddForiGMLast = ActiveCell.Address
    lstAddForiGMLast = ActiveCell.End(xlToRight).Address
    
    rngiGMCumm = Range(fstAddForiGMLast, lstAddForiGMLast).Address
    
    Range(rngiGMCumm).Select
    
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(6).name = "=Revenue_Forcast!" & seriesAddRevenue
    ActiveChart.FullSeriesCollection(6).Values = "=Revenue_Forcast!" & fstAddForLastYear & ":" & lstAddForLastYear
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(7).name = "=Revenue_Forcast!" & seriesAddiGM
    ActiveChart.FullSeriesCollection(7).Values = "=Revenue_Forcast!" & fstAddForiGMLast & ":" & lstAddForiGMLast
    ActiveChart.FullSeriesCollection(6).AxisGroup = 1
    ActiveChart.FullSeriesCollection(6).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(5).ChartType = xlLineMarkers
    ActiveChart.FullSeriesCollection(7).ChartType = xlLineMarkers
    ActiveChart.FullSeriesCollection(7).AxisGroup = 2
    ActiveChart.ChartStyle = 279
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Revenue - Forecast"
     With ActiveChart.Parent
         .Height = 320 ' resize
         .Width = 900  ' resize
         .Top = 100    ' reposition
         .Left = 300   ' reposition
     End With
     ActiveChart.FullSeriesCollection(6).Select
    ActiveChart.FullSeriesCollection(6).ApplyDataLabels
    ActiveChart.FullSeriesCollection(6).DataLabels.Select
    Selection.NumberFormat = "#,##0.00"
    Selection.NumberFormat = "#,##0"
    Selection.NumberFormat = "0"
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "#,##0.00"
    Selection.NumberFormat = "#,##0"
    Selection.NumberFormat = "0"
    Selection.Position = xlLabelPositionCenter
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(7).Select
    ActiveChart.SetElement (msoElementDataLabelCallout)
    ActiveChart.FullSeriesCollection(7).DataLabels.Select
    Selection.ShowCategoryName = False
    Selection.Position = xlLabelPositionBelow
    ActiveChart.FullSeriesCollection(5).Select
    ActiveChart.SetElement (msoElementDataLabelCallout)
    ActiveChart.FullSeriesCollection(5).DataLabels.Select
    Selection.Position = xlLabelPositionBelow
    Selection.ShowCategoryName = False
    
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(6).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(7).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
    End With
    ActiveChart.FullSeriesCollection(5).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.1499999762
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.1499999762
    End With
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(5).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.1499999762
    End With
    Range("A1:S31").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    
        ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(6).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    Selection.Format.TextFrame2.TextRange.Font.Size = 12
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Size = 12
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    
    ActiveWorkbook.Sheets("Contract_Dynamics").Activate
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1")).Select
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1", _
        "[C,S] System Code Material (Material no of  R Eq) 1")).Select
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1", _
        "[C,S] System Code Material (Material no of  R Eq) 1", "Market 2")).Select
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1", _
        "[C,S] System Code Material (Material no of  R Eq) 1", "Market 2", "Country 2" _
        )).Select
    Selection.Copy
    Sheets("Revenue_Forcast").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    'calling a function for trends
    Creating_Trends_Market_Dynamics nxtYearFormulaAdd
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.Workbooks(inputFileNameContracts).Close False
    Application.Workbooks(marketInputFile).Close False
    
End Sub

Public Sub Contract_Penetration()

Dim NCNotPresent(20) As String
Dim ncNt As Integer
Dim inputFileNameContracts As String
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
Dim strtMonth As String

On Error Resume Next
Application.DisplayAlerts = False
Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
If Application.FileDialog(msoFileDialogFilePicker).Show <> -1 Then
MsgBox "No File is Selected!"
End
End If

inputRevenue = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = ActiveWorkbook.name

'Copy Data from SAP file
strtMonth = Format(Now() - 31, "mmmyyyy")
marketInputFile = "Market_Groups_Markets_Country.xlsx"
marketInputFile = Replace(inputRevenue, inputFileNameContracts, marketInputFile)
Application.Workbooks.Open (marketInputFile), False

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "Contract_Penetration_" & Format(Now, "mmmyy") & ".xlsx"
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
If Dir(revenueOutputGlobal) = "" Then
    Application.Workbooks.Add
    ActiveWorkbook.SaveAs fileName:=revenueOutputGlobal, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    revenueOutputGlobal = ActiveWorkbook.name
Else
    Application.Workbooks.Open (revenueOutputGlobal), False
    revenueOutputGlobal = ActiveWorkbook.name
End If

'Paste Copied data in new workbook
Workbooks(revenueOutputGlobal).Activate

'delete Data Sheet if present
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    If ws.name = "CP Trend" Or ws.name = "CP Select Vs Global" Or ws.name = "Data" Then
        ws.Delete
    End If
Next

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find(what:="Profit Center", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="Profit Center", lookat:=xlWhole, After:=ActiveCell).Select

'Putting names in blank cells
Do Until ActiveCell.Offset(1, 0).Value = "" And ActiveCell.Offset(0, 1).Value = ""
    If ActiveCell.Value = "" Then
        ActiveCell.Value = ActiveCell.Offset(0, -1).Value & " " & "A"
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 1).Select
    End If
    
    If ActiveCell.Offset(-1, 0).Value = ActiveCell.Value Then
        ActiveCell.Value = Replace(ActiveCell.Offset(-1, 0).Value, " ", "_")
    End If
Loop

ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find(what:="Profit Center", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="Profit Center", lookat:=xlWhole, After:=ActiveCell).Select

fstAddForPivot = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

'Paste Copied data in new workbook
Workbooks(revenueOutputGlobal).Activate

Sheets.Add
With ActiveSheet.Range("A:A")
    .PasteSpecial xlPasteValues
End With

ActiveSheet.name = "Data"

Dim pvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim pvtFld As PivotField
Dim lastRow
Dim lastColumn
Dim rngDataForPivot As String
Dim pvtItem As PivotItem

marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="Country Code", lookat:=xlWhole).Select
Dim marketFSTAdd As String
Dim marketLSTAdd As String

marketFSTAdd = ActiveCell.Address
Selection.SpecialCells(xlCellTypeLastCell).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

'Adding Market column
Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Material", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Dim marketRNG As Range
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="Company code", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="Company code", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "Market"

Dim lstPasteRNG As String
Dim fstPasteRNG As String
Dim lookForVal As String
Dim rngStringMarket As String

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)

Range(rngStringMarket).Delete

'adding Fiscal Year/Period Column
Application.Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Calendar Year/Month", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="Calendar Year/Month", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.Value = "Calender Year/Period"

ActiveCell.Offset(1, 1).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, -1).Select
lookForVal = ActiveCell.Offset(0, 1).Address(False, False)

ActiveCell.Formula = "=MID(" & lookForVal & ", 4, 4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & lookForVal & ", 1, 2)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)


'Adding 6NC Groups
marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Material", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="Material", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="Material", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "Groups (6NC)"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
Range(rngStringMarket).Select
Selection.Delete

'Adding Countries
marketInputFile = "Market_Groups_Markets_Country.xlsx"

Application.Workbooks(marketInputFile).Activate
ActiveWorkbook.Sheets("Sheet1").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="Country Name", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
marketLSTAdd = ActiveCell.Address
ActiveSheet.Range(marketFSTAdd, marketLSTAdd).Select
Selection.Copy

Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="Material", lookat:=xlWhole).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)
marketRNG.Sort Columns, Orientation:=xlLeftToRight

ActiveSheet.UsedRange.Find(what:="Company code", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="Company code", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Value = "Country"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
Range(rngStringMarket).Select
Selection.Delete

ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Select

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

Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!" & rngDataForPivot, Version:=xlPivotTableVersion15)
'create a PivotTable report based on a Pivot Cache, using the PivotCache.CreatePivotTable method. TableDestination is mandatory to specify in this method.

'create PivotTable in a new worksheet:
Sheets.Add
ActiveSheet.name = "Pivot"
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R40C1", TableName:="cpPivotTable", DefaultVersion:=xlPivotTableVersion15)

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
    Cells(40, 1).Select
    With ActiveSheet.PivotTables("cpPivotTable").PivotFields("Market")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("cpPivotTable").PivotFields("Market")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("cpPivotTable").PivotFields("Material")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("cpPivotTable").CalculatedFields.Add "CP%", _
        "='Under_" & Chr(10) & "Contract' /('IB_Units_" & Chr(10) & "Total' -'Under_" & Chr(10) & "Installation' -'Under_" & Chr(10) & "Warranty' )" _
        , True
    ActiveSheet.PivotTables("cpPivotTable").PivotFields("CP%").Orientation = _
        xlDataField
    Range("C42").Select
    With ActiveSheet.PivotTables("cpPivotTable").PivotFields("Sum of CP%")
        .NumberFormat = "0%"
    End With
    With ActiveSheet.PivotTables("cpPivotTable").PivotFields( _
        "Calender Year/Period")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("cpPivotTable").ColumnGrand = False
    'ActiveSheet.PivotTables("cpPivotTable").PivotFields("Market").AutoSort _
        xlDescending, "Sum of �� �Total Contract Revenue"

pvtTbl.ManualUpdate = False
        
        ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
        ActiveCell.Offset(1, 0).Select
        Dim fstCPTableAdd As String
        Dim lstCPTableAdd As String
        
        fstCPTableAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToRight).Select
        ActiveCell.End(xlDown).Select
        lstCPTableAdd = ActiveCell.Address
        ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
        Range(fstCPTableAdd, lstCPTableAdd).Copy
        
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(3, 0).Select
        ActiveCell.PasteSpecial xlPasteAll
        
        ActiveCell.Formula = "=" & fstCPTableAdd
        ActiveCell.Copy
        Range(ActiveCell.Address, ActiveCell.SpecialCells(xlCellTypeLastCell).Address).Select
        Selection.PasteSpecial xlPasteFormulas
        
        Dim rngForChart As String
        rngForChart = Selection.Address
        
        ActiveCell.End(xlToRight).Select
        Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Copy
        ActiveCell.PasteSpecial xlPasteValues
        ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
        ActiveSheet.PivotTables("cpPivotTable").RowGrand = False
        
        Range(rngForChart).Select
        ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Pivot!" & rngForChart)
    ActiveChart.ChartStyle = 279
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Contract Pentration - Market Trend"
    ActiveChart.FullSeriesCollection(17).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Application.CutCopyMode = False
     With ActiveChart.Parent
         .Height = 220 ' resize
         .Width = 530  ' resize
         .Top = 10    ' reposition
         .Left = 285   ' reposition
     End With
    Range("A38").Select
    
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpPivotTable"), _
        "Groups (6NC)").Slicers.Add ActiveSheet, , "Groups (6NC)", "Groups (6NC)", _
        10, 10, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpPivotTable"), _
        "Material").Slicers.Add ActiveSheet, , "Material", "Material", 210, 150, _
        144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpPivotTable"), _
        "Market").Slicers.Add ActiveSheet, , "Market", "Market", 10, 150, 144, _
        198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpPivotTable"), _
        "Country").Slicers.Add ActiveSheet, , "Country", "Country", 210, 10, 144 _
        , 198.75
    ActiveSheet.Shapes.Range(Array("Country")).Select
    
    ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
    ActiveSheet.PivotTables("cpPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(3, 0).Select
    
    ActiveCell.PasteSpecial xlPasteAll
    Dim pvtNameMaterial As String
    pvtNameMaterial = ActiveCell.PivotTable.name
    With ActiveSheet.PivotTables(pvtNameMaterial).PivotFields("Market")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pvtNameMaterial).PivotFields("Groups (6NC)")
        .Orientation = xlColumnField
        .Position = 1
    End With
        
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        fstCPTableAdd = ActiveCell.Address(False, False)
        ActiveCell.End(xlToRight).Select
        ActiveCell.End(xlDown).Select
        lstCPTableAdd = ActiveCell.Address
        ActiveCell.End(xlToLeft).Select
        Range(fstCPTableAdd, lstCPTableAdd).Copy
        
        ActiveCell.Offset(3, 0).Select
        ActiveCell.PasteSpecial xlPasteAll
        
        ActiveCell.Formula = "=IFERROR(" & fstCPTableAdd & ",)"
        ActiveCell.Copy
        Dim fstTrendAdd As String
        Dim lstTrendAdd As String
        
        fstTrendAdd = ActiveCell.Address
        ActiveCell.End(xlToRight).Select
        ActiveCell.End(xlDown).Select
        lstTrendAdd = ActiveCell.Address
        Range(fstTrendAdd, lstTrendAdd).Select
        Selection.PasteSpecial xlPasteFormulas
        Dim rngMaterialTrend As String
        rngMaterialTrend = Selection.Address
                
    ActiveSheet.Shapes.AddChart2(279, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Pivot!" & rngMaterialTrend)
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Contract Pentration - Product Trend"
    ActiveChart.PlotBy = xlColumns
     With ActiveChart.Parent
         .Height = 220 ' resize
         .Width = 550  ' resize
         .Top = 250    ' reposition
         .Left = 300   ' reposition
     End With
        Range("A1:V33").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Making charts non movable
    Dim objChart As ChartObject
    
    For Each objChart In ActiveSheet.ChartObjects
        objChart.Select
        Selection.Placement = xlFreeFloating
    Next
    
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.name = "CP Trend"
    
'Creating chart for Select Vs Global
    Sheets.Add
    ActiveSheet.name = "Pivot"
    
'Create a new pivot table
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R40C1", TableName:="cpGlobalPivotTable", DefaultVersion:=xlPivotTableVersion15)

'change style of the new PivotTable:
pvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
pvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
pvtTbl.ManualUpdate = True
    
    With ActiveSheet.PivotTables("cpGlobalPivotTable").PivotFields( _
        "Calender Year/Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("cpGlobalPivotTable").AddDataField ActiveSheet. _
        PivotTables("cpGlobalPivotTable").PivotFields("CP%"), "Sum of CP%", xlSum
    With ActiveSheet.PivotTables("cpGlobalPivotTable").PivotFields("Material")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("cpGlobalPivotTable")
        .ColumnGrand = False
        .RowGrand = False
    End With
    With ActiveSheet.PivotTables("cpGlobalPivotTable").PivotFields("Sum of CP%")
        .NumberFormat = "0%"
    End With

pvtTbl.ManualUpdate = False

    ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
    Dim fstTableAdd As String
    Dim lstTableAdd As String
    
    fstTableAdd = ActiveCell.Offset(1, 0).Address(False, False)
    
    ActiveCell.Offset(1, 1).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstTableAdd = ActiveCell.Address
    ActiveCell.End(xlToLeft).Select
    Range(fstTableAdd, lstTableAdd).Copy
    ActiveCell.Offset(3, 0).Select
    ActiveCell.PasteSpecial xlPasteAll
    ActiveCell.Formula = "=" & fstTableAdd
    ActiveCell.Copy
    
    Dim fstFormulaAdd As String
    Dim lstFormulaAdd As String
    Dim fstAddToCutForGlobal As String
    
    fstAddToCutForGlobal = ActiveCell.Address ' store address to get back for pasting below column
    fstFormulaAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstFormulaAdd = ActiveCell.Address
    
    Range(fstFormulaAdd, lstFormulaAdd).Select
    Selection.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(1, 0).Value = "CP%" & "-" & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Value = ""
    
    ActiveCell.Offset(0, 1).Select
    Do Until Mid(ActiveCell.Offset(0, 1).Value, 1, 4) <> Mid(ActiveCell.Value, 1, 4)
        ActiveCell.Offset(0, 1).Select
    Loop
    
    ActiveCell.Offset(0, 1).Select
    Dim fstAddToCut As String
    Dim lstAddToCut As String
    
    fstAddToCut = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddToCut = ActiveCell.Address
    
    Range(fstAddToCut, lstAddToCut).Cut
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(4, 1).Select
    
    ActiveSheet.Paste
    ActiveCell.Offset(1, -1).Value = "CP%" & "-" & Mid(ActiveCell.Value, 1, 4)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlToRight).Select
    
    Dim addCompare As String
    addCompare = ActiveCell.Address(False, False)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=MID(" & addCompare & ",1,4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & addCompare & ",6,2)+1"
    ActiveCell.Copy
    Do Until Mid(ActiveCell.Offset(0, -1).Value, 6, 2) >= 11
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveCell.End(xlToLeft).Select
    'converting value to date for forecast
    Do Until ActiveCell.Value = ""
        
        ActiveCell.Offset(-1, 0).Value = CDate(ActiveCell.Value & "-" & 1)
        ActiveCell.Offset(0, 1).Select
    Loop
    
    Dim forecastAddSelect As String
    Dim fstysAdd As String
    Dim lstysAdd As String
    Dim fstxsAdd As String
    Dim lstxsAdd As String
    
    ActiveCell.Offset(-1, -1).Select
    forecastAddSelect = ActiveCell.Address
    lstysAdd = ActiveCell.Offset(0, -1).Address
    lstxsAdd = ActiveCell.Offset(2, -1).Address
    ActiveCell.End(xlToLeft).Select
    fstysAdd = ActiveCell.Address
    fstxsAdd = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToRight).Select
    
    ActiveCell.Offset(2, 0).Formula = "=FORECAST(" & forecastAddSelect & "," & fstxsAdd & ":" & lstxsAdd & "," & fstysAdd & ":" & lstysAdd & ")"
    ActiveCell.Offset(2, 0).NumberFormat = "0%"
    
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(2, 0).Select
    Dim fstToPasteAdd As String
    
    'Storing address to get back
    fstToPasteAdd = ActiveCell.Address
    
    'Creating Global table
    ActiveSheet.UsedRange.Find(what:="Sum of CP%", lookat:=xlWhole).Select
    ActiveSheet.PivotTables("cpGlobalPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(-4, 3).Select
        
    ActiveCell.PasteSpecial xlPasteAll
    ActiveCell.End(xlDown).Select
    
    fstTableAdd = ActiveCell.Offset(1, 0).Address(False, False)
    
    ActiveCell.Offset(1, 1).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstTableAdd = ActiveCell.Address
    ActiveCell.End(xlToLeft).Select
    Range(fstTableAdd, lstTableAdd).Copy
    ActiveCell.Offset(3, 0).Select
    ActiveCell.PasteSpecial xlPasteAll
    ActiveCell.Formula = "=" & fstTableAdd
    ActiveCell.Copy
    
    Dim fstAddToCutForGlobalCP As String
    fstAddToCutForGlobalCP = ActiveCell.Address
    
    fstFormulaAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstFormulaAdd = ActiveCell.Address
    
    Range(fstFormulaAdd, lstFormulaAdd).Select
    Selection.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(1, 0).Value = "Global CP%" & Mid(ActiveCell.Offset(0, 1).Value, 1, 4)
    ActiveCell.Value = ""
    
    ActiveCell.Offset(0, 1).Select
    Do Until Mid(ActiveCell.Offset(0, 1).Value, 1, 4) <> Mid(ActiveCell.Value, 1, 4)
        ActiveCell.Offset(0, 1).Select
    Loop
    
    ActiveCell.Offset(0, 1).Select
    
    fstAddToCut = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddToCut = ActiveCell.Address
    
    Range(fstAddToCut, lstAddToCut).Cut
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(4, 1).Select
    
    ActiveSheet.Paste
    ActiveCell.Offset(1, -1).Value = "Global CP%" & "-" & Mid(ActiveCell.Value, 1, 4)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlToRight).Select
    
    addCompare = ActiveCell.Address(False, False)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=MID(" & addCompare & ",1,4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & addCompare & ",6,2)+1"
    ActiveCell.Copy
    Do Until Mid(ActiveCell.Offset(0, -1).Value, 6, 2) >= 11
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    ActiveCell.End(xlToLeft).Select
    'converting value to date for forecast
    Do Until ActiveCell.Value = ""
        
        ActiveCell.Offset(-1, 0).Value = CDate(ActiveCell.Value & "-" & 1)
        ActiveCell.Offset(0, 1).Select
    Loop
    
    ActiveCell.Offset(-1, -1).Select
    forecastAddSelect = ActiveCell.Address
    lstysAdd = ActiveCell.Offset(0, -1).Address
    lstxsAdd = ActiveCell.Offset(2, -1).Address
    ActiveCell.End(xlToLeft).Select
    fstysAdd = ActiveCell.Address
    fstxsAdd = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToRight).Select
    
    ActiveCell.Offset(2, 0).Formula = "=FORECAST(" & forecastAddSelect & "," & fstxsAdd & ":" & lstxsAdd & "," & fstysAdd & ":" & lstysAdd & ")"
    ActiveCell.Offset(2, 0).NumberFormat = "0%"
    
    fstAddToCut = ActiveCell.Offset(2, 0).Address
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(2, -1).Select
    lstAddToCut = ActiveCell.Address
    
    Range(fstAddToCut, lstAddToCut).Cut
    
    Range(fstToPasteAdd).Select
    ActiveCell.Offset(1, -1).Select
    ActiveSheet.Paste
    
    Range(fstAddToCutForGlobalCP).Select
    ActiveCell.Offset(1, 0).Select
    Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address).Cut
    
    Range(fstAddToCutForGlobal).Select
    ActiveCell.Offset(2, 0).Select
    ActiveSheet.Paste
    
    Dim fstAddForChartAddition As String
    Dim lstAddForChartAddition As String
    
    fstAddForChartAddition = ActiveCell.Offset(-1, 0).Address
    lstAddForChartAddition = ActiveCell.End(xlToRight).Address
    
    ActiveCell.Offset(3, 0).Select
    Dim fstAddForChartCreate As String
    Dim lstAddForChartCreate As String
    
    fstAddForChartCreate = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.End(xlDown).Select
    lstAddForChartCreate = ActiveCell.Address
    
    Range(fstAddForChartAddition).Select
    Dim seriesThreeAdd As String
    Dim seriesThreeFstAdd As String
    Dim seriesThreeLstAdd As String
    Dim seriesFourAdd As String
    Dim seriesFourFstAdd As String
    Dim seriesFourLstAdd As String
    
    ActiveCell.Offset(1, 0).Select
    seriesThreeAdd = ActiveCell.Offset(-1, 0).Address
    seriesThreeFstAdd = ActiveCell.Offset(-1, 1).Address
    seriesFourAdd = ActiveCell.Address
    seriesFourFstAdd = ActiveCell.Offset(0, 1).Address
    ActiveCell.End(xlToRight).Select
    seriesThreeLstAdd = ActiveCell.Offset(-1, 0).Address
    seriesFourLstAdd = ActiveCell.Address
    
    Range(fstAddForChartCreate, lstAddForChartCreate).Select
    ActiveSheet.Shapes.AddChart2(279, xlColumnClustered).Select
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).name = "=Pivot!" & seriesThreeAdd
    ActiveChart.FullSeriesCollection(3).Values = "=Pivot!" & seriesThreeFstAdd & ":" & seriesThreeLstAdd
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).name = "=Pivot!" & seriesFourAdd
    ActiveChart.FullSeriesCollection(4).Values = "=Pivot!" & seriesFourFstAdd & ":" & seriesFourLstAdd
    ActiveChart.FullSeriesCollection(3).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).ChartType = xlLineMarkers
    ActiveChart.FullSeriesCollection(4).ChartType = xlLineMarkers
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.1499999762
    End With
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(4).ApplyDataLabels
    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    ActiveChart.Legend.Position = xlLegendPositionRight
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.ChartTitle.Text = "Contract Pentration - Select Vs Global"
    Application.CutCopyMode = False
     With ActiveChart.Parent
         .Height = 370 ' resize
         .Width = 700  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
    Range("A38").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpGlobalPivotTable") _
        , "Groups (6NC)").Slicers.Add ActiveSheet, , "Groups (6NC) 1", "Groups (6NC)", _
        5, 150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpGlobalPivotTable") _
        , "Material").Slicers.Add ActiveSheet, , "Material 1", "Material", 210, _
        5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpGlobalPivotTable") _
        , "Market").Slicers.Add ActiveSheet, , "Market 1", "Market", 210, 150, _
        144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("cpGlobalPivotTable") _
        , "Country").Slicers.Add ActiveSheet, , "Country 1", "Country", 5, 5 _
        , 144, 198.75
    ActiveSheet.Shapes.Range(Array("Country 1")).Select
    
    Range("A1:X29").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Application.Workbooks(inputFileNameContracts).Close False
    Application.Workbooks(marketInputFile).Close False
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.name = "CP Select Vs Global"
    ActiveWorkbook.Save
    Sheet1.lstBx6NC.MultiSelect = fmMultiSelectSingle
    Sheet1.lstBx6NC.Value = ""
    Sheet1.lstBx6NC.MultiSelect = fmMultiSelectMulti
    Sheet1.comb6NC2.Value = ""
End Sub

Public Sub Creating_Trends_Market_Dynamics(nxtYearFormulaAdd As String)
'Adding Trend above the chart
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find(what:="Sum of �� �Total Contract Revenue", lookat:=xlWhole).Select
    ActiveCell.End(xlToRight).Select
    ActiveSheet.Cells(4, 11).Value = ActiveCell.Value
    Range(nxtYearFormulaAdd).Select
    ActiveSheet.Cells(5, 11).Value = Application.WorksheetFunction.Sum(Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address))
    ActiveSheet.Cells(4, 11).Select
    
    ActiveCell.Offset(-1, 0).Value = "Global"
    ActiveCell.Offset(-1, -1).Value = "Current"
    ActiveCell.Offset(-1, 1).Value = "% Contribution"
    ActiveCell.Offset(0, -2).Value = "ITM"
    ActiveCell.Offset(1, -2).Value = "YTD"
    
    Range(nxtYearFormulaAdd).Select
    ActiveCell.End(xlToRight).Select
    ActiveSheet.Cells(4, 10).Formula = "=" & ActiveCell.Address
    ActiveCell.Offset(2, 0).Select
    ActiveSheet.Cells(5, 10).Formula = "=" & ActiveCell.Address
    ActiveSheet.Cells(4, 12).Select
    ActiveCell.Formula = "=" & ActiveCell.Offset(0, -2).Address(False, False) & "/" & ActiveCell.Offset(0, -1).Address(False, False)
    ActiveCell.NumberFormat = "0%"
    ActiveCell.Copy
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteFormulas
    ActiveCell.NumberFormat = "0%"
    Range(nxtYearFormulaAdd).Select
    Dim fstVerayance As String
    Dim lstVerayance As String
    Dim currentYr As String
    
    ActiveCell.End(xlToRight).Select
    fstVerayance = ActiveCell.Offset(2, 0).Address(False, False)
    lstVerayance = ActiveCell.Offset(-6, 0).Address(False, False)
    currentYr = ActiveCell.Offset(-1, 0).Address(False, False)
    ActiveSheet.Cells(4, 5).Select
    ActiveCell.Formula = "=IFERROR((" & fstVerayance & "-" & lstVerayance & ")/" & lstVerayance & ",)"
    ActiveCell.NumberFormat = "0%"
    ActiveCell.Offset(-1, 0).Value = "VLY"
    ActiveCell.Offset(0, 2).Formula = "=" & currentYr
    ActiveCell.Offset(-1, 2).Value = "Current Month"
    ActiveCell.Offset(-1, 1).Value = "Trend"
    
    'for trend arrow
    Range(nxtYearFormulaAdd).Select
    ActiveCell.End(xlToRight).Select
    Dim fstAddForTrend As String
    Dim lstAddForTrend As String
    
    fstAddForTrend = ActiveCell.Address
    lstAddForTrend = ActiveCell.Offset(0, -1).Address
    
    ActiveSheet.Cells(4, 6).Select
    ActiveCell.Formula = "=IF(" & fstAddForTrend & ">AVERAGE(" & nxtYearFormulaAdd & ":" & lstAddForTrend & "),CHAR(233),IF(" & fstAddForTrend & "<AVERAGE(" & nxtYearFormulaAdd & ":" & lstAddForTrend & "),CHAR(234),CHAR(232)))"
    '=IF(G48>AVERAGE($C$48:F48),CHAR(233),IF(G48<AVERAGE($C$48:F48),CHAR(234),CHAR(232)))
     With Selection.Font
        .name = "Wingdings"
    End With
    
    Range("J4:K5").Select
    Selection.NumberFormat = "#,##0"
    
    Range("E3:L5").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("F4").Select
    With Selection.Font
        .Color = -16711681
        .TintAndShade = 0
    End With
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$F$4=CHAR(233)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$F$4=CHAR(234)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select
    
End Sub

Public Sub Creating_Trend_Drops_Joins()
Range("K2").Select
    ActiveCell.FormulaR1C1 = "Drop"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "Current"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "Global"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "YTD"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "VLY"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Range("K2:N5").Select
    Selection.Copy
    Range("P2").Select
    ActiveSheet.Paste
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Join"
    
    ActiveSheet.Cells(4, 15).Value = lastmonthVal
    ActiveSheet.Cells(3, 15).Value = "Current Month"

    Application.CutCopyMode = False
    
    Dim currMonth As String
    Dim YTDMonth As String
    Dim currmonthDrop As String
    Dim YTDMonthDrop As String
    
    currMonth = Format(lastmonthVal, "mmm") & "-" & Format(lastmonthVal, "yy")
    YTDMonth = "Jan" & "-" & Format(Now(), "yy")
    
    ActiveSheet.UsedRange.Find(what:="Join Total", lookat:=xlWhole).Select
    ActiveSheet.Cells(5, 12).Formula = "=IFERROR(SUM(" & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.End(xlToRight).Address & "),)"
    ActiveCell.Offset(3, 0).Select
    ActiveSheet.Cells(5, 17).Formula = "=IFERROR(SUM(" & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.End(xlToRight).Address & "),)"
    
    ActiveSheet.UsedRange.Find(what:="Join Total", lookat:=xlWhole).Select
    Do Until ActiveCell.Offset(-1, 0).Value = YTDMonth
        ActiveCell.Offset(0, 1).Select
    Loop
    
    Dim fstAddForTrend As String
    Dim lstAddForTrend As String
    
    fstAddForTrend = ActiveCell.Address
    currmonthDrop = ActiveCell.Offset(3, 0).Address
    
    Do Until ActiveCell.Offset(-1, 0).Value = currMonth
        ActiveCell.Offset(0, 1).Select
    Loop
    
    lstAddForTrend = ActiveCell.Address
    YTDMonthDrop = ActiveCell.Offset(3, 0).Address
    
    
    ActiveSheet.Cells(4, 12).Formula = "=Sum(" & fstAddForTrend & ":" & lstAddForTrend & ")"
    ActiveSheet.Cells(4, 17).Formula = "=Sum(" & currmonthDrop & ":" & YTDMonthDrop & ")"
    
    Range(lstAddForTrend).Select
    ActiveSheet.Cells(4, 13).Formula = "=IFERROR((" & ActiveCell.Address & "-" & ActiveCell.Offset(0, -12).Address & ")" & "/" & ActiveCell.Offset(0, -12).Address & ",)"
    ActiveSheet.Cells(4, 18).Formula = "=IFERROR((" & ActiveCell.Offset(3, 0).Address & "-" & ActiveCell.Offset(3, -12).Address & ")" & "/" & ActiveCell.Offset(3, -12).Address & ",)"
    
    ActiveSheet.Cells(4, 14).Formula = "=IF(" & lstAddForTrend & ">AVERAGE(" & fstAddForTrend & ":" & Range(lstAddForTrend).Offset(0, -1).Address & "),CHAR(233),IF(" & lstAddForTrend & "<AVERAGE(" & fstAddForTrend & ":" & Range(lstAddForTrend).Offset(0, -1).Address & "),CHAR(234),CHAR(232)))"
    ActiveSheet.Cells(4, 19).Formula = "=IF(" & YTDMonthDrop & ">AVERAGE(" & currmonthDrop & ":" & Range(YTDMonthDrop).Offset(0, -1).Address & "),CHAR(233),IF(" & YTDMonthDrop & "<AVERAGE(" & currmonthDrop & ":" & Range(YTDMonthDrop).Offset(0, -1).Address & "),CHAR(234),CHAR(232)))"
    
    Range("L4:N4").Copy
    Range("L5").PasteSpecial xlPasteValues
    Range("L4:N4").Copy
    Range("L5").PasteSpecial xlPasteFormats
    Range("Q4:S4").Copy
    Range("Q5").PasteSpecial xlPasteValues
    Range("Q4:S4").Copy
    Range("Q5").PasteSpecial xlPasteFormats
    
    Range("N4").Select
    With Selection.Font
        .name = "Wingdings"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16711681
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("S4").Select
    With Selection.Font
        .name = "Wingdings"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16711681
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("N4").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$N$4=CHAR(233)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$N$4=CHAR(234)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("S4").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$S$4=CHAR(233)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$S$4=CHAR(234)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.ChartObjects("JoinsAndDrops").Activate
    ActiveSheet.ChartObjects("JoinsAndDrops").Activate
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnStacked
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Position = xlLabelPositionInsideBase
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Position = xlLabelPositionInsideBase
    ActiveSheet.Shapes("JoinsAndDrops").ScaleWidth 1.008, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("JoinsAndDrops").ScaleHeight 0.7864285714, msoFalse, _
        msoScaleFromBottomRight
    Range("K2:S5").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("M4").Select
    Selection.Style = "Percent"
    Range("R4").Select
    Selection.Style = "Percent"
    Range("K2:S5").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
End Sub

Public Sub Calculating_Data_Downloaded_Date()
Dim fstPasteRNG As String
Dim lstPasteRNG As String
Dim lookForVal As String


ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.Value = "Fiscal Year/Period"

ActiveCell.Offset(1, 1).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, -1).Select
lookForVal = ActiveCell.Offset(0, 1).Address(False, False)

ActiveCell.Formula = "=MID(" & lookForVal & ", 5, 4)" & "&" & Chr(34) & "-" & Chr(34) & "&" & "MID(" & lookForVal & ", 2, 2)"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)

Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("AC2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range(fstPasteRNG & ":" & lstPasteRNG)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
lastmonthVal = ActiveCell.Value
End Sub
