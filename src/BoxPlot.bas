Attribute VB_Name = "BoxPlot"
Public Sub BoxPlot()

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

'On Error Resume Next

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

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "Contracts_BoxPlot_" & Format(Now, "mmmyy") & ".xlsm"
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
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R50C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)

'change style of the new PivotTable:
pvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
pvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
pvtTbl.ManualUpdate = True

Dim pvtTblName As String
pvtTblName = pvtTbl.name

    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Market")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Range("A2").Select
    With ActiveSheet.PivotTables(pvtTblName)
        .ColumnGrand = False
        .RowGrand = False
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C] Contract Material Line Item")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
        pvtTblName).PivotFields("    Contract" & Chr(10) & "Net Value"), _
        "Count of     Contract" & Chr(10) & "Net Value", xlCount
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Count of     Contract" & Chr(10) & "Net Value")
        .Caption = "Sum of     Contract"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C] Contract Material Line Item").AutoSort xlDescending, "Sum of     Contract"
        
    pvtTbl.ManualUpdate = False
    
    ActiveSheet.Cells(50, 1).Select
    Dim fstAdd As String
    Dim lstAdd As String
    
    fstAdd = ActiveCell.Offset(1, 1).Address(False, False)
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveSheet.Cells(30, 1).Value = "Product Group"
    ActiveSheet.Cells(31, 1).Value = "Price SWO's"
    ActiveSheet.Cells(32, 1).Value = "Mean"
    ActiveSheet.Cells(33, 1).Value = "Min"
    ActiveSheet.Cells(34, 1).Value = "Q1"
    ActiveSheet.Cells(35, 1).Value = "Median"
    ActiveSheet.Cells(36, 1).Value = "P95"
    ActiveSheet.Cells(37, 1).Value = "Max"
    ActiveSheet.Cells(39, 1).Value = "25th PCT"
    ActiveSheet.Cells(40, 1).Value = "50th PCT"
    ActiveSheet.Cells(41, 1).Value = "95th PCT"
    ActiveSheet.Cells(43, 1).Value = "Min"
    ActiveSheet.Cells(44, 1).Value = "Max"
    
    ActiveSheet.Cells(30, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=" & fstAdd
    ActiveCell.Offset(1, 0).Formula = "=IFERROR(SUM(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(2, 0).Formula = "=IFERROR(AVERAGE(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(3, 0).Formula = "=IFERROR(Min(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(4, 0).Formula = "=IFERROR(PERCENTILE.EXC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.25),)"
    ActiveCell.Offset(5, 0).Formula = "=IFERROR(MEDIAN(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(6, 0).Formula = "=IFERROR(PERCENTILE.EXC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.95),)"
    ActiveCell.Offset(7, 0).Formula = "=IFERROR(MAX(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(9, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(10, 0).Formula = "=IFERROR(" & ActiveCell.Offset(5, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(11, 0).Formula = "=IFERROR(" & ActiveCell.Offset(6, 0).Address(False, False) & "-" & ActiveCell.Offset(5, 0).Address(False, False) & ",)"
    ActiveCell.Offset(13, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & "-" & ActiveCell.Offset(3, 0).Address(False, False) & ",)"
    ActiveCell.Offset(14, 0).Formula = "=IFERROR(" & ActiveCell.Offset(7, 0).Address(False, False) & "-" & ActiveCell.Offset(6, 0).Address(False, False) & ",)"
    
    Range(ActiveCell.Address, ActiveCell.Offset(14, 1).Address).Copy
    Do Until ActiveCell.Offset(21, 0).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    
End Sub

