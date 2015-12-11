Attribute VB_Name = "BoxPlot"
Public Sub BoxPlot_Calculations()

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

'Selection for Modality
If Sheet1.combModality.value = "Modality" Or Sheet1.combModality.value = "" Then
     MsgBox "Please Select a Modality Group!"
     End
End If

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
'marketInputFile = "Market_Groups_Markets_Country.xlsx"
'marketInputFile = Replace(inputRevenue, inputFileNameContracts, marketInputFile)
'Application.Workbooks.Open (marketInputFile), False

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


Workbooks(revenueOutputGlobal).Activate
Dim ws As Worksheet
For Each ws In Workbooks(revenueOutputGlobal).Sheets
    If ws.name = "BoxPlot-Revenue" Or ws.name = "Data" Or ws.name = "BoxPlot-Cost" Or ws.name = "RevenueVsMarket" Or ws.name = "Filtered-Data-Revenue" Or ws.name = "Filtered-Data-Cost" Or ws.name = "CostVsMarket" Then
        ws.Delete
    End If
Next

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, after:=ActiveCell).Select

'Putting names in blank cells
Do Until ActiveCell.Offset(1, 0).value = "" And ActiveCell.Offset(0, 1).value = ""
    If ActiveCell.value = "" Then
        ActiveCell.value = ActiveCell.Offset(0, -1).value & " " & "A"
        ActiveCell.Offset(0, 1).Select
    Else
        ActiveCell.Offset(0, 1).Select
    End If
    
    If ActiveCell.value = "EUR" Then
        ActiveCell.value = ActiveCell.Offset(-1, 0).value
    End If
Loop

ActiveSheet.Cells(1, 1).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole, after:=ActiveCell).Select
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
'marketInputFile = "Market_Groups_Markets_Country.xlsx"

modalityVal = Sheet1.combModality.value

'Application.Workbooks(marketInputFile).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("Markets").Activate
Dim revenueval As String
revenueval = ActiveSheet.Cells(2, 16).value
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
Dim marketFSTAdd As String
Dim marketLSTAdd As String

marketFSTAdd = ActiveCell.Address
ActiveCell.Offset(0, 3).Select
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
ActiveCell.Offset(0, -1).value = "System Code (6NC)"

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
ActiveCell.Formula = "=IFERROR(IF(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "4" & "," & "False)=" & Chr(34) & modalityVal & Chr(34) & ",VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")," & Chr(34) & "Others" & Chr(34) & ")"
'"=IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "2" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

'Adding Market column

'Application.Workbooks(marketInputFile).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("Markets").Activate
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
ActiveCell.Offset(0, -1).value = "Market"

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

'adding Fiscal Year/Period Column
Application.Workbooks(revenueOutputGlobal).Activate
ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.value = "Fiscal Year/Period"

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

'Adding contracts type

'Application.Workbooks(marketInputFile).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("Markets").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="Contract Type Material", lookat:=xlWhole).Select

marketFSTAdd = ActiveCell.Address
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

ActiveSheet.UsedRange.Find(what:="[C] Contract Material Line Item", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C] Contract Material Line Item", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).value = "Contract Type Material"

rngStringMarket = marketRNG.Address
ActiveCell.Offset(1, 0).Select
fstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstPasteRNG = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
lookForVal = ActiveCell.Address(False, False)

ActiveCell.Offset(0, -1).Select
ActiveCell.Formula = "=IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "1" & "," & "False)," & Chr(34) & "Others" & Chr(34) & ")"
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

ActiveWorkbook.Sheets("Data").Activate
'Set wsData = Worksheets("Data")
'
'
'lastRow = wsData.Cells(rows.Count, 1).End(xlUp).Row
'lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
'
'Set rngData = wsData.Cells(1, 1).Resize(lastRow, lastColumn)
'rngDataForPivot = rngData.Address
'rngData.Select
'
'Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!" & rngDataForPivot, Version:=xlPivotTableVersion15)
'Sheets.Add
'ActiveSheet.name = "Pivot"
'Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R50C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'
'
'pvtTbl.TableStyle2 = "PivotStyleMedium3"
'
'
'pvtTbl.InGridDropZones = True
'
'pvtTbl.ManualUpdate = True
'
'Dim pvtTblName As String
'pvtTblName = pvtTbl.name
'
'   With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'        "Contract Type Material")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    Range("A52").Select
'    ActiveSheet.PivotTables(pvtTblName).PivotFields("Contract Type Material"). _
'        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
'        False, False)
'    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'        "[C,S] Reference Equipment")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'        "[C,S] Reference Equipment").Subtotals = Array(False, False, False, False, _
'        False, False, False, False, False, False, False, False)
'    Range("A54").Select
'    With ActiveSheet.PivotTables(pvtTblName)
'        .ColumnGrand = False
'        .RowGrand = False
'    End With
'    With ActiveSheet.PivotTables(pvtTblName)
'        .InGridDropZones = True
'        .RowAxisLayout xlTabularRow
'    End With
'    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Fiscal Year/Period")
'        .Orientation = xlColumnField
'        .Position = 1
'    End With
'    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
'        pvtTblName).PivotFields("    Total Contract Revenue"), _
'        "Count of     Total Contract Revenue", xlCount
'    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'        "Count of     Total Contract Revenue")
'        .Caption = "Sum of     Total Contract Revenue"
'        .Function = xlSum
'    End With
'    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
'        "Contract Type Material").AutoSort xlDescending, _
'        "Sum of     Total Contract Revenue"
'
'    pvtTbl.ManualUpdate = False
'
'    ActiveSheet.Cells(51, 2).Select
'    Dim fstAdd As String
'    Dim lstAdd As String
'    Dim tenthAdd As String
'    Dim lineItemAdd As String
'
'    lineItemAdd = ActiveCell.Address(False, False)
'    fstAdd = ActiveCell.Offset(0, 1).Address(False, False)
'
'    ActiveCell.End(xlDown).Select
'    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
'
'    ActiveSheet.Cells(30, 2).value = "Product Group"
'    ActiveSheet.Cells(31, 2).value = "Revenue"
'    ActiveSheet.Cells(32, 2).value = "Mean"
'    ActiveSheet.Cells(33, 2).value = "Min"
'    ActiveSheet.Cells(34, 2).value = "Q1"
'    ActiveSheet.Cells(35, 2).value = "Median"
'    ActiveSheet.Cells(36, 2).value = "P75"
'    ActiveSheet.Cells(37, 2).value = "P95"
'    ActiveSheet.Cells(38, 2).value = "Max"
'    ActiveSheet.Cells(40, 2).value = "25th PCT"
'    ActiveSheet.Cells(41, 2).value = "50th PCT"
'    ActiveSheet.Cells(42, 2).value = "75th PCT"
'    ActiveSheet.Cells(43, 2).value = "95th PCT"
'    ActiveSheet.Cells(45, 2).value = "Min"
'    ActiveSheet.Cells(46, 2).value = "Max"
'
'    ActiveSheet.Cells(30, 2).Select
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Formula = "=" & fstAdd
'    ActiveCell.Offset(1, 0).Formula = "=IFERROR(SUBTOTAL(109," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
'    ActiveCell.Offset(2, 0).Formula = "=IFERROR(SUBTOTAL(1," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
'    ActiveCell.Offset(3, 0).Formula = "=IFERROR(Min(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
'    ActiveCell.Offset(4, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.25),)"
'    ActiveCell.Offset(5, 0).Formula = "=IFERROR(MEDIAN(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
'    ActiveCell.Offset(6, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.75),)"
'    ActiveCell.Offset(7, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.95),)"
'    ActiveCell.Offset(8, 0).Formula = "=IFERROR(MAX(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
'    ActiveCell.Offset(10, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
'    ActiveCell.Offset(11, 0).Formula = "=IFERROR(" & ActiveCell.Offset(5, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
'    ActiveCell.Offset(12, 0).Formula = "=IFERROR(" & ActiveCell.Offset(6, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
'    ActiveCell.Offset(13, 0).Formula = "=IFERROR(" & ActiveCell.Offset(7, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
'    ActiveCell.Offset(15, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & "-" & ActiveCell.Offset(3, 0).Address(False, False) & ",)"
'    ActiveCell.Offset(16, 0).Formula = "=IFERROR(" & ActiveCell.Offset(8, 0).Address(False, False) & "-" & ActiveCell.Offset(7, 0).Address(False, False) & ",)"
'
'    Dim fstChartAdd As String
'    Dim lstChartAdd As String
'    Dim fstAddToCopy As String
'    Dim lstAddToCopy As String
'    Dim chartName As String
'
'    fstAddToCopy = ActiveCell.Address
'    lstAddToCopy = ActiveCell.Offset(16, 0).Address
'    fstChartAdd = ActiveCell.Offset(0, -1).Address
'
'    ActiveSheet.Range(fstAddToCopy, lstAddToCopy).Copy
'
'    Do Until ActiveCell.Offset(21, 1).value = ""
'        ActiveCell.Offset(0, 1).Select
'        ActiveCell.PasteSpecial xlPasteFormulas
'    Loop
'
'    lstChartAdd = ActiveCell.Offset(16, 0).Address
'    Dim minAdd As Range
'    Dim maxAdd As String
'    Dim chartRNG As Range
'    Set chartRNG = Range("Pivot!" & fstChartAdd & ":" & lstChartAdd)
'    Cells(45, 3).Select
'    Set minAdd = Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address)
'    ActiveCell.Offset(1, 0).Select
'    maxAdd = Range(ActiveCell, ActiveCell.End(xlToRight)).Address
'    chartRNG.Select
'    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
'    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked).Select
'    ActiveChart.SetSourceData Source:=chartRNG
'    ActiveChart.PlotBy = xlRows
'    chartName = ActiveChart.name
'    With ActiveChart.Parent
'         .Height = 300 ' resize
'         .Width = 600  ' resize
'         .Top = 10    ' reposition
'         .Left = 300   ' reposition
'     End With
'
'    For i = 1 To 9
'        ActiveChart.FullSeriesCollection(i).IsFiltered = True
'    Next
'    For i = 14 To 16
'        ActiveChart.FullSeriesCollection(i).IsFiltered = True
'    Next
'
'    ActiveChart.ChartTitle.Text = "Box Plot Revenue"
'    ActiveSheet.ChartObjects("Chart 1").Activate
'    ActiveChart.ChartArea.Select
'    ActiveChart.FullSeriesCollection(13).HasErrorBars = True
'    ActiveChart.FullSeriesCollection(13).ErrorBars.Select
'    ActiveChart.FullSeriesCollection(13).ErrorBar Direction:=xlY, Include:= _
'        xlPlusValues, Type:=xlCustom, Amount:=minAdd.value
'
'    ActiveChart.Axes(xlCategory).Select
'    Selection.TickLabelPosition = xlLow
'    ActiveChart.Axes(xlValue).Select
'    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
'    ActiveChart.ChartArea.Select
'    ActiveChart.ChartStyle = 279
'    ActiveSheet.ChartObjects("Chart 1").Select
'    ActiveChart.ChartArea.Font.Size = 12
'    Selection.Placement = xlFreeFloating
'        ActiveChart.Legend.Select
'    Selection.Position = xlTop
'
'    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
'        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC)", _
'        "System Code (6NC)", 5, 5, 144, 198.75
'    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
'        "[C,S] System Code Material (Material no of  R Eq)").Slicers.Add ActiveSheet, _
'        , "[C,S] System Code Material (Material no of  R Eq)", _
'        "[C,S] System Code Material (Material no of  R Eq)", 210, 150, 144, 198.75
'    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
'        "Contract Type Material").Slicers.Add ActiveSheet, , _
'        "Contract Type Material", "Contract Type Material", 5, 150 _
'        , 144, 198.75
'        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
'        "Market").Slicers.Add ActiveSheet, , "Market", "Market", 210, 5, 144, _
'        198.75
'    ActiveSheet.Shapes.Range(Array("Market")).Select
'    ActiveSheet.Shapes.Range(Array("Contract Type Material")).Select
'
'Range("A1:P28").Select
'With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorAccent1
'        .TintAndShade = -0.499984740745262
'        .PatternTintAndShade = 0
'    End With
'ActiveSheet.Cells(1, 1).Select
'ActiveSheet.name = "BoxPlot-Revenue"
'
'Sheets("BoxPlot-Revenue").Select
'    Sheets("BoxPlot-Revenue").Copy Before:=Sheets(2)
'    ActiveSheet.name = "BoxPlot-Cost"
'    Range("A51").Select
'    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
'        "PivotTable1").PivotFields("   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"), "Count of   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", _
'        xlCount
'    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
'        "Sum of     Total Contract Revenue").Orientation = xlHidden
'    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
'        "Count of   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract")
'        .Caption = "Sum of Sum of    Total"
'        .Function = xlSum
'    End With
'    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
'        "[C] Contract Material Line Item").AutoSort xlDescending, "   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"
'
'    ActiveSheet.ChartObjects("Chart 1").Activate
'    ActiveChart.ChartTitle.Text = "Box Plot Cost"
'    ActiveSheet.Cells(1, 1).Select
'
    MarketVsRevenue_Boxplot inputFileNameContracts, revenueval
    
    MarketVsCost_Boxplot inputFileNameContracts, revenueval
    
    Sheets("CostVsMarket").Move Before:=Sheets(2)
    Sheets("RevenueVsMarket").Select
End Sub

Public Sub MarketVsRevenue_Boxplot(inputFileNameContracts As String, revenueval As String)
'Creating Market Vs Revenue BoxPlots
Sheets("Data").Activate
Sheets("Data").Select
ActiveSheet.UsedRange.Find(what:="    Total Contract Revenue", lookat:=xlWhole).Select
Dim cellFld As Integer
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:=">=0" _
        , Operator:=xlAnd, Criteria2:="<=" & revenueval
        
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Start Date (Header)", lookat:=xlWhole).Select
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:="<>#"

'Calculating duration
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight

ActiveCell.value = "Duration"
ActiveCell.Offset(1, 0).Select
ActiveCell.Formula = "=(DATE(RIGHT(" & ActiveCell.Offset(0, -1).Address(False, False) & ",4),MID(" & ActiveCell.Offset(0, -1).Address(False, False) & ",4,2),MID(" & ActiveCell.Offset(0, -1).Address(False, False) & ",1,2))-DATE(RIGHT(" & ActiveCell.Offset(0, -2).Address(False, False) & ",4),MID(" & ActiveCell.Offset(0, -2).Address(False, False) & ",4,2),MID(" & ActiveCell.Offset(0, -2).Address(False, False) & ",1,2)))/30"
Dim fstDateDiffAdd As String
Dim lstDateDiffAdd As String
ActiveCell.Copy
fstDateDiffAdd = ActiveCell.Offset(1, 0).Address
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
lstDateDiffAdd = ActiveCell.Offset(0, 1).Address
Range(fstDateDiffAdd, lstDateDiffAdd).PasteSpecial xlPasteFormulas

ActiveSheet.UsedRange.Find(what:="Duration", lookat:=xlWhole).Select
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:=">6"

ActiveCell.SpecialCells(xlCellTypeVisible).Select
Selection.Copy
Sheets.Add
ActiveSheet.Paste
ActiveSheet.name = "Filtered-Data-Revenue"
    Sheets("Filtered-Data-Revenue").Activate

ActiveSheet.UsedRange.Find(what:="Duration", lookat:=xlWhole).Select
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:=">13"

ActiveSheet.UsedRange.Find(what:="    Total Contract Revenue", lookat:=xlWhole).Select
Dim fstcel As Integer
Dim lstcel As Integer
Dim diffCel As Integer
fstcel = ActiveCell.Column
ActiveSheet.UsedRange.Find(what:="Duration", lookat:=xlWhole).Select
lstcel = ActiveCell.Column
diffCel = fstcel - lstcel
ActiveCell.AutoFilter

ActiveSheet.UsedRange.Find(what:="    Total Contract Revenue", lookat:=xlWhole).Select

ActiveCell.EntireColumn.Insert xlToRight
ActiveCell.Offset(1, 0).Select
ActiveCell.Formula = "=IF(" & ActiveCell.Offset(0, -diffCel).Address(False, False) & "," & ActiveCell.Offset(0, 1).Address(False, False) & "/" & ActiveCell.Offset(0, -diffCel).Address(False, False) & "*12," & ActiveCell.Offset(0, 1).Address(False, False) & ")"
ActiveCell.Copy
fstDateDiffAdd = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
lstDateDiffAdd = ActiveCell.Offset(0, -1).Address
Range(fstDateDiffAdd, lstDateDiffAdd).PasteSpecial xlPasteFormulas
Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Copy
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteValues
ActiveCell.Offset(0, -1).Select
ActiveCell.EntireColumn.Delete

Dim tblRNG As String
Dim tblAdd As String

ActiveSheet.UsedRange.Select
tblAdd = Selection.Address
tblRNG = Application.ConvertFormula(Formula:=tblAdd, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlR1C1)
tblRNG = "Filtered-Data-Revenue!" & tblRNG
Sheets.Add
    ActiveSheet.name = "Pivot"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        tblRNG, Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Pivot!R50C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15

Dim pvtTblName As String
ActiveSheet.Cells(50, 1).Select
pvtTblName = ActiveCell.PivotTable.name
   
   With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Contract Type Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("A52").Select
    ActiveSheet.PivotTables(pvtTblName).PivotFields("Contract Type Material"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName)
        .ColumnGrand = False
        .RowGrand = False
    End With
    With ActiveSheet.PivotTables(pvtTblName)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Market")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
        pvtTblName).PivotFields("    Total Contract Revenue"), _
        "Count of     Total Contract Revenue", xlCount
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Count of     Total Contract Revenue")
        .Caption = "Sum of     Total Contract Revenue"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Contract Type Material").AutoSort xlDescending, _
        "Sum of     Total Contract Revenue"
    ActiveSheet.PivotTables(pvtTblName).PivotFields("Market").ShowAllItems = _
        True
    ActiveCell.PivotTable.ManualUpdate = False
    
    ActiveSheet.Cells(51, 2).Select
    Dim fstAdd As String
    Dim lstAdd As String
    Dim tenthAdd As String
    Dim lineItemAdd As String
    
    lineItemAdd = ActiveCell.Address(False, False)
    fstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveSheet.Cells(30, 2).value = "Product Group"
    ActiveSheet.Cells(31, 2).value = "Revenue"
    ActiveSheet.Cells(32, 2).value = "Mean"
    ActiveSheet.Cells(33, 2).value = "Min"
    ActiveSheet.Cells(34, 2).value = "Q1"
    ActiveSheet.Cells(35, 2).value = "Median"
    ActiveSheet.Cells(36, 2).value = "P75"
    ActiveSheet.Cells(37, 2).value = "P95"
    ActiveSheet.Cells(38, 2).value = "Max"
    ActiveSheet.Cells(40, 2).value = "25th PCT"
    ActiveSheet.Cells(41, 2).value = "50th PCT"
    ActiveSheet.Cells(42, 2).value = "75th PCT"
    ActiveSheet.Cells(43, 2).value = "95th PCT"
    ActiveSheet.Cells(45, 2).value = "Min"
    ActiveSheet.Cells(46, 2).value = "Max"
    
    ActiveSheet.Cells(30, 2).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=" & fstAdd
    ActiveCell.Offset(1, 0).Formula = "=IFERROR(SUBTOTAL(109," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(2, 0).Formula = "=IFERROR(SUBTOTAL(1," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(3, 0).Formula = "=IFERROR(Min(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(4, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.25),)"
    ActiveCell.Offset(5, 0).Formula = "=IFERROR(MEDIAN(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(6, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.75),)"
    ActiveCell.Offset(7, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.95),)"
    ActiveCell.Offset(8, 0).Formula = "=IFERROR(MAX(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(10, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(11, 0).Formula = "=IFERROR(" & ActiveCell.Offset(5, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(12, 0).Formula = "=IFERROR(" & ActiveCell.Offset(6, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(13, 0).Formula = "=IFERROR(" & ActiveCell.Offset(7, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(15, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & "-" & ActiveCell.Offset(3, 0).Address(False, False) & ",)"
    ActiveCell.Offset(16, 0).Formula = "=IFERROR(" & ActiveCell.Offset(8, 0).Address(False, False) & "-" & ActiveCell.Offset(7, 0).Address(False, False) & ",)"
    
    Dim fstChartAdd As String
    Dim lstChartAdd As String
    Dim fstAddToCopy As String
    Dim lstAddToCopy As String
    Dim chartName As String
    
    fstAddToCopy = ActiveCell.Address
    lstAddToCopy = ActiveCell.Offset(16, 0).Address
    fstChartAdd = ActiveCell.Offset(0, -1).Address
    
    ActiveSheet.Range(fstAddToCopy, lstAddToCopy).Copy
    
    Do Until ActiveCell.Offset(21, 1).value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    lstChartAdd = ActiveCell.Offset(16, 0).Address
    Dim minAdd As Range
    Dim maxAdd As String
    Dim chartRNG As Range
    Set chartRNG = Range("Pivot!" & fstChartAdd & ":" & lstChartAdd)
    Cells(45, 3).Select
    Set minAdd = Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address)
    ActiveCell.Offset(1, 0).Select
    maxAdd = Range(ActiveCell, ActiveCell.End(xlToRight)).Address
    chartRNG.Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked).Select
    ActiveChart.SetSourceData Source:=chartRNG
    ActiveChart.PlotBy = xlRows
    chartName = ActiveChart.name
    With ActiveChart.Parent
         .Height = 300 ' resize
         .Width = 600  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
    
    For i = 1 To 9
        ActiveChart.FullSeriesCollection(i).IsFiltered = True
    Next
    For i = 14 To 16
        ActiveChart.FullSeriesCollection(i).IsFiltered = True
    Next
    
    ActiveChart.ChartTitle.Text = "Box Plot Revenue"
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(13).HasErrorBars = True
    ActiveChart.FullSeriesCollection(13).ErrorBars.Select
    ActiveChart.FullSeriesCollection(13).ErrorBar Direction:=xlY, Include:= _
        xlPlusValues, Type:=xlCustom, Amount:=minAdd.value

    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlLow
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.ChartArea.Select
    ActiveChart.ChartStyle = 279
    ActiveSheet.ChartObjects("Chart 1").Select
    ActiveChart.ChartArea.Font.Size = 12
    Selection.Placement = xlFreeFloating
        ActiveChart.Legend.Select
    Selection.Position = xlTop
    
Range("A1:P28").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
ActiveSheet.Cells(1, 1).Select
ActiveSheet.name = "RevenueVsMarket"

ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC) 2", _
        "System Code (6NC)", 5, 5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "[C,S] Reference Equipment").Slicers.Add ActiveSheet, , _
        "[C,S] Reference Equipment", "[C,S] Reference Equipment", 210, 150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Contract Type Material").Slicers.Add ActiveSheet, , _
        "Contract Type Material 2", "Contract Type Material", 5, _
        150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Fiscal Year/Period").Slicers.Add ActiveSheet, , "Fiscal Year/Period", _
        "Fiscal Year/Period", 210, 5, 144, 198.75


End Sub

Public Sub MarketVsCost_Boxplot(inputFileNameContracts As String, revenueval As String)

'Creating Market Vs Cost BoxPlots
Sheets("Data").Activate
Sheets("Data").Select
ActiveCell.AutoFilter
ActiveSheet.UsedRange.Find(what:="   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", lookat:=xlWhole).Select
Dim cellFld As Integer
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:=">=0" _
        , Operator:=xlAnd, Criteria2:="<=" & revenueval
        
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Start Date (Header)", lookat:=xlWhole).Select
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:="<>#"

ActiveSheet.UsedRange.Find(what:="Duration", lookat:=xlWhole).Select
cellFld = ActiveCell.Column
ActiveCell.AutoFilter Field:=cellFld, Criteria1:=">6"

ActiveCell.SpecialCells(xlCellTypeVisible).Select
Selection.Copy
Sheets.Add
ActiveSheet.Paste
ActiveSheet.name = "Filtered-Data-Cost"
    Sheets("Filtered-Data-Cost").Activate

ActiveSheet.UsedRange.Find(what:="   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", lookat:=xlWhole).Select
Dim fstcel As Integer
Dim lstcel As Integer
Dim diffCel As Integer
fstcel = ActiveCell.Column
ActiveSheet.UsedRange.Find(what:="Duration", lookat:=xlWhole).Select
lstcel = ActiveCell.Column
diffCel = fstcel - lstcel
ActiveCell.AutoFilter

ActiveSheet.UsedRange.Find(what:="   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", lookat:=xlWhole).Select

ActiveCell.EntireColumn.Insert xlToRight
ActiveCell.Offset(1, 0).Select
ActiveCell.Formula = "=IF(" & ActiveCell.Offset(0, -diffCel).Address(False, False) & "," & ActiveCell.Offset(0, 1).Address(False, False) & "/" & ActiveCell.Offset(0, -diffCel).Address(False, False) & "*12," & ActiveCell.Offset(0, 1).Address(False, False) & ")"
ActiveCell.Copy
fstDateDiffAdd = ActiveCell.Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
lstDateDiffAdd = ActiveCell.Offset(0, -1).Address
Range(fstDateDiffAdd, lstDateDiffAdd).PasteSpecial xlPasteFormulas
Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Copy
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteValues
ActiveCell.Offset(0, -1).Select
ActiveCell.EntireColumn.Delete

Dim tblRNG As String
Dim tblAdd As String

ActiveSheet.UsedRange.Select
tblAdd = Selection.Address
tblRNG = Application.ConvertFormula(Formula:=tblAdd, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlR1C1)
tblRNG = "Filtered-Data-Cost!" & tblRNG
Sheets.Add
    ActiveSheet.name = "Pivot"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        tblRNG, Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Pivot!R50C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15

Dim pvtTblName As String
ActiveSheet.Cells(50, 1).Select
pvtTblName = ActiveCell.PivotTable.name
   
   With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Contract Type Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("A52").Select
    ActiveSheet.PivotTables(pvtTblName).PivotFields("Contract Type Material"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName)
        .ColumnGrand = False
        .RowGrand = False
    End With
    With ActiveSheet.PivotTables(pvtTblName)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Market")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
        pvtTblName).PivotFields("   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Count of   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of   swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Contract Type Material").AutoSort xlDescending, _
        "Sum of     Total Contract Revenue"
    ActiveSheet.PivotTables(pvtTblName).PivotFields("Market").ShowAllItems = _
        True
    ActiveCell.PivotTable.ManualUpdate = False
    
    ActiveSheet.Cells(51, 2).Select
    Dim fstAdd As String
    Dim lstAdd As String
    Dim tenthAdd As String
    Dim lineItemAdd As String
    
    lineItemAdd = ActiveCell.Address(False, False)
    fstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    
    ActiveSheet.Cells(30, 2).value = "Product Group"
    ActiveSheet.Cells(31, 2).value = "Cost"
    ActiveSheet.Cells(32, 2).value = "Mean"
    ActiveSheet.Cells(33, 2).value = "Min"
    ActiveSheet.Cells(34, 2).value = "Q1"
    ActiveSheet.Cells(35, 2).value = "Median"
    ActiveSheet.Cells(36, 2).value = "P75"
    ActiveSheet.Cells(37, 2).value = "P95"
    ActiveSheet.Cells(38, 2).value = "Max"
    ActiveSheet.Cells(40, 2).value = "25th PCT"
    ActiveSheet.Cells(41, 2).value = "50th PCT"
    ActiveSheet.Cells(42, 2).value = "75th PCT"
    ActiveSheet.Cells(43, 2).value = "95th PCT"
    ActiveSheet.Cells(45, 2).value = "Min"
    ActiveSheet.Cells(46, 2).value = "Max"
    
    ActiveSheet.Cells(30, 2).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "=" & fstAdd
    ActiveCell.Offset(1, 0).Formula = "=IFERROR(SUBTOTAL(109," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(2, 0).Formula = "=IFERROR(SUBTOTAL(1," & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(3, 0).Formula = "=IFERROR(Min(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(4, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.25),)"
    ActiveCell.Offset(5, 0).Formula = "=IFERROR(MEDIAN(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(6, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.75),)"
    ActiveCell.Offset(7, 0).Formula = "=IFERROR(PERCENTILE.INC(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & ",0.95),)"
    ActiveCell.Offset(8, 0).Formula = "=IFERROR(MAX(" & Range(fstAdd).Offset(1, 0).Address(False, False) & ":" & lstAdd & "),)"
    ActiveCell.Offset(10, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(11, 0).Formula = "=IFERROR(" & ActiveCell.Offset(5, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(12, 0).Formula = "=IFERROR(" & ActiveCell.Offset(6, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(13, 0).Formula = "=IFERROR(" & ActiveCell.Offset(7, 0).Address(False, False) & "-" & ActiveCell.Offset(4, 0).Address(False, False) & ",)"
    ActiveCell.Offset(15, 0).Formula = "=IFERROR(" & ActiveCell.Offset(4, 0).Address(False, False) & "-" & ActiveCell.Offset(3, 0).Address(False, False) & ",)"
    ActiveCell.Offset(16, 0).Formula = "=IFERROR(" & ActiveCell.Offset(8, 0).Address(False, False) & "-" & ActiveCell.Offset(7, 0).Address(False, False) & ",)"
    
    Dim fstChartAdd As String
    Dim lstChartAdd As String
    Dim fstAddToCopy As String
    Dim lstAddToCopy As String
    Dim chartName As String
    
    fstAddToCopy = ActiveCell.Address
    lstAddToCopy = ActiveCell.Offset(16, 0).Address
    fstChartAdd = ActiveCell.Offset(0, -1).Address
    
    ActiveSheet.Range(fstAddToCopy, lstAddToCopy).Copy
    
    Do Until ActiveCell.Offset(21, 1).value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    lstChartAdd = ActiveCell.Offset(16, 0).Address
    Dim minAdd As Range
    Dim maxAdd As String
    Dim chartRNG As Range
    Set chartRNG = Range("Pivot!" & fstChartAdd & ":" & lstChartAdd)
    Cells(45, 3).Select
    Set minAdd = Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address)
    ActiveCell.Offset(1, 0).Select
    maxAdd = Range(ActiveCell, ActiveCell.End(xlToRight)).Address
    chartRNG.Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked).Select
    ActiveChart.SetSourceData Source:=chartRNG
    ActiveChart.PlotBy = xlRows
    chartName = ActiveChart.name
    With ActiveChart.Parent
         .Height = 300 ' resize
         .Width = 600  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
    
    For i = 1 To 9
        ActiveChart.FullSeriesCollection(i).IsFiltered = True
    Next
    For i = 14 To 16
        ActiveChart.FullSeriesCollection(i).IsFiltered = True
    Next
    
    ActiveChart.ChartTitle.Text = "Box Plot Cost"
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(13).HasErrorBars = True
    ActiveChart.FullSeriesCollection(13).ErrorBars.Select
    ActiveChart.FullSeriesCollection(13).ErrorBar Direction:=xlY, Include:= _
        xlPlusValues, Type:=xlCustom, Amount:=minAdd.value

    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlLow
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.ChartArea.Select
    ActiveChart.ChartStyle = 279
    ActiveSheet.ChartObjects("Chart 1").Select
    ActiveChart.ChartArea.Font.Size = 12
    Selection.Placement = xlFreeFloating
        ActiveChart.Legend.Select
    Selection.Position = xlTop
    
Range("A1:P28").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
ActiveSheet.Cells(1, 1).Select
ActiveSheet.name = "CostVsMarket"

ActiveSheet.PivotTables("PivotTable1").PivotSelect _
        "'[C,S] Reference Equipment'[All]", xlLabelOnly, True
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC)", _
        "System Code (6NC)", 5, 5, 144, 198.75
ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Market").Slicers.Add ActiveSheet, , "Market", "Market", 210, 150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Contract Type Material").Slicers.Add ActiveSheet, , "Contract Type Material", _
        "Contract Type Material", 5, _
        150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTable1"), _
        "Fiscal Year/Period").Slicers.Add ActiveSheet, , "Fiscal Year/Period 1", _
        "Fiscal Year/Period", 210, 5, 144, 198.75

    ThisWorkbook.Sheets("UI").Activate
Application.Workbooks(revenueOutputGlobal).Save
Workbooks(inputFileNameContracts).Close False
'Workbooks(marketInputFile).Close False

End Sub
