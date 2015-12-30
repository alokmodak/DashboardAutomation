Attribute VB_Name = "PieChartSWOCost"
Public Sub PieChart_SWOCost()

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
Dim modalityVal As String

'On Error Resume Next

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

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "PieChart_SWOCost_" & Format(Now, "mmmyy") & ".xlsx"
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
    If ws.name = "SWO_Cost_PieChart" Or ws.name = "Pivot" Or ws.name = "Data" Then
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

Dim pvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim pvtFld As PivotField
Dim lastRow
Dim lastColumn
Dim rngDataForPivot As String
Dim pvtItem As PivotItem

'marketInputFile = "Market_Groups_Markets_Country.xlsx"
'
'Application.Workbooks(marketInputFile).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("Markets").Activate
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
ActiveCell.Offset(0, -1).value = "Market"

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

'Adding 6NC Names column
'marketInputFile = "Market_Groups_Markets_Country.xlsx"
modalityVal = Sheet1.combModality.value

'Application.Workbooks(marketInputFile).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("Markets").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.AutoFilter 'two times autofilter to clear all the filters
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select

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
ActiveCell.PasteSpecial xlPasteAll
Set marketRNG = Range(Selection.Address)

ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert xlToRight
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.Offset(0, -1).value = "System Code (6NC)"

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
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.UsedRange.AutoFilter
ActiveSheet.UsedRange.Find(what:="[S] SWO Order", lookat:=xlWhole).Select
Dim fstAddForSort As String
Dim lstAddForSort As String

fstAddForSort = ActiveCell.Address
lstAddForSort = ActiveCell.End(xlDown).Address

ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
        fstAddForSort & ":" & lstAddForSort), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.UsedRange.Find(what:="[C] Contract Material Line Item", lookat:=xlWhole).Select
    Dim filterColumn As Integer
    Dim filterColumnAdd2 As String
    filterColumn = ActiveCell.Column
    filterColumnAdd2 = ActiveCell.Address
    ActiveSheet.UsedRange.AutoFilter Field:=filterColumn, Criteria1:="#"
   
    ActiveSheet.UsedRange.Find(what:="{S] SWO Activity Type", lookat:=xlWhole).Select
    Dim copyCellVal As String
    
    ActiveCell.EntireRow.Hidden = True
    ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible).Select
    copyCellVal = ActiveCell.Address(False, False)
    Range(filterColumnAdd2).Select
    ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible).Select
    ActiveCell.Formula = "=" & copyCellVal
    ActiveCell.Copy
    ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible).Select
    Selection.PasteSpecial xlPasteFormulas
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    Range(ActiveCell.Offset(1, -1).Address, ActiveCell.Offset(0, 1).End(xlDown).Address).ClearContents
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Cells(1, 1).Select
    ActiveCell.EntireRow.Hidden = False
    
    ActiveSheet.UsedRange.Find(what:="[S] SWO Order", lookat:=xlWhole).Select
    ActiveCell.EntireColumn.Insert xlToRight
    ActiveCell.value = "Parts/Non-Parts"
    ActiveCell.Offset(1, 0).Select
    Dim formulaTxt As String
    formulaTxt = "IF(AND(OR(" & ActiveCell.Offset(0, 1).Address(False, False) & "<>" & ActiveCell.Offset(-1, 1).Address(False, False) & "," & _
                        ActiveCell.Offset(-1, 0).Address(False, False) & "<>" & Chr(34) & "Parts" & Chr(34) & ")" & _
                        ",COUNT(FIND({" & Chr(34) & "A" & Chr(34) & "," & Chr(34) & "B" & Chr(34) & "," & Chr(34) & "C" & Chr(34) & "," & Chr(34) & "D" & Chr(34) & "," & Chr(34) & "E" & Chr(34) & "," & Chr(34) & "F" & Chr(34) & "," & Chr(34) & "G" & Chr(34) & "," & Chr(34) & "H" & Chr(34) & "," & Chr(34) & "I" & Chr(34) & "," & Chr(34) & "J" & Chr(34) & "," & Chr(34) & "K" & Chr(34) & "," & Chr(34) & "L" & Chr(34) & "," & Chr(34) & "M" & Chr(34) & "," & Chr(34) & "N" & Chr(34) & "," & Chr(34) & "O" & Chr(34) & "," & Chr(34) & "P" & Chr(34) & "," & Chr(34) & "Q" & Chr(34) & "," & Chr(34) & "R" & Chr(34) & "," & Chr(34) & "S" & Chr(34) & "," & Chr(34) & "T" & Chr(34) & "," & Chr(34) & "U" & Chr(34) & "," & Chr(34) & "V" & Chr(34) & "," & Chr(34) & "W" & Chr(34) & "," & Chr(34) & "X" & Chr(34) & "," & Chr(34) & "Y" & Chr(34) & "," & Chr(34) & "Z" & Chr(34) & "," & Chr(34) & "#" & Chr(34) & "}," & _
                        ActiveCell.Offset(0, -2).Address(False, False) & _
                        "))>0)," & Chr(34) & "Non-Parts" & Chr(34) & "," & Chr(34) & "Parts" & Chr(34) & ")"
                        
    ActiveCell.Formula = "=" & formulaTxt
    '=IF(AND(OR(Z2<>Z1,Y1<>"Parts"),COUNT(FIND({"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","#"},V2))>0),"Non-Parts","Parts")
    ActiveCell.Copy
    Range(ActiveCell.Offset(1, -1).Address, ActiveCell.Offset(1, -1).End(xlDown).Address).Select
    Selection.Offset(0, 1).Select
    Selection.PasteSpecial xlPasteFormulas
    
    ActiveCell.End(xlUp).Select
    ActiveCell.EntireColumn.Copy
    ActiveCell.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    ActiveWorkbook.Sheets("Data").Activate
        ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
        ActiveSheet.Cells(1, 3).Select
        ActiveCell.EntireColumn.Delete
        ActiveSheet.Cells(1, 4).Select
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.Offset(0, 1).Select
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.Offset(0, 1).Select
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.Offset(0, 5).Select
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.Offset(0, 3).Select
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete
        ActiveCell.EntireColumn.Delete

ActiveSheet.UsedRange.Select

Set wsData = Worksheets("Data")

'A Pivot Cache represents the memory cache for a PivotTable report. Each Pivot Table report has one cache only. Create a new PivotTable cache, and then create a new PivotTable report based on the cache.

'determine source data range (dynamic):
'last row in column no. 1:
lastRow = wsData.Cells(rows.Count, 1).End(xlUp).Row
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

    ActiveSheet.PivotTables("marketPivotTable").AddDataField ActiveSheet. _
        PivotTables("marketPivotTable").PivotFields( _
        "      swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Count of       swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo labour cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("marketPivotTable").AddDataField ActiveSheet. _
        PivotTables("marketPivotTable").PivotFields( _
        "      swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Count of       swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo material cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("marketPivotTable").AddDataField ActiveSheet. _
        PivotTables("marketPivotTable").PivotFields( _
        "      swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Count of       swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo travel cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("marketPivotTable").AddDataField ActiveSheet. _
        PivotTables("marketPivotTable").PivotFields( _
        "      other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields( _
        "Count of       other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       other swo cost"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("marketPivotTable").PivotFields("Parts/Non-Parts")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("marketPivotTable").PivotFields("Parts/Non-Parts"). _
        ClearAllFilters
    ActiveSheet.PivotTables("marketPivotTable").PivotFields("Parts/Non-Parts"). _
        CurrentPage = "Parts"
    pvtTbl.ManualUpdate = False

ActiveSheet.PivotTables("marketPivotTable").PivotSelect "", xlDataAndLabel, _
        True
    Selection.Copy
    Range("G38").Select
    ActiveSheet.Paste
    Dim pvtname As String
    pvtname = ActiveCell.PivotTable.name
    
    ActiveSheet.PivotTables(pvtname).PivotFields("Parts/Non-Parts"). _
        ClearAllFilters
    ActiveSheet.PivotTables(pvtname).PivotFields("Parts/Non-Parts"). _
        CurrentPage = "Non-Parts"
        
    ActiveSheet.PivotTables("marketPivotTable").PivotFields("[S] SWO Order"). _
        AutoSort xlDescending, "Sum of       swo material cost"
    ActiveSheet.PivotTables(pvtname).PivotFields("[S] SWO Order").AutoSort _
        xlDescending, "Sum of       swo material cost"
        
        Range("A25").Select
    ActiveCell.FormulaR1C1 = "Parts Labour Cost"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "Parts Material Cost"
    Range("A27").Select
    ActiveCell.FormulaR1C1 = "Parts Travel Cost"
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Parts Other Cost"
    Range("A29").Select
    ActiveCell.FormulaR1C1 = "Non-Parts Labpur Cost"
    Range("A30").Select
    ActiveCell.FormulaR1C1 = "Non-Parts Material Cost"
    Range("A31").Select
    ActiveCell.FormulaR1C1 = "Non-Parts Travel Cost"
    Range("A32").Select
    ActiveCell.FormulaR1C1 = "Non-Parts Other Cost"
    
    ActiveSheet.Cells(41, 1).Select
    Dim fstChartAdd As String
    Dim lstChartAdd As String
    
    fstChartAdd = ActiveCell.Offset(1, 0).Address(True, False)
    lstChartAdd = ActiveCell.End(xlDown).Address
    
    ActiveSheet.Cells(25, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 1).Address(False, False) & ")"
    ActiveSheet.Cells(26, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 2).Address(False, False) & ")"
    ActiveSheet.Cells(27, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 3).Address(False, False) & ")"
    ActiveSheet.Cells(28, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 4).Address(False, False) & ")"
    ActiveSheet.Cells(29, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 7).Address(False, False) & ")"
    ActiveSheet.Cells(30, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 8).Address(False, False) & ")"
    ActiveSheet.Cells(31, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 9).Address(False, False) & ")"
    ActiveSheet.Cells(32, 2).Formula = "=SUM(" & Range(fstChartAdd, lstChartAdd).Offset(0, 10).Address(False, False) & ")"
    
    
    Range("A25:B32").Select
    ActiveSheet.Shapes.AddChart2(253, xlPie).Select
    ActiveChart.SetSourceData Source:=Range("Pivot!$A$25:$B$32")
    With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 350  ' resize
         .Top = 10    ' reposition
         .Left = 300   ' reposition
     End With
     
     
     Range("B40").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "System Code (6NC)").Slicers.Add ActiveSheet, , "System Code (6NC) 1", _
        "System Code (6NC)", 5, 5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "Market").Slicers.Add ActiveSheet, , "Market 1", "Market", 210, 5, 144 _
        , 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "Country A").Slicers.Add ActiveSheet, , "Country A 1", "Country A", 210, _
        150, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("marketPivotTable"), _
        "Fiscal Year/Period").Slicers.Add ActiveSheet, , "Fiscal Year/Period 1", _
        "Fiscal Year/Period", 5, 150, 144, 198.75
    ActiveSheet.Shapes.Range(Array("Fiscal Year/Period 1")).Select
    ActiveSheet.Shapes.Range(Array("System Code (6NC) 1")).Select
    ActiveSheet.Shapes.Range(Array("Market 1")).Select
    ActiveSheet.Shapes.Range(Array("Country A 1")).Select
    ActiveSheet.Shapes.Range(Array("Fiscal Year/Period 1")).Select
    
    ActiveSheet.name = "SWO_Cost_PieChart"
    
    Sheets.Add
ActiveSheet.name = "Pivot"

Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R40C1", TableName:="paratoPivotTable", DefaultVersion:=xlPivotTableVersion15)

'change style of the new PivotTable:
pvtTbl.TableStyle2 = "PivotStyleMedium3"

'to view the PivotTable in Classic Pivot Table Layout, set InGridDropZones property to True, else set to False:
pvtTbl.InGridDropZones = True

'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.
pvtTbl.ManualUpdate = True

With ActiveSheet.PivotTables("paratoPivotTable").PivotFields("[S] SWO Order")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("paratoPivotTable").PivotFields("[S] SWO Order"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("paratoPivotTable").PivotFields("[S] SWO Order"). _
        LayoutForm = xlTabular
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields( _
        "[C] Contract Material Line Item")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields("Parts/Non-Parts")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("paratoPivotTable").AddDataField ActiveSheet. _
        PivotTables("paratoPivotTable").PivotFields( _
        "      swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields( _
        "Count of       swo labour cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo labour cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("paratoPivotTable").AddDataField ActiveSheet. _
        PivotTables("paratoPivotTable").PivotFields( _
        "      swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields( _
        "Count of       swo material cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo material cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("paratoPivotTable").AddDataField ActiveSheet. _
        PivotTables("paratoPivotTable").PivotFields( _
        "      swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields( _
        "Count of       swo travel cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       swo travel cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("paratoPivotTable").AddDataField ActiveSheet. _
        PivotTables("paratoPivotTable").PivotFields( _
        "      other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract"), _
        "Count of       other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract", xlCount
    With ActiveSheet.PivotTables("paratoPivotTable").PivotFields( _
        "Count of       other swo cost" & Chr(10) & "settled to" & Chr(10) & "contract")
        .Caption = "Sum of       other swo cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("paratoPivotTable").PivotFields("Parts/Non-Parts"). _
        ClearAllFilters
    ActiveSheet.PivotTables("paratoPivotTable").PivotFields("Parts/Non-Parts"). _
        CurrentPage = "Parts"
     
    pvtTbl.ManualUpdate = False
   
    ActiveSheet.UsedRange.Find(what:="Row Labels", lookat:=xlWhole).Select
    Range(ActiveCell.Address, ActiveCell.SpecialCells(xlCellTypeLastCell).Offset(-1, 0).Address).Copy
    Sheets.Add
    ActiveSheet.Cells(41, 1).Select
    ActiveSheet.Paste
    
    ThisWorkbook.Sheets("UI").Activate
    Workbooks(inputFileNameContracts).Close False
    Workbooks(revenueOutputGlobal).Save
    
End Sub

