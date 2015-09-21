Attribute VB_Name = "DiffusionRate"
Public Sub DiffusionRate()

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

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "ContractsDiffusion_Rate_" & Format(Now, "mmmyy") & ".xlsm"
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
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
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
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("D5").Select
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type")
        For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type").PivotItems
            If pvtItem.name = "ZCSW" Then
            .PivotItems("ZCSW").Visible = True
            Else
            pvtItem.Visible = False
            End If
        Next
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields("[C,S] Contract Type"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").ClearAllFilters
    With ActiveSheet.PivotTables(pvtTblName)
        .ColumnGrand = False
        .RowGrand = False
    End With
    
Dim filterSelectedValues As Integer
    Dim secondLoop As Integer
    secondLoop = 1
    
    ActiveWorkbook.Sheets("Pivot").Activate
'turn on automatic update / calculation in the Pivot Table
pvtTbl.ManualUpdate = False
Set PvtTblCache = Nothing

ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
Dim fstZCSWAdd As String
Dim lstZCSWAdd As String
fstZCSWAdd = ActiveCell.Address
Range(Mid(ActiveCell.Address, 2, 2) & Rows.Count).End(xlUp).Select
lstZCSWAdd = ActiveCell.Address

Range(fstZCSWAdd, lstZCSWAdd).Copy
Sheets.Add
ActiveSheet.name = "Contracts-Data"
ActiveCell.PasteSpecial xlPasteAll

Dim eqRNG As String
Dim celVal As Variant

eqRNG = Selection.Address
For Each celVal In Range(eqRNG)
    If ActiveCell.Offset(1, 0).Value = "" Then
        ActiveCell.EntireRow.Delete
    End If
    If ActiveCell.Value = "" Or ActiveCell.Value = "Grand Total" Then
        ActiveCell.EntireRow.Delete
        ActiveCell.Offset(-1, 0).Select
    End If
    ActiveCell.Offset(1, 0).Select
Next

Dim fstFilterVal1 As String
Dim lstFilterVal1 As String

ActiveSheet.UsedRange.Copy

fstFilterVal1 = ActiveCell.End(xlUp).Address
lstFilterVal1 = ActiveCell.Address
'ActiveWorkbook.Sheets("Pivot").Delete

ActiveWorkbook.Sheets("Contracts-Data").Activate
ActiveSheet.UsedRange.Copy

ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.Cells(1, 1).Select

Dim fstFilterAdd As String
Dim lstFilterAdd As String
Dim filterRNG As String

fstFilterAdd = ActiveCell.Address
ActiveCell.SpecialCells(xlCellTypeLastCell).Select
lstFilterAdd = ActiveCell.Address
filterRNG = Range(fstFilterAdd, lstFilterAdd).Address

ActiveCell.Offset(10, 10).Select
ActiveCell.PasteSpecial xlPasteAll
Dim filterRNG1 As String
filterRNG1 = Selection.Address

Range(filterRNG).AdvancedFilter Action:=xlFilterInPlace, CriteriaRange _
        :=Range(filterRNG1), Unique:=False

Range(filterRNG1).Delete
ActiveCell.SpecialCells(xlCellTypeVisible).Select
Selection.Copy
Sheets.Add
ActiveSheet.Paste
ActiveSheet.name = "Filtered-Data"
ActiveSheet.UsedRange.Find(what:="{C,S] Fiscal Year/Period", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert Shift:=xlToRight

ActiveCell.Value = "Fiscal Year/Period"
ActiveCell.Offset(1, 0).Select
Dim fstAddForYear As String
Dim fstAddRNG2 As String
Dim lstAddRNG2 As String
fstAddForYear = ActiveCell.Offset(0, 1).Address(False, False)
fstAddRNG2 = ActiveCell.Address
ActiveCell.Formula = "=RIGHT(" & fstAddForYear & ",4)"
Range(fstAddForYear).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
lstAddRNG2 = ActiveCell.Address
Range(fstAddForYear).Select
ActiveCell.Offset(0, -1).Copy
Range(fstAddRNG2, lstAddRNG2).Select
Selection.PasteSpecial xlPasteFormulas
Selection.Copy
Selection.PasteSpecial xlPasteValues

ActiveSheet.UsedRange.Select
Dim pivoRNG As String
pivoRNG = Selection.Address
Dim sourceData1 As String


sourceData1 = "Filtered-Data!" & pivoRNG
Application.CutCopyMode = False
Sheets.Add
ActiveSheet.name = "Revenue"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Filtered-Data!R1C1:R53020C47", Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Revenue!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15
ActiveWorkbook.Sheets("Pivot").Delete
End Sub
