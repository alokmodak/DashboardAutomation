Attribute VB_Name = "DiffusionRate"
Public Sub DiffusionRate_Calculations()

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

'Deleting # lines
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Start Date (Header)", lookat:=xlWhole).Select

ActiveCell.AutoFilter field:=ActiveCell.Column, Criteria1:="#", Operator:=xlFilterValues
ActiveCell.EntireRow.Hidden = True
ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Select
Selection.EntireRow.Delete
ActiveSheet.UsedRange.EntireRow.Hidden = False

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
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
lstZCSWAdd = ActiveCell.Address

Range(fstZCSWAdd, lstZCSWAdd).Copy
Sheets.Add
ActiveSheet.name = "Contracts-Data"
ActiveCell.PasteSpecial xlPasteAll

Dim eqRNG As String
Dim celVal As Variant

Range(ActiveCell.Address, Cells(Rows.Count, ActiveCell.Column).End(xlUp).Address).Select
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

ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
ActiveCell.Offset(0, 2).Value = "IB Year"
ActiveCell.Offset(1, 0).Select

ActiveCell.Offset(0, 2).Formula = "=RIGHT(" & ActiveCell.Offset(0, 1).Address(False, False) & ",4)"
Dim fstCopyVal As String
Dim lstCopyVal As String
fstCopyVal = ActiveCell.Offset(0, 2).Address
ActiveCell.End(xlDown).Select
lstCopyVal = ActiveCell.Offset(0, 2).Address

Range(fstCopyVal).Select
ActiveCell.Copy

Range(fstCopyVal, lstCopyVal).PasteSpecial xlPasteFormulas
Selection.Copy
Selection.PasteSpecial xlPasteValues

Dim fstFilterVal1 As String
Dim lstFilterVal1 As String


ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
fstFilterVal1 = ActiveCell.Address
lstFilterVal1 = ActiveCell.End(xlDown).Address
ActiveWorkbook.Sheets("Pivot").Delete

ActiveWorkbook.Sheets("Contracts-Data").Activate
Range(fstFilterVal1, lstFilterVal1).Select

Dim filterRNG1 As String
Dim n As Long
filterRNG1 = Selection.Address
i = 1
n = Range(filterRNG1).Count
ReDim aryData(1 To n) As String
For Each cell In Range(filterRNG1).Cells
    aryData(i) = cell
    i = i + 1
Next

ActiveWorkbook.Sheets("Data").Activate
ActiveSheet.Cells(1, 1).Select

Dim fstFilterAdd As String
Dim lstFilterAdd As String
Dim filterRNG As String


fstFilterAdd = ActiveCell.Address
ActiveCell.SpecialCells(xlCellTypeLastCell).Select
lstFilterAdd = ActiveCell.Address
filterRNG = Range(fstFilterAdd, lstFilterAdd).Address

ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select

'Filter range from the data
Range(filterRNG).AutoFilter field:=4, Criteria1:=aryData, Operator:=xlFilterValues

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

Do Until ActiveCell.Value = ""
    If InStr(1, ActiveCell.Value, ".", vbTextCompare) Then
        ActiveCell.Value = Replace(ActiveCell.Value, ".", "")
        ActiveCell.Value = ActiveCell.Value & "0"
    Else
        ActiveCell.Value = ActiveCell.Value
    End If
    ActiveCell.Offset(1, 0).Select
Loop

ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
ActiveCell.EntireColumn.Insert Shift:=xlToRight
ActiveCell.Value = "IB Year"
ActiveCell.Offset(1, 0).Select
ActiveCell.Formula = "=VLOOKUP(" & ActiveCell.Offset(0, 1).Address(False, False) & ",'Contracts-Data'!" & Sheets("Contracts-Data").UsedRange.Address & ",3,FALSE)"

Dim fstAddToCopyIB As String
Dim lstAddToCopyIB As String

ActiveCell.Copy
ActiveCell.Offset(0, 1).Select
fstAddToCopyIB = ActiveCell.Offset(0, -1).Address
ActiveCell.End(xlDown).Select
lstAddToCopyIB = ActiveCell.Offset(0, -1).Address

Range(fstAddToCopyIB, lstAddToCopyIB).Select
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
Application.ReferenceStyle = xlR1C1
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceData1, Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Revenue!R30C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15

With ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Year/Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("    Total Contract Revenue"), _
        "Count of     Total Contract Revenue", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Count of     Total Contract Revenue")
        .Caption = "Sum of     Total Contract Revenue"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("System Code (6NC)")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("IB Year")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = False
        .RowGrand = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Year/Period"). _
        ShowAllItems = True
Application.ReferenceStyle = xlA1
        
        ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    
    pvtTblName = ActiveCell.PivotTable.name
    Set pvtTbl = ActiveCell.PivotTable
    
    For Each pvtFld In pvtTbl.PivotFields
        If pvtFld.Caption <> "IB Year" Then
            pvtFld.Orientation = xlHidden
        End If
    Next
        
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Sum of     Total Contract Revenue").Orientation = xlHidden
    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
        pvtTblName).PivotFields("[C,S] Reference Equipment"), _
        "Sum of [C,S] Reference Equipment", xlSum
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Sum of [C,S] Reference Equipment")
        .Caption = "Count of [C,S] Reference Equipment"
        .Function = xlCount
    End With
    
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    Range("A36").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(pvtTblName), _
        "IB Year").Slicers.Add ActiveSheet, , "IB Year", "IB Year", 289.5, 334.5, 144, _
        198.75
    ActiveSheet.Shapes.Range(Array("IB Year")).Select
    ActiveSheet.Shapes("IB Year").IncrementLeft -297
    ActiveSheet.Shapes("IB Year").IncrementTop -141
    ActiveWorkbook.SlicerCaches("Slicer_IB_Year").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
    
    Dim productAdd As String
    ActiveSheet.UsedRange.Find(what:="Sum of     Total Contract Revenue", lookat:=xlWhole).Select
    productAdd = ActiveCell.Offset(2, 0).Value
    
    ActiveSheet.Cells(1, 1).Value = ""
    ActiveSheet.Cells(12, 3).Value = productAdd
    ActiveSheet.Cells(13, 3).Value = "IB Count"
    
    Dim ibCountRng As String
    ActiveSheet.UsedRange.Find(what:="Count of [C,S] Reference Equipment", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    
    ibCountRng = Range(ActiveCell.Address, ActiveCell.End(xlDown).Address).Address
    
    ActiveSheet.Cells(14, 3).Formula = "=COUNT(" & ibCountRng & ")"
    ActiveSheet.Cells(14, 4).Formula = "=" & Cells(14, 3).Address & "*55000"
    
    ActiveSheet.Cells(16, 3).Value = "Diffusion Rate"
    ActiveSheet.UsedRange.Find(what:="Sum of     Total Contract Revenue", lookat:=xlWhole).Select
    Dim yrAdd As String
    Dim divAdd As String
    yrAdd = ActiveCell.Offset(1, 1).Address(False, False)
    divAdd = ActiveCell.Offset(2, 1).Address(False, False)
    
    ActiveSheet.Cells(17, 3).Formula = "=" & yrAdd
    ActiveSheet.Cells(17, 3).Select
    ActiveCell.Copy
    Do Until ActiveCell.Offset(15, 0).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    ActiveCell.End(xlToLeft).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Formula = "=" & divAdd & "/$D$14"
    ActiveCell.Copy
    
    Do Until ActiveCell.Offset(-1, 0).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    Range(ActiveCell.Address, ActiveCell.End(xlToLeft).Address).NumberFormat = "0%"
    
    Sheets("Revenue").Select
    Sheets("Revenue").Copy Before:=Sheets(2)
    Sheets("Revenue (2)").Select
    Sheets("Revenue (2)").name = "Cost"
    Range("A30").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Sum of     Total Contract Revenue").Orientation = xlHidden
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total SWO cost"), "Count of Total SWO cost", _
        xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Count of Total SWO cost")
        .Caption = "Sum of Total SWO cost"
        .Function = xlSum
    End With
    ActiveSheet.UsedRange.Find(what:="Count of [C,S] Reference Equipment", lookat:=xlWhole).Select
    pvtTblName = ActiveCell.PivotTable.name
    Dim fstNewAdd As String
    fstNewAdd = ActiveCell.Address
    
    ActiveSheet.PivotTables(pvtTblName).AddDataField ActiveSheet.PivotTables( _
        pvtTblName).PivotFields("Total SWO cost"), "Count of Total SWO cost", _
        xlCount
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Count of Total SWO cost")
        .Caption = "Sum of Total SWO cost"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Count of [C,S] Reference Equipment").Orientation = xlHidden
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Fiscal Year/Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    ActiveSheet.Cells(16, 3).Value = "Bath Tub"
    ActiveCell.Offset(18, 3).Select
    Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address).NumberFormat = "0"
    With Range(ActiveCell.Address, ActiveCell.End(xlToRight).Address).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range(fstNewAdd).Select
    Dim fstCostIbCountAdd As String
    Dim lstCostIbCountAdd As String
    
    fstCostIbCountAdd = ActiveCell.Offset(1, 0).Address(False, False)
    ActiveCell.End(xlToLeft).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstCostIbCountAdd = ActiveCell.Address(False, False)
    
    Cells(14, 3).Select
    ActiveCell.Formula = "=IFERROR(Count(" & fstCostIbCountAdd & ":" & lstCostIbCountAdd & "),)"
    ActiveCell.Copy
    Do Until ActiveCell.Offset(3, 0).Value = ""
        ActiveCell.Offset(0, 1).Select
        ActiveCell.PasteSpecial xlPasteFormulas
    Loop
    
    ActiveSheet.UsedRange.Find(what:="Sum of Total SWO cost", lookat:=xlWhole).Select
    Dim fstCostAdd As String
    fstCostAdd = ActiveCell.Offset(2, 1).Address(False, False)
    ActiveSheet.Cells(18, 3).Select
    ActiveCell.Formula = "=IFERROR(" & fstCostAdd & "/" & ActiveCell.Offset(-4, 0).Address(False, False) & ",)"
    ActiveCell.Copy
    
    Do Until ActiveCell.Offset(-1, 1).Value = ""
     ActiveCell.Offset(0, 1).Select
     ActiveCell.PasteSpecial xlPasteFormulas
    Loop
        
ActiveWorkbook.Sheets("Data").Delete

End Sub
