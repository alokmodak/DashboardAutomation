Attribute VB_Name = "COntractsJoinDropEmailSending"
Public Sub Send_Email_For_Contracts_Drops_Joins()

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
Dim userRequiremetsFile As String
Dim FSO As FileSystemObject
Dim exportFolderPath As String
Dim exportFolder As String

On Error Resume Next

Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
If Application.FileDialog(msoFileDialogFilePicker).Show <> -1 Then
MsgBox "No File is Selected!"
End
End If

inputRevenue = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = ActiveWorkbook.name

Set FSO = New scripting.FileSystemObject
exportFolderPath = FSO.GetParentFolderName(inputRevenue)
exportFolder = exportFolderPath & "\ExportedFiles\"

If Not FSO.FolderExists(exportFolder) Then
    FSO.CreateFolder (exportFolder)
End If

'Copy Data from SAP file
strtMonth = Format(Now() - 31, "mmmyyyy")
marketInputFile = "Market_Groups_Markets_Country.xlsx"
marketInputFile = Replace(inputRevenue, inputFileNameContracts, marketInputFile)
Application.Workbooks.Open (marketInputFile), False

'opening Requrements file
'userRequiremetsFile = "User_Requirements.xlsx"
'userRequiremetsFile = Replace(inputRevenue, inputFileNameContracts, userRequiremetsFile)
'Application.Workbooks.Open (userRequiremetsFile), False
Dim reqFile As String
reqFile = ThisWorkbook.name

Workbooks(inputFileNameContracts).Activate
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate

revenueOutputGlobal = Left(inputRevenue, InStrRev(inputRevenue, "\") - 1) & "\" & "Contracts-Drops&Joins_" & Format(Now, "mmmyy") & ".xlsm"
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
ActiveCell.Formula = "=IF(IFERROR(VLOOKUP(" & lookForVal & "," & rngStringMarket & "," & "3" & "," & "FALSE)<" & Chr(61) & "YEAR(TODAY()),)," & Chr(34) & "Yes" & Chr(34) & "," & Chr(34) & "No" & Chr(34) & ")"
'=IF(VLOOKUP(B31,$AS$30:$AU$64,3,FALSE)<YEAR(TODAY()),"EOL","Not EOL")
ActiveCell.Copy
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).PasteSpecial xlPasteAll
ActiveSheet.Range(fstPasteRNG, lstPasteRNG).Select
Selection.Copy
Selection.PasteSpecial (xlValues)
marketRNG.Delete

Calculating_Data_Downloaded_Date
ActiveCell.EntireColumn.Delete

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
Set pvtTbl = PvtTblCache.CreatePivotTable(TableDestination:="Pivot!R3C1", TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)

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
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 1
    End With
    
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
        "[C,S] Ship-To Party Line Item")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Ship-To Party Line Item").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Ship-To Party Line Item A")
        .Orientation = xlRowField
        .Position = 3
    End With
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Ship-To Party Line Item A").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Ship-to City")
        .Orientation = xlRowField
        .Position = 4
    End With
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields("Ship-to City").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "    Contract" & Chr(10) & "Net Value")
        .Orientation = xlRowField
        .Position = 5
    End With
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields("    Contract" & Chr(10) & "Net Value"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
 
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "EOL Status")
        .Orientation = xlRowField
        .Position = 6
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "EOL Status").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "System Code (6NC)")
        .Orientation = xlRowField
        .Position = 7
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "System Code (6NC)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlRowField
        .Position = 8
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Market").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Market")
        .Orientation = xlRowField
        .Position = 9
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(pvtTblName).PivotFields("Country A")
        .Orientation = xlRowField
        .Position = 10
    End With
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "Country A").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .Orientation = xlRowField
        .Position = 11
    End With
    With ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 12
    End With
    
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
    
    ActiveSheet.PivotTables(pvtTblName).PivotFields( _
        "[C,S] Contract End Date (Header)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type")
        .Orientation = xlRowField
        .Position = 13
    End With
    
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

'Starting the process to filter data and collect drops and joins
Dim cellVal

Dim fstRNGForSendMail As String
Dim lstRNGForSendMail As String
Dim filterVal As Variant
Dim filterValCountry As String
Dim filterValMarket As String
Dim filterNum As Integer
Dim countryFilterVal As Variant
Dim marketFilterVal As Variant
Dim dateVal As String
Dim toEmailAdd As String
Dim found As Boolean
Dim subject As String
Dim txtBody As String

Application.Workbooks(reqFile).Activate
ActiveWorkbook.Sheets("Contracts_Drop-Requirements").Activate
ActiveSheet.UsedRange.Find(what:="Market", lookat:=xlWhole).Select
fstRNGForSendMail = ActiveCell.Offset(1, 0).Address
lstRNGForSendMail = ActiveCell.End(xlDown).Address

For Each cellVal In Range(fstRNGForSendMail, lstRNGForSendMail)
Application.Workbooks(reqFile).Activate
ActiveWorkbook.Sheets("Contracts_Drop-Requirements").Activate
ActiveCell.Offset(1, 0).Select
filterValMarket = ActiveCell.Value
filterValCountry = ActiveCell.Offset(0, 1).Value
dateVal = ActiveCell.Offset(0, 2)
toEmailAdd = ActiveCell.Offset(0, 3)
subject = ActiveSheet.Cells(6, 9).Value
txtBody = ActiveSheet.Cells(7, 9).Value


Workbooks(revenueOutputGlobal).Activate
ActiveSheet.PivotTables(pvtTblName).PivotFields("Country A").ClearAllFilters

ActiveSheet.PivotTables(pvtTblName).PivotFields("Market").ClearAllFilters
            
    countryFilterVal = Split(filterValCountry, ";")
        
        Workbooks(revenueOutputGlobal).Activate
        For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields("Country A").PivotItems
            With pvtItem
                found = False
                For filterNum = 0 To UBound(countryFilterVal)
                    If pvtItem = countryFilterVal(filterNum) Then found = True
                Next
                If Not found Then
                    pvtItem.Visible = False
                End If
            End With
        Next
        If filterValCountry = "" Then ActiveSheet.PivotTables(pvtTblName).PivotFields("Country A").ClearAllFilters
        
    marketFilterVal = Split(filterValMarket, ";")
    Workbooks(revenueOutputGlobal).Activate
        For Each pvtItem In ActiveSheet.PivotTables(pvtTblName).PivotFields("Market").PivotItems
            With pvtItem
                found = False
                For filterNum = 0 To UBound(marketFilterVal)
                    If pvtItem = marketFilterVal(filterNum) Then found = True
                Next
                If Not found Then
                    pvtItem.Visible = False
                End If
            End With
        Next
        If filterValMarket = "" Then ActiveSheet.PivotTables(pvtTblName).PivotFields("Market").ClearAllFilters

'Copy Pivot table values to new sheet
ActiveWorkbook.Sheets("Pivot").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.SpecialCells(xlCellTypeLastCell).Select
ActiveCell.End(xlUp).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

ActiveWorkbook.Sheets.Add
With ActiveSheet.Cells(2, 1)
    .PasteSpecial xlPasteValues
End With

ActiveSheet.name = "Contracts-Joins&Drops"

ActiveWorkbook.Sheets("Contracts-Joins&Drops").Activate

ActiveSheet.Cells(2, 1).Select
Dim fstTableAdd As String
fstTableAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select

monthsForTable = DateAdd("m", -1, dateVal)

ActiveCell.Offset(0, 1).Select
For monthCellForTable = 2 To 3
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
            If ActiveCell.Offset(0, 10).Value = "" Then
                ActiveCell.Offset(0, 10).Value = ActiveCell.Offset(-1, 10).Value
            End If
            If ActiveCell.Offset(0, 11).Value = "" Then
                ActiveCell.Offset(0, 11).Value = ActiveCell.Offset(-2, 10).Value
            End If
            duration = DateDiff("m", Replace(ActiveCell.Offset(0, 10).Value, ".", "/"), Replace(ActiveCell.Offset(0, 11).Value, ".", "/"))
            i = 1
            Do Until ActiveCell.Offset(i, 0).Value <> "" Or i > 20
            'exit loop for last cell
                If ActiveCell.Offset(i, 12).Value = "" Then
                Exit Do
                End If
            If ActiveCell.Offset(i, 10).Value = "" Then
                ActiveCell.Offset(i, 10).Value = ActiveCell.Offset(-1, 10).Value
            End If
            If ActiveCell.Offset(i, 11).Value = "" Then
                ActiveCell.Offset(i, 11).Value = ActiveCell.Offset(-2, 11).Value
            End If
            duration = duration + DateDiff("m", Replace(ActiveCell.Offset(i, 10).Value, ".", "/"), Replace(ActiveCell.Offset(i, 11).Value, ".", "/"))
            i = i + 1
            Loop
        
            monthCellForTable = 13
            For i = 1 To 2
            
        Dim k As Integer
        k = 0
        Do
        'exit for last cell
        If ActiveCell.Offset(k, 12).Value = "" Then
            Exit Do
        End If
                fstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 10).Value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 10).Value, ".", "/", 4)), 1)
                lstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 11).Value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 11).Value, ".", "/", 4)) + 1, 0)
                
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
            If ActiveCell.Offset(0, 5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "EOL"
            End If
            
        'condition for After warranty
        If ActiveCell.Offset(0, 12).Value = "ZCSW" Then
        j = 1
        zcswVal = True
        Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
        'condition for last row exit loop
            If ActiveCell.Offset(j, 12).Value <> "ZCSW" Then
                If ActiveCell.Offset(1, 12).Value = "" Then
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
            If ActiveCell.Offset(0, 5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 2).Value = "EOL"
            End If
            
            If ActiveCell.Offset(0, 12).Value = "ZCSW" Then
            j = 1
            zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
            'condition for last row exit loop
                If ActiveCell.Offset(j, 12).Value <> "ZCSW" Then
                    If ActiveCell.Offset(1, 12).Value = "" Then
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
   If ActiveCell.Offset(0, 5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 1).Value = "EOL"
            End If
            
'condition for After warranty
If ActiveCell.Offset(0, 12).Value = "ZCSW" Then
   j = 1
zcswVal = True
            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                    'condition for last row exit loop
                    If ActiveCell.Offset(j, 12).Value <> "ZCSW" Then
                        If ActiveCell.Offset(1, 12).Value = "" Then
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
                    If ActiveCell.Offset(0, 5).Value = "EOL" Then
                ActiveCell.Offset(0, monthCellForTable - 1).Value = "EOL"
            End If
            
                    'condition for After warranty
                    If ActiveCell.Offset(0, 12).Value = "ZCSW" Then
                        j = 1
                        zcswVal = True
                            Do Until ActiveCell.Offset(j, 0) <> "" Or j > 20
                                'condition for last row exit loop
                                If ActiveCell.Offset(j, 12).Value <> "ZCSW" Then
                                    If ActiveCell.Offset(1, 12).Value = "" Then
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
    
ActiveSheet.Cells(1, 1).Select

Application.DisplayAlerts = False
    Sheets("Contracts-Joins&Drops").Select
    Sheets("Contracts-Joins&Drops").Move
    Dim flName As String
    flName = exportFolder & filterValCountry & "_" & filterValMarket & "-Drops.xlsx"
    ActiveWorkbook.SaveAs fileName:=flName, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
Application.DisplayAlerts = True

ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", lookat:=xlWhole).Select
ActiveCell.Offset(0, 1).EntireColumn.Delete
ActiveCell.Offset(0, 2).EntireColumn.Delete
ActiveCell.Offset(0, 2).EntireColumn.Delete
ActiveCell.Offset(0, 3).EntireColumn.Delete
ActiveCell.Offset(0, 3).EntireColumn.Delete
ActiveCell.Offset(0, 2).EntireColumn.Delete

Do Until ActiveCell.Value = ""
    If ActiveCell.Offset(0, 1).Value = "" Then
        ActiveCell.EntireRow.Delete
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Loop

Dim cutAdd As String
Dim pasteAdd As String

ActiveSheet.UsedRange.Select
Selection.Columns.AutoFit
ActiveSheet.UsedRange.Find(what:="[C,S] Reference Equipment", lookat:=xlWhole).Select
ActiveCell.Value = "Equipment ID"
ActiveSheet.UsedRange.Find(what:="[C,S] Ship-To Party Line Item", lookat:=xlWhole).Select
ActiveCell.Value = "Customer ID"
ActiveSheet.UsedRange.Find(what:="[C,S] Ship-To Party Line Item A", lookat:=xlWhole).Select
ActiveCell.Value = "Customer Name"
ActiveSheet.UsedRange.Find(what:="Ship-to City", lookat:=xlWhole).Select
ActiveCell.Value = "Location"
ActiveSheet.UsedRange.Find(what:="EOL Status", lookat:=xlWhole).Select
ActiveCell.Value = "EOL ?"
ActiveSheet.UsedRange.Find(what:="    Contract" & Chr(10) & "Net Value", lookat:=xlWhole).Select
ActiveCell.Value = "Contract Value"
cutAdd = ActiveCell.Address
ActiveSheet.UsedRange.Find(what:="System Code (6NC)", lookat:=xlWhole).Select
ActiveCell.Value = "System Name"
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", lookat:=xlWhole).Select
ActiveCell.Value = "System 6NC"
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Start Date (Header)", lookat:=xlWhole).Select
ActiveCell.Value = "Contract Start"
ActiveSheet.UsedRange.Find(what:="[C,S] Contract End Date (Header)", lookat:=xlWhole).Select
ActiveCell.Value = "Contract End"
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", lookat:=xlWhole).Select
ActiveCell.Value = "Contract Type"
pasteAdd = ActiveCell.Address
    Range(cutAdd).Select
    Range(ActiveCell.Address & ":" & ActiveCell.End(xlDown).Address).Select
    Selection.Cut
    Range(pasteAdd).Select
    Selection.Insert Shift:=xlToLeft

ActiveSheet.UsedRange.Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveSheet.UsedRange.Find(what:="Country A", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

ActiveSheet.Cells(1, 1).Select
ActiveWorkbook.Save

Send_Email_Via_OutlookInbox flName, toEmailAdd, subject, txtBody

Next cellVal

ThisWorkbook.Sheets("UI").Activate
Application.Workbooks(marketInputFile).Close False
'Application.Workbooks(reqFile).Close False
Application.Workbooks(revenueOutputGlobal).Close False

End Sub
