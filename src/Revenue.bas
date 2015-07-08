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
Dim zcswVal As Boolean

'Copy Data from SAP file
Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
Application.Workbooks.Add
ActiveWorkbook.SaveAs fileName:="D:\Philips\Assignments\Revenue\ContractDynamics_Waterfall_Jul15.xlsx", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
outputFile = ActiveWorkbook.name
inputRevenue = "D:\Philips\Assignments\Revenue\ContractDynamics_Waterfall.xlsx"
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = Split(inputRevenue, "\")(UBound(Split(inputRevenue, "\")))
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

Application.Workbooks(outputFile).Activate
ActiveWorkbook.Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Data!R1C1:R74904C23", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet4!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    ActiveSheet.name = "Pivot"
    ActiveWorkbook.Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Country")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Reference Equipment")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("A5").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Reference Equipment") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With ActiveSheet.PivotTables("PivotTable1")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Reference Equipment")
        .PivotItems("#").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract End Date (Header)")
        .Orientation = xlRowField
        .Position = 3
    End With
    Range("B6").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)").Subtotals = Array(False, False, False, False _
        , False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract Start Date (Header)")
        .PivotItems("#").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract End Date (Header)")
        .PivotItems("#").Visible = False
    End With
    Range("C7").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] Contract End Date (Header)").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("D5").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type")
        .PivotItems("#").Visible = False
        .PivotItems("MV").Visible = False
        .PivotItems("ZPO").Visible = False
        .PivotItems("ZSO").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[C,S] Contract Type"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[C,S] System Code Material (Material no of  R Eq)").CurrentPage = "718074"

'Copy Pivot table values to new sheet
ActiveSheet.UsedRange.Find(what:="[C,S] Contract Type", LookAt:=xlWhole).Select
fstAddForPivot = ActiveCell.Address
ActiveCell.End(xlDown).Select
ActiveCell.End(xlToLeft).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

ActiveWorkbook.Sheets.Add
With ActiveSheet.Cells(2, 27)
    .PasteSpecial xlPasteValues
End With
ActiveSheet.name = "Endura"

ActiveWorkbook.Sheets("Pivot").delete
ActiveWorkbook.Sheets("Data").delete

ActiveWorkbook.Sheets("Endura").Activate
ActiveSheet.Cells(2, 27).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).value = "ContractYearBand"

monthsForTable = DateAdd("m", -24, Date)

ActiveCell.Offset(0, 1).Select
For monthCellForTable = 2 To 37
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = monthsForTable
    ActiveCell.NumberFormat = "[$-409]mmm-yy;@"
    monthsForTable = DateAdd("m", 1, monthsForTable)
Next
ActiveCell.End(xlToLeft).Select

ActiveCell.Offset(1, 0).Select
fstAddForPivot = ActiveCell.Address
Range(Mid(ActiveCell.Address, 2, 2) & Rows.Count).End(xlUp).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot).Select

topCelVal = 1

'Loop for each row individually to calculate values
For Each cell In Range(fstAddForPivot, lstAddForPivot)

If ActiveCell.value <> "" Then
    duration = DateDiff("m", Replace(ActiveCell.Offset(0, 1).value, ".", "/"), Replace(ActiveCell.Offset(0, 2).value, ".", "/"))
    i = 1
    Do Until ActiveCell.Offset(i, 0).value = ""
    duration = duration + DateDiff("m", Replace(ActiveCell.Offset(i, 1).value, ".", "/"), Replace(ActiveCell.Offset(i, 2).value, ".", "/"))
    i = i + 1
    Loop
    
    'Conditions for ContractYearBand
    If duration <= 12 Then
        ActiveCell.Offset(0, 4).value = "0To1Year"
    ElseIf 13 >= duration <= 36 Then
        ActiveCell.Offset(0, 4).value = "2To3Years"
    ElseIf 37 >= duration <= 60 Then
        ActiveCell.Offset(0, 4).value = "3To5Years"
    ElseIf duration >= 61 Then
        ActiveCell.Offset(0, 4).value = "MoreThan5Years"
    End If
    
    'condition for After warranty
    If ActiveCell.Offset(0, 3).value = "ZCSW" Then
        i = 1
        zcswVal = True
            Do Until ActiveCell.Offset(i, 0) = ""
                If ActiveCell.Offset(i, 3).value <> "ZCSW" Then
                    zcswVal = False
                End If
            i = i + 1
            Loop
            If zcswVal = True Then
                ActiveCell.Offset(0, 4).value = "AfterWarranty"
            End If
    End If
    monthCellForTable = 5
    For i = 1 To 36
        If ActiveCell.Offset(0, 1).value = "" Then
            Exit For
        End If
        fstVal = DateSerial(Year(Replace(ActiveCell.Offset(0, 1).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(0, 1).value, ".", "/", 4)), 1)
        lstVal = DateSerial(Year(Replace(ActiveCell.Offset(0, 2).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(0, 2).value, ".", "/", 4)) + 1, 0)
        
        If fstVal <= CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) And CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) <= lstVal Then
            ActiveCell.Offset(0, monthCellForTable).value = "Yes"
        Else
            ActiveCell.Offset(0, monthCellForTable).value = "No"
        End If
        monthCellForTable = monthCellForTable + 1
    Next

Else

   'Loop for Yes No values with skip from ContractYearBand
   monthCellForTable = 5
    For i = 1 To 36
        If ActiveCell.Offset(0, 1).value = "" Then
            Exit For
        End If
        fstVal = DateSerial(Year(Replace(ActiveCell.Offset(0, 1).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(0, 1).value, ".", "/", 4)), 1)
        lstVal = DateSerial(Year(Replace(ActiveCell.Offset(0, 2).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(0, 2).value, ".", "/", 4)) + 1, 0)
        
        If fstVal <= CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) And CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) <= lstVal Then
            ActiveCell.Offset(0, monthCellForTable).value = "Yes"
        Else
            ActiveCell.Offset(0, monthCellForTable).value = "No"
        End If
        monthCellForTable = monthCellForTable + 1
    Next
End If

topCelVal = topCelVal + 1
ActiveCell.Offset(1, 0).Select
Next

End Sub
