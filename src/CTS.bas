Attribute VB_Name = "CTS"

Sub createPivotTableAggregatedData()

Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim rngData As String
Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim strtPt As String
Dim SrcData As String
Dim wsData As Worksheet
Dim wsPtTable As Worksheet
Dim pvtExcel As String
Dim wsptName  As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String

'pt.ManualUpdate = False
myPath = "D:\Philips\DashboardAutomation"
    pvtExcel = myPath & "\" & Dir(myPath & "\" & "MMBusC_InputData_" & "*.xls*")  'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    Sheets("Aggr. SWO Data CV").Activate
    Range("A2").Select
    fstCellAdd = ActiveCell.Address
    Range("A2").End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    
    ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=3, Criteria1:="=Buildingblocks Aggregated"
    Range("A2").Offset(1, 0).Select
    Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String
    fstFiltCellAdd = ActiveCell.Address
    Range("A2").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
    Range(fstFiltCellAdd1).End(xlToRight).Select
    fstFiltCellAdd2 = ActiveCell.Address
   ' lastFiltCellAdd = ActiveCell.Address
   Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
    ActiveSheet.Range("F1").Select
    Selection.UnMerge
   'ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=4, Criteria1:="=Non-Parts Aggregated"
    Range("F2").Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range("F2").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
Dim Max, tenPercentofMax, cellVal
ActiveSheet.Range(fstFiltCellAdd, fstFiltCellAdd1).Select
Max = Application.WorksheetFunction.Max(ActiveSheet.Range(fstFiltCellAdd, fstFiltCellAdd1))
    tenPercentofMax = Max / 10
    Dim rows As Range, cell As Range, value As Long

Set cell = Range("F3")
Do Until cell.value = ""
    value = Val(cell.value)
    If (value < tenPercentofMax) Then
        If rows Is Nothing Then
            Set rows = cell.EntireRow
        Else
            Set rows = Union(cell.EntireRow, rows)
        End If
    End If
    Set cell = cell.Offset(1)
Loop

If Not rows Is Nothing Then rows.Delete
    [n3].Resize(, 1).EntireColumn.Insert
  
Dim lastRow As Integer

 Workbooks(myPvtWorkBook).Sheets(2).Activate
'    ActiveSheet.ClearAllFilters

    lastRow = ActiveSheet.Range("M" & ActiveSheet.rows.Count).End(xlUp).Row
    Range("N2").value = "Total Cost of Parts & Non-Parts"
    Range("N3").FormulaR1C1 = _
        "=IF(OR(RC[-10]=""Non-Parts Aggregated"",RC[-10]=""Parts Aggregated""),(RC[-8]*RC[-7]*100)+RC[-5]*200,0)"
      '  Range("N3", "N" & Cells(rows.Count, 1).End(xlUp).Row).FillDown
    Range("N3").AutoFill Destination:=Range("N3:N" & lastRow)
    Range("N3:N" & lastRow).Select
    Calculate
    Selection.Copy
    Range("N3:N" & lastRow).PasteSpecial xlPasteValues
       
    ActiveWorkbook.Sheets(2).Activate
    Sheets.Add After:=Worksheets(Worksheets.Count)

    Set wsPtTable = Worksheets(Sheets.Count)

    Set wsPtTable = Worksheets(3)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets(2).Activate

    Set wsData = Worksheets("Aggr. SWO Data CV")
    Worksheets("Aggr. SWO Data CV").Activate
    sourceSheet = ActiveSheet.name

    ActiveSheet.Cells(2, 1).Select
    fstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select

    lstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
    Sheets("Sheet1").Activate
    rngData = fstadd & ":" & lstadd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
                
        wsPtTable.Activate
        
    Set pt = wsPtTable.PivotTables("PivotTable1")
    Set pf = pt.PivotFields("SubSystem")
    pf.Orientation = xlRowField
    pf.Position = 1
    Set pf = pt.PivotFields("BuildingBlock")
    pf.Orientation = xlRowField
    pf.Position = 2
    Set pf = pt.PivotFields("Part12NC")
    pf.Orientation = xlColumnField
    pf.Position = 1
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Calls (#)"), "# of Calls", xlSum
        
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Avg. MTTR/Call (hrs)"), "MTTR/Call (hrs)", xlSum
    
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Avg. ETTR (days)"), "ETTR (days)", xlSum
    
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Avg. Visits/call (#)"), "Visits/call (#)", xlSum
        
        ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
        
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Costs/part (EUR)"), "Costs/part (EUR)", xlSum
    
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Cost of Parts & Non-Parts"), _
        "#Total Cost of Parts & Non-Parts", xlSum
            
    
        
    With ActiveSheet.PivotTables("PivotTable1")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    
        ActiveSheet.PivotTables("PivotTable1").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    With pt.PivotFields("Part12NC")
    pf.Orientation = xlColumnField
    pf.Position = 2
    End With
    
    Dim pvtItm As PivotItem
    Set PvtTbl = Worksheets("Sheet1").PivotTables("PivotTable1")
    PvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    PvtTbl.RefreshTable

    Columns("C:C").Select
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 13012579
        .TintAndShade = 0
    End With
 
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 13012579
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").Select
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 2668287
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 2668287
        .TintAndShade = 0
    End With
    
   With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With
    
  Columns("E:E").Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 20
        .Percent = True
    End With
    With Selection.FormatConditions(1).Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("F23").Select
    
    Columns("F:F").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("F19").Select
    
    Columns("G:G").Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 10
        .Percent = True
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("H:H").Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 10
        .Percent = True
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("I:J").Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 20
        .Percent = True
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    

    Columns("I:I").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("C:N").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 11
End With
    Range("M2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Columns("A:B").Select
    With Selection
        .ColumnWidth = 15
    End With
    Cells(1, 1).Select

    ActiveWindow.Zoom = 85
Worksheets("Sheet1").PivotTables("PivotTable1").PreserveFormatting = True

  '  pt.ManualUpdate = True
    MsgBox "Pivot Table is created succesfully", vbOKOnly

End Sub

Public myWorkBook As String
Sub CR()
    installFlName = ThisWorkbook.Path & "\" & "Veradius_Aug_2015_Jun_2013" & ".xlsx"
    Application.Workbooks.Open (installFlName), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstadd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstadd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstadd, lstadd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstadd, lstadd).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
        Workbooks("KPI's_NewVer_1.0_change_R").Activate
        Sheets("Sheet7").Select
        Range("A:A").Select
        On Error Resume Next
        Selection.EntireRow.Select
        Selection.EntireRow.Delete
        Application.Columns.Ungroup
        rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        Sheets("KPI-All").Select
        ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
        Selection.Copy
        Sheets("Sheet7").Select
        Range("a1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("1:1").Select
        Selection.EntireRow.Delete
        Columns("E:J").Select
        Selection.EntireColumn.Delete
         Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "All Systems"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "All Blocks"
    'Range("B5").Select
    Windows("CTS_Guidelines.xlsx").Activate
    Sheets("Sheet2").Activate
    Sheets("Sheet2").UsedRange.Find(what:="CR / Sys / Yr", lookat:=xlWhole).Select

    Selection.EntireColumn.Select
    Selection.Copy
    
    Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
    Range("C1").Select
    ActiveSheet.Paste
    Range("F3").Select
    Application.CutCopyMode = False
  
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]+RC[-2])/2047"
    Range("F3").Select
    Selection.AutoFill Destination:=Range("F3:F41")
    Range("F3:F41").Select
    Range("F3:F41").NumberFormat = "0.00"

    Calculate
    
    Range("A2").Select
    fstadd1 = ActiveCell.Address
    Sheets("Sheet7").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    lstadd2 = ActiveCell.Address
    Range(fstadd1, lstadd2).Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    Sheets("Sheet7").Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C91")
    Range("C3:C91").Select
    Calculate
    'Columns("AE:AE").Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("AE3").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    'Range("AE3").Select
    'Selection.AutoFill Destination:=Range("AE3:AE91")
    'Range("AE3:AE91").Select
    'Calculate
    
        conditionalfrmngCR
        addTBLHeadingCR
        createSheetCR
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub

Sub conditionalfrmngCR()

    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Cells(3, 7).Select
    ActiveCell.End(xlDown).Select
    lstRowAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    Range(lstRowAdd).Select
    Sheets("Sheet7").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    
    Range("E3").Select
    pkAdd = ActiveCell.Address
    fstCellAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    mioflstcell = Left(fstCellAdd, 3)
    midoflstadd = Mid(lstRowAdd, 4)
    Add = mioflstcell & midoflstadd
    ActiveSheet.Range(fstCellAdd, Add).Select
        
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 2668287
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 2668287
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("F3").Select
    
    pkAdd1 = ActiveCell.Address
    fstCellAdd1 = ActiveCell.Address(ReferenceStyle:=xlA1)
    mioflstcelll = Left(fstCellAdd1, 3)
    midoflstadd = Mid(lstRowAdd, 4)
    add1 = mioflstcelll & midoflstadd
    ActiveSheet.Range(fstCellAdd1, add1).Select
        
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
        
    Range(fstCellAdd, add1).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormat = "General"

    End With
    
End Sub

Sub addTBLHeadingCR()

    Range("E1").Select
    ActiveCell.value = "MAT # of Calls profiles"
    
    Range("E2").Select
    ActiveCell.value = "Non-Parts"
    Range("F2").Select
    ActiveCell.value = "Parts"
    Range("G2").Select
    ActiveCell.value = "CR / Sys / ITM"
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("H1").Select
    ActiveCell.value = "Current Year CR / Sys"
    Range("H2").Select
    ActiveCell.value = "ITM"
    Range("I2").Select
    ActiveCell.value = "IMQ"
    Range("J2").Select
    ActiveCell.value = "YTD"
    Range("K2").Select
    ActiveCell.value = "MAT"
    Range("H1:K1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L1").Select
    ActiveCell.value = "VLY"
    Range("L2").Select
    ActiveCell.value = "ITM"
    Range("M2").Select
    ActiveCell.value = "IMQ"
    Range("N2").Select
    ActiveCell.value = "YTD"
    Range("O2").Select
    ActiveCell.value = "MAT"
    Range("L1:O1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("P1").Select
    ActiveCell.value = "Crossover"
     Range("P2").Select
    ActiveCell.value = "Trigger"
    Range("Q2").Select
    ActiveCell.value = "M12"
    Range("R2").Select
    ActiveCell.value = "M11"
    Range("S2").Select
    ActiveCell.value = "M10"
    Range("T2").Select
    ActiveCell.value = "M9"
    Range("U2").Select
    ActiveCell.value = "M8"
    Range("V2").Select
    ActiveCell.value = "M7"
    Range("W2").Select
    ActiveCell.value = "M6"
    Range("X2").Select
    ActiveCell.value = "M5"
    Range("Y2").Select
    ActiveCell.value = "M4"
    Range("Z2").Select
    ActiveCell.value = "M3"
    Range("AA2").Select
    ActiveCell.value = "M2"
    Range("AB2").Select
    ActiveCell.value = "M1"
    Range("D1").Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D1").Select
    Selection.Copy
    Range("D1:F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("A1:A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("B1:B2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("C1:C2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("D2:AB2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("G1:I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("J1:M1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("N1:P1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Range("H1:K1").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L1:O1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveSheet.UsedRange.Select
    Selection.RowHeight = 15
    Range("H1:P2").Select
    Selection.Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlRight
    End With
Sheets("Sheet7").UsedRange.Find(what:="CR / Sys / ITM", lookat:=xlWhole).Select
Sheets("Sheet7").UsedRange.Find(what:="ITM", After:=ActiveCell, lookat:=xlWhole).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.EntireColumn.Select
Selection.ColumnWidth = 6

End Sub
Sub createSheetCR()
CRPivotTable
CRITM
CRIMQ
CRYTD
CRMAT
CRITMPrvs
CRIMQPrvs
CRYTDPrvs
CRMATPrvs
CRMonthlyCR
CRfinalFormatting
End Sub
Sub CRITM()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
Set PvtTbl = Worksheets("Sheet7").PivotTables("PivotTable1")
fixedDate = 201406
'currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2)
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = startDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

      
    Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 8).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
    Cells(x, 8).NumberFormat = "0"
  Next x
             
End Sub
Sub CRIMQ()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = Worksheets("Sheet7").PivotTables("PivotTable1")
PvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, startDate), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, startDate), "yyyy" & "-" & "mm")

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 9).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 3
    Cells(x, 9).NumberFormat = "0"
    
    'Application.WorksheetFunction.RoundUp (Cells(x, 8).Value)
    'Application.RoundUp (Cells(x, 9).Value)
    Next x
End Sub
Sub CRMAT()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")
   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 11).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 12
    Cells(x, 11).NumberFormat = "0"
    Next x

End Sub
Sub CRITMPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = endDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

    Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 12).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
    Cells(x, 12).NumberFormat = "0"
    Next x
    
End Sub
Sub CRIMQPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
prvsIMQ = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = Worksheets("Sheet7").PivotTables("PivotTable1")
PvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 13).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 3
    Cells(x, 13).NumberFormat = "0"
    Next x
End Sub

Sub CRMATPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(DateAdd("yyyy", -1, Date), "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -2, Date), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 15).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 12
    Cells(x, 15).NumberFormat = "0"
    Next x

End Sub
Sub CRYTD()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 5, 2)

endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")

   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 10).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / EndDateMonth
    Cells(x, 10).NumberFormat = "0"
    Next x


End Sub

Sub CRYTDPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
StartDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 5, 2)

endDate = Format(DateAdd("m", -EndDateMonth, StartDate1), "yyyy" & "-" & "mm")

   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > StartDate1 Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    Dim x As Long
    lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 14).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / EndDateMonth
    Cells(x, 14).NumberFormat = "0"
    Next x


End Sub

Sub MonthlyCR()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")
   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
Cells(3, 16).Select
i = 17
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
    Else
            pf.CurrentPage = pvtItm.Caption
            Dim x As Long
            lr = Worksheets("Sheet7").Cells(rows.Count, "C").End(xlUp).Row
            Rng = Range("AE3:AJ91")
            
            If i <= 28 Then
            For x = 2 To lr
            On Error Resume Next
            Cells(x, i).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
            'Round (Cells(x, i).Value)

            Next x
             
    End If
    i = i + 1
    End If
Next pvtItm
   
End Sub
Sub finalFormattingCR()

Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H3:O43").Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[68]C)"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    
    Sheets("Sheet7").Select
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]>RC[-3]"
    Range("P3").Select
    Selection.AutoFill Destination:=Range("P3:P91")
    Range("P3:P91").Select
    Calculate
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P2").Select
    Selection.End(xlDown).Select
    Range("Q91").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("Q3:AB91").Select
    Range("Q91").Activate
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("Q3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Q3:AB91").Select
    
    Application.CutCopyMode = False
    Range("AC3:AC91").Select
    Range("$AC$3:$AC$91").SparklineGroups.Add Type:=xlSparkLine, SourceData:= _
        "Q3:AB91"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Range("AC4").Select
    ActiveCell.FormulaR1C1 = ""
    Range("AB2").Select
    Selection.Copy
    Range("AC2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("G3").Select
    Selection.End(xlDown).Select
    Range("H41").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H41:AC1048576").Select
   
    Selection.ClearContents
    
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 10
        .Percent = False
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("P3").Select
    Range(Selection, Selection.End(xlDown)).Select
   
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TRUE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 240
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
   
End Sub

Sub CRPivotTable()
    
    Workbooks("Veradius_Aug_2015_Jun_2013.xlsx").Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Sheets.Add
    pvtSheetName = ActiveSheet.name
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Aggr. SWO Data CV!R1C1:R2981C21", Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:=pvtSheetName & "!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15
    Sheets(pvtSheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SubSystem")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Calls (#)"), "Count of Total Calls (#)", _
        xlCount
    With ActiveSheet.PivotTables("PivotTable1")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A7").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Range("B7").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Range("C7").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock")
        .PivotItems("Buildingblocks Aggregated").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("132205111168").Visible = False
        .PivotItems("132205111173").Visible = False
        .PivotItems("132205111189").Visible = False
        .PivotItems("132205411156").Visible = False
        .PivotItems("242202800229").Visible = False
        .PivotItems("242203000196").Visible = False
        .PivotItems("242208600166").Visible = False
        .PivotItems("242208600546").Visible = False
        .PivotItems("242212000636").Visible = False
        .PivotItems("242212916007").Visible = False
        .PivotItems("242254944424").Visible = False
        .PivotItems("243050000065").Visible = False
        .PivotItems("251278502015").Visible = False
        .PivotItems("251278502075").Visible = False
        .PivotItems("252204314008").Visible = False
        .PivotItems("252240109011").Visible = False
        .PivotItems("252272808005").Visible = False
        .PivotItems("262200130078").Visible = False
        .PivotItems("262285521091").Visible = False
        .PivotItems("282206502596").Visible = False
        .PivotItems("451000035931").Visible = False
        .PivotItems("451210045251").Visible = False
        .PivotItems("451210177141").Visible = False
        .PivotItems("451210498891").Visible = False
        .PivotItems("451210788004").Visible = False
        .PivotItems("451213056931").Visible = False
        .PivotItems("451213118611").Visible = False
        .PivotItems("451213406061").Visible = False
        .PivotItems("451213435751").Visible = False
        .PivotItems("451214843172").Visible = False
        .PivotItems("451220106902").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("451220109901").Visible = False
        .PivotItems("451291694535").Visible = False
        .PivotItems("451291768093").Visible = False
        .PivotItems("451291768123").Visible = False
        .PivotItems("451980042911").Visible = False
        .PivotItems("451980044391").Visible = False
        .PivotItems("452201101011").Visible = False
        .PivotItems("452205900151").Visible = False
        .PivotItems("452205901511").Visible = False
        .PivotItems("452209000971").Visible = False
        .PivotItems("452209004501").Visible = False
        .PivotItems("452209004841").Visible = False
        .PivotItems("452209008352").Visible = False
        .PivotItems("452209014208").Visible = False
        .PivotItems("452209017181").Visible = False
        .PivotItems("452209018231").Visible = False
        .PivotItems("452209018241").Visible = False
        .PivotItems("452209024951").Visible = False
        .PivotItems("452210280082").Visible = False
        .PivotItems("452210280102").Visible = False
        .PivotItems("452210280122").Visible = False
        .PivotItems("452210280142").Visible = False
        .PivotItems("452210355241").Visible = False
        .PivotItems("452210457286").Visible = False
        .PivotItems("452210459405").Visible = False
        .PivotItems("452210459421").Visible = False
        .PivotItems("452210459453").Visible = False
        .PivotItems("452210466113").Visible = False
        .PivotItems("452210601902").Visible = False
        .PivotItems("452212624336").Visible = False
        .PivotItems("452212650053").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452212650054").Visible = False
        .PivotItems("452212672141").Visible = False
        .PivotItems("452212702938").Visible = False
        .PivotItems("452212825601").Visible = False
        .PivotItems("452212857534").Visible = False
        .PivotItems("452212876313").Visible = False
        .PivotItems("452212876642").Visible = False
        .PivotItems("452212876681").Visible = False
        .PivotItems("452212876761").Visible = False
        .PivotItems("452212876781").Visible = False
        .PivotItems("452212876791").Visible = False
        .PivotItems("452212905322").Visible = False
        .PivotItems("452212905782").Visible = False
        .PivotItems("452212906782").Visible = False
        .PivotItems("452213171001").Visible = False
        .PivotItems("452214239252").Visible = False
        .PivotItems("452216422653").Visible = False
        .PivotItems("452216424613").Visible = False
        .PivotItems("452216424614").Visible = False
        .PivotItems("452216424624").Visible = False
        .PivotItems("452216424625").Visible = False
        .PivotItems("452216424661").Visible = False
        .PivotItems("452216424962").Visible = False
        .PivotItems("452216501133").Visible = False
        .PivotItems("452216501143").Visible = False
        .PivotItems("452216502181").Visible = False
        .PivotItems("452216503993").Visible = False
        .PivotItems("452216505961").Visible = False
        .PivotItems("452216506872").Visible = False
        .PivotItems("452216506882").Visible = False
        .PivotItems("452216507203").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452216507204").Visible = False
        .PivotItems("452216507249").Visible = False
        .PivotItems("452216508221").Visible = False
        .PivotItems("452216508231").Visible = False
        .PivotItems("452216508292").Visible = False
        .PivotItems("452216508451").Visible = False
        .PivotItems("452221042312").Visible = False
        .PivotItems("452230005332").Visible = False
        .PivotItems("452230007341").Visible = False
        .PivotItems("452230007351").Visible = False
        .PivotItems("452230007361").Visible = False
        .PivotItems("452230009052").Visible = False
        .PivotItems("452230009062").Visible = False
        .PivotItems("452230009181").Visible = False
        .PivotItems("452230014021").Visible = False
        .PivotItems("452230014031").Visible = False
        .PivotItems("452230014042").Visible = False
        .PivotItems("452230014043").Visible = False
        .PivotItems("452230014122").Visible = False
        .PivotItems("452230014123").Visible = False
        .PivotItems("452230014191").Visible = False
        .PivotItems("452230014292").Visible = False
        .PivotItems("452230014301").Visible = False
        .PivotItems("452230014342").Visible = False
        .PivotItems("452230014521").Visible = False
        .PivotItems("452230014661").Visible = False
        .PivotItems("452230014894").Visible = False
        .PivotItems("452230014906").Visible = False
        .PivotItems("452230014971").Visible = False
        .PivotItems("452230016714").Visible = False
        .PivotItems("452230019854").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452230019872").Visible = False
        .PivotItems("452230019881").Visible = False
        .PivotItems("452230019891").Visible = False
        .PivotItems("452230019921").Visible = False
        .PivotItems("452230019931").Visible = False
        .PivotItems("452230019941").Visible = False
        .PivotItems("452230019952").Visible = False
        .PivotItems("452230019971").Visible = False
        .PivotItems("452230019981").Visible = False
        .PivotItems("452230019991").Visible = False
        .PivotItems("452230026001").Visible = False
        .PivotItems("452230026021").Visible = False
        .PivotItems("452230026031").Visible = False
        .PivotItems("452230026041").Visible = False
        .PivotItems("452230026051").Visible = False
        .PivotItems("452230026061").Visible = False
        .PivotItems("452230026062").Visible = False
        .PivotItems("452230026101").Visible = False
        .PivotItems("452230026172").Visible = False
        .PivotItems("452230026173").Visible = False
        .PivotItems("452230026211").Visible = False
        .PivotItems("452230026221").Visible = False
        .PivotItems("452230026251").Visible = False
        .PivotItems("452230028542").Visible = False
        .PivotItems("452230028571").Visible = False
        .PivotItems("452230028581").Visible = False
        .PivotItems("452230028721").Visible = False
        .PivotItems("452230028921").Visible = False
        .PivotItems("452230028931").Visible = False
        .PivotItems("452230028941").Visible = False
        .PivotItems("452230029041").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452230029081").Visible = False
        .PivotItems("452230029101").Visible = False
        .PivotItems("452230029111").Visible = False
        .PivotItems("452230029132").Visible = False
        .PivotItems("452230029153").Visible = False
        .PivotItems("452230029154").Visible = False
        .PivotItems("452230029161").Visible = False
        .PivotItems("452230029172").Visible = False
        .PivotItems("452250066171").Visible = False
        .PivotItems("452250066241").Visible = False
        .PivotItems("452298037862").Visible = False
        .PivotItems("452298037871").Visible = False
        .PivotItems("452298037881").Visible = False
        .PivotItems("452298037891").Visible = False
        .PivotItems("452298038081").Visible = False
        .PivotItems("452298038352").Visible = False
        .PivotItems("452298038422").Visible = False
        .PivotItems("452298038603").Visible = False
        .PivotItems("453560087111").Visible = False
        .PivotItems("453561153714").Visible = False
        .PivotItems("453561153724").Visible = False
        .PivotItems("453561158311").Visible = False
        .PivotItems("453561219221").Visible = False
        .PivotItems("453564260891").Visible = False
        .PivotItems("453564347391").Visible = False
        .PivotItems("453566440751").Visible = False
        .PivotItems("453567914251").Visible = False
        .PivotItems("453580488115").Visible = False
        .PivotItems("453580488116").Visible = False
        .PivotItems("455300002861").Visible = False
        .PivotItems("459800000561").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800000573").Visible = False
        .PivotItems("459800001091").Visible = False
        .PivotItems("459800001443").Visible = False
        .PivotItems("459800001444").Visible = False
        .PivotItems("459800001445").Visible = False
        .PivotItems("459800001681").Visible = False
        .PivotItems("459800002531").Visible = False
        .PivotItems("459800002541").Visible = False
        .PivotItems("459800002572").Visible = False
        .PivotItems("459800002573").Visible = False
        .PivotItems("459800002582").Visible = False
        .PivotItems("459800002591").Visible = False
        .PivotItems("459800002601").Visible = False
        .PivotItems("459800002612").Visible = False
        .PivotItems("459800002613").Visible = False
        .PivotItems("459800002623").Visible = False
        .PivotItems("459800002625").Visible = False
        .PivotItems("459800002631").Visible = False
        .PivotItems("459800003061").Visible = False
        .PivotItems("459800003142").Visible = False
        .PivotItems("459800011751").Visible = False
        .PivotItems("459800011761").Visible = False
        .PivotItems("459800011771").Visible = False
        .PivotItems("459800012171").Visible = False
        .PivotItems("459800013612").Visible = False
        .PivotItems("459800013631").Visible = False
        .PivotItems("459800020611").Visible = False
        .PivotItems("459800034242").Visible = False
        .PivotItems("459800037342").Visible = False
        .PivotItems("459800048691").Visible = False
        .PivotItems("459800061991").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800062051").Visible = False
        .PivotItems("459800062091").Visible = False
        .PivotItems("459800063241").Visible = False
        .PivotItems("459800066331").Visible = False
        .PivotItems("459800066333").Visible = False
        .PivotItems("459800066334").Visible = False
        .PivotItems("459800066335").Visible = False
        .PivotItems("459800066611").Visible = False
        .PivotItems("459800070281").Visible = False
        .PivotItems("459800070291").Visible = False
        .PivotItems("459800072661").Visible = False
        .PivotItems("459800072671").Visible = False
        .PivotItems("459800099201").Visible = False
        .PivotItems("459800108861").Visible = False
        .PivotItems("459800125171").Visible = False
        .PivotItems("459800125172").Visible = False
        .PivotItems("459800128372").Visible = False
        .PivotItems("459800148821").Visible = False
        .PivotItems("459800151382").Visible = False
        .PivotItems("459800151421").Visible = False
        .PivotItems("459800151422").Visible = False
        .PivotItems("459800151442").Visible = False
        .PivotItems("459800153441").Visible = False
        .PivotItems("459800155441").Visible = False
        .PivotItems("459800162341").Visible = False
        .PivotItems("459800164611").Visible = False
        .PivotItems("459800173431").Visible = False
        .PivotItems("459800173432").Visible = False
        .PivotItems("459800196181").Visible = False
        .PivotItems("459800219711").Visible = False
        .PivotItems("459800220541").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800238711").Visible = False
        .PivotItems("459800240611").Visible = False
        .PivotItems("459800240621").Visible = False
        .PivotItems("459800240681").Visible = False
        .PivotItems("459800240683").Visible = False
        .PivotItems("459800240691").Visible = False
        .PivotItems("459800240731").Visible = False
        .PivotItems("459800240741").Visible = False
        .PivotItems("459800240781").Visible = False
        .PivotItems("459800240801").Visible = False
        .PivotItems("459800240821").Visible = False
        .PivotItems("459800240841").Visible = False
        .PivotItems("459800240961").Visible = False
        .PivotItems("459800260121").Visible = False
        .PivotItems("459800267151").Visible = False
        .PivotItems("459800274261").Visible = False
        .PivotItems("459800295301").Visible = False
        .PivotItems("459800319211").Visible = False
        .PivotItems("459800319212").Visible = False
        .PivotItems("459800320161").Visible = False
        .PivotItems("459800359091").Visible = False
        .PivotItems("459800359092").Visible = False
        .PivotItems("459800372151").Visible = False
        .PivotItems("459800418511").Visible = False
        .PivotItems("459800440311").Visible = False
        .PivotItems("459800609732").Visible = False
        .PivotItems("459800671421").Visible = False
        .PivotItems("459800766581").Visible = False
        .PivotItems("867000053429").Visible = False
        .PivotItems("929900059707").Visible = False
        .PivotItems("989600007772").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("989600008501").Visible = False
        .PivotItems("989600180815").Visible = False
        .PivotItems("989600193001").Visible = False
        .PivotItems("989600204612").Visible = False
        .PivotItems("989600206924").Visible = False
        .PivotItems("989600216801").Visible = False
        .PivotItems("989601000652").Visible = False
        .PivotItems("989601023201").Visible = False
        .PivotItems("989601041312").Visible = False
        .PivotItems("989601041313").Visible = False
        .PivotItems("989601063621").Visible = False
        .PivotItems("989601065321").Visible = False
        .PivotItems("989670000011").Visible = False
        .PivotItems("989710002291").Visible = False
        .PivotItems("989710005263").Visible = False
        .PivotItems("989710006151").Visible = False
        .PivotItems("991920050193").Visible = False
        .PivotItems("991920050194").Visible = False
        .PivotItems("991920160462").Visible = False
        .PivotItems("991932050882").Visible = False
        .PivotItems("991932050883").Visible = False
        .PivotItems("991932050912").Visible = False
        .PivotItems("991932050913").Visible = False
        .PivotItems("991932050923").Visible = False
        .PivotItems("991932051114").Visible = False
        .PivotItems("991932212002").Visible = False
        .PivotItems("991932212011").Visible = False
        .PivotItems("991932472041").Visible = False
        .PivotItems("All Aggregated").Visible = False
        '.PivotItems("(blank)").Visible = False
    End With
    Range("C4").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"
    Range("D4").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
    Range("B6").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
    ActiveWindow.SmallScroll ToRight:=16
    Windows("Veradius_Aug_2015_Jun_2013.xlsx").Activate
    ActiveSheet.PivotTables("PivotTable1").Location = _
        "'[KPI''s_NewVer_1.0_Change_R.xlsm]Sheet7'!$AK$3"
        Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
        Sheets("Sheet7").Activate
    'ActiveSheet.PivotTables("PivotTable1").PivotSelect "Period", xlButton, True
    'ActiveSheet.PivotTables("PivotTable1").Location = "Sheet7!$AK$3"
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[5]"
    Range("AF3").Select
    Selection.Copy
    
    Range("AF3,AF91").Select
    
    Range("AF3,AF3:AJ91").Select
    ActiveSheet.Paste
    
    Range("AE3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    Range("AE3").Select
    Selection.Copy
    
    Range("AE3:AE91").Select
    ActiveSheet.Paste
    
End Sub

Public myWorkBook As String
Sub MTTR()
    installFlName = ThisWorkbook.Path & "\" & "Veradius_Aug_2015_Jun_2013" & ".xlsx"
    Application.Workbooks.Open (installFlName), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstadd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstadd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstadd, lstadd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstadd, lstadd).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
        Workbooks("KPI's_NewVer_1.0_change_R").Activate
        Worksheets.Add().name = "MTTR"

        Sheets("MTTR").Select
        Range("A:A").Select
        On Error Resume Next
        Selection.EntireRow.Select
        Selection.EntireRow.Delete
        Application.Columns.Ungroup
        rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        Sheets("KPI-All").Select
        ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
        Selection.Copy


        Sheets("MTTR").Select
        Range("a1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("1:1").Select
        Selection.EntireRow.Delete
        Columns("E:J").Select
        Selection.EntireColumn.Delete
         Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "All Systems"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "All Blocks"
    'Range("B5").Select
    Windows("CTS_Guidelines.xlsx").Activate
    Sheets("Sheet2").Activate
    Sheets("Sheet2").UsedRange.Find(what:="MTTR/ Sys / Yr", lookat:=xlWhole).Select

    Selection.EntireColumn.Select
    Selection.Copy
    Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
    Range("C1").Select
    ActiveSheet.Paste
    Range("F3").Select
    Application.CutCopyMode = False
  
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]+RC[-2])/2047"
    Range("F3").Select
    Selection.AutoFill Destination:=Range("F3:F41")
    Range("F3:F41").Select
    Range("F3:F41").NumberFormat = "0.00"

    Calculate
    
    Range("A2").Select
    fstadd1 = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    lstadd2 = ActiveCell.Address
    Range(fstadd1, lstadd2).Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    Sheets("MTTR").Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C91")
    Range("C3:C91").Select
    Calculate
    'Columns("AE:AE").Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("AE3").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    'Range("AE3").Select
    'Selection.AutoFill Destination:=Range("AE3:AE91")
    'Range("AE3:AE91").Select
    'Calculate
    
        conditionalfrmngMTTR
        addTBLHeadingMTTR
        createSheetMTTR
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub

Sub conditionalfrmngMTTR()

    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Cells(3, 7).Select
    ActiveCell.End(xlDown).Select
    lstRowAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    Range(lstRowAdd).Select
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    
    Range("E3").Select
    pkAdd = ActiveCell.Address
    fstCellAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    mioflstcell = Left(fstCellAdd, 3)
    midoflstadd = Mid(lstRowAdd, 4)
    Add = mioflstcell & midoflstadd
    ActiveSheet.Range(fstCellAdd, Add).Select
        
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 2668287
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 2668287
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("F3").Select
    
    pkAdd1 = ActiveCell.Address
    fstCellAdd1 = ActiveCell.Address(ReferenceStyle:=xlA1)
    mioflstcelll = Left(fstCellAdd1, 3)
    midoflstadd = Mid(lstRowAdd, 4)
    add1 = mioflstcelll & midoflstadd
    ActiveSheet.Range(fstCellAdd1, add1).Select
        
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
        
    Range(fstCellAdd, add1).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormat = "General"

    End With
    
End Sub

Sub addTBLHeadingMTTR()

    Range("E1").Select
    ActiveCell.value = "MAT # of Calls profiles"
    
    Range("E2").Select
    ActiveCell.value = "Non-Parts"
    Range("F2").Select
    ActiveCell.value = "Parts"
    Range("G2").Select
    ActiveCell.value = "MTTR/ Sys / Yr"
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("H1").Select
    ActiveCell.value = "Current Year MTTR/ Sys"
    Range("H2").Select
    ActiveCell.value = "ITM"
    Range("I2").Select
    ActiveCell.value = "IMQ"
    Range("J2").Select
    ActiveCell.value = "YTD"
    Range("K2").Select
    ActiveCell.value = "MAT"
    Range("H1:K1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L1").Select
    ActiveCell.value = "VLY"
    Range("L2").Select
    ActiveCell.value = "ITM"
    Range("M2").Select
    ActiveCell.value = "IMQ"
    Range("N2").Select
    ActiveCell.value = "YTD"
    Range("O2").Select
    ActiveCell.value = "MAT"
    Range("L1:O1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("P1").Select
    ActiveCell.value = "Crossover"
     Range("P2").Select
    ActiveCell.value = "Trigger"
    Range("Q2").Select
    ActiveCell.value = "M12"
    Range("R2").Select
    ActiveCell.value = "M11"
    Range("S2").Select
    ActiveCell.value = "M10"
    Range("T2").Select
    ActiveCell.value = "M9"
    Range("U2").Select
    ActiveCell.value = "M8"
    Range("V2").Select
    ActiveCell.value = "M7"
    Range("W2").Select
    ActiveCell.value = "M6"
    Range("X2").Select
    ActiveCell.value = "M5"
    Range("Y2").Select
    ActiveCell.value = "M4"
    Range("Z2").Select
    ActiveCell.value = "M3"
    Range("AA2").Select
    ActiveCell.value = "M2"
    Range("AB2").Select
    ActiveCell.value = "M1"
    Range("D1").Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D1").Select
    Selection.Copy
    Range("D1:F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("A1:A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("B1:B2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("C1:C2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("D2:AB2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("G1:I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("J1:M1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("N1:P1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Range("H1:K1").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L1:O1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveSheet.UsedRange.Select
    Selection.RowHeight = 15
    Range("H1:P2").Select
    Selection.Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlRight
    End With
Sheets("MTTR").UsedRange.Find(what:="MTTR/ Sys / Yr", lookat:=xlWhole).Select
Sheets("MTTR").UsedRange.Find(what:="ITM", After:=ActiveCell, lookat:=xlWhole).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.EntireColumn.Select
Selection.ColumnWidth = 6

End Sub
Sub createSheetMTTR()
MTTRPivotTable
MTTRITM
MTTRIMQ
MTTRYTD
MTTRMAT
MTTRITMPrvs
MTTRIMQPrvs
MTTRYTDPrvs
MTTRMATPrvs
MTTRMonthly
MTTRfinalFormatting
End Sub



Sub MTTRITM()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
Set PvtTbl = Worksheets("MTTR").PivotTables("PivotTable1")
fixedDate = 201406
'currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2)
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = startDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

      
    Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 8).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
    Cells(x, 8).NumberFormat = "0.00"
  Next x
             
End Sub
Sub MTTRIMQ()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = Worksheets("MTTR").PivotTables("PivotTable1")
PvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, startDate), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, startDate), "yyyy" & "-" & "mm")

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 9).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 3
    Cells(x, 9).NumberFormat = "0.00"
    
    'Application.WorksheetFunction.RoundUp (Cells(x, 8).Value)
    'Application.RoundUp (Cells(x, 9).Value)
    Next x
End Sub
Sub MTTRMAT()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")
   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 11).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 12
    Cells(x, 11).NumberFormat = "0.00"
    Next x

End Sub
Sub MTTRITMPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = endDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

    Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 12).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
    Cells(x, 12).NumberFormat = "0.00"
    Next x
    
End Sub
Sub MTTRIMQPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
prvsIMQ = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")

Set PvtTbl = Worksheets("MTTR").PivotTables("PivotTable1")
PvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 13).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 3
    Cells(x, 13).NumberFormat = "0.00"
    Next x
End Sub

Sub MTTRMATPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(DateAdd("yyyy", -1, Date), "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -2, Date), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")

Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 15).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / 12
    Cells(x, 15).NumberFormat = "0.00"
    Next x

End Sub
Sub MTTRYTD()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 5, 2)

endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")

   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 10).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / EndDateMonth
    Cells(x, 10).NumberFormat = "0.00"
    Next x


End Sub

Sub MTTRYTDPrvs()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
StartDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 5, 2)

endDate = Format(DateAdd("m", -EndDateMonth, StartDate1), "yyyy" & "-" & "mm")

   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > StartDate1 Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    Dim x As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    Rng = Range("AE3:AJ91")
            
    For x = 3 To lr
    On Error Resume Next
    Cells(x, 14).value = (Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)) / EndDateMonth
    Cells(x, 14).NumberFormat = "0.00"
    Next x


End Sub

Sub MTTRMonthly()
Dim PvtTbl As PivotTable
Dim pvtItm As PivotItem
Dim visPvtItm As String
Dim pf As PivotField
fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
EndDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, EndDate1), "yyyy" & "-" & "mm")
   
Set PvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = PvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
Cells(3, 16).Select
i = 17
For Each pvtItm In PvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
    Else
            pf.CurrentPage = pvtItm.Caption
            Dim x As Long
            lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
            Rng = Range("AE3:AJ91")
            
            If i <= 28 Then
            For x = 2 To lr
            On Error Resume Next
            Cells(x, i).value = Application.WorksheetFunction.VLookup(Cells(x, 3).value, Rng, 6, False)
            Cells(x, i).NumberFormat = "0.00"

            Next x
             
    End If
    i = i + 1
    End If
Next pvtItm
   
End Sub
Sub MTTRfinalFormatting()

Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H3:O43").Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[68]C)"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    
    Sheets("MTTR").Select
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]>RC[-3]"
    Range("P3").Select
    Selection.AutoFill Destination:=Range("P3:P91")
    Range("P3:P91").Select
    Calculate
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P2").Select
    Selection.End(xlDown).Select
    Range("Q91").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("Q3:AB91").Select
    Range("Q91").Activate
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("Q3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Q3:AB91").Select
    
    Application.CutCopyMode = False
    Range("AC3:AC91").Select
    Range("$AC$3:$AC$91").SparklineGroups.Add Type:=xlSparkLine, SourceData:= _
        "Q3:AB91"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Range("AC4").Select
    ActiveCell.FormulaR1C1 = ""
    Range("AB2").Select
    Selection.Copy
    Range("AC2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("G3").Select
    Selection.End(xlDown).Select
    Range("H41").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H41:AC1048576").Select
   
    Selection.ClearContents
    
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 10
        .Percent = False
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("P3").Select
    Range(Selection, Selection.End(xlDown)).Select
   
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TRUE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 240
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.Zoom = 85
   
End Sub

Sub MTTRPivotTable()
    
    Workbooks("Veradius_Aug_2015_Jun_2013.xlsx").Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Sheets.Add
    pvtSheetName = ActiveSheet.name
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Aggr. SWO Data CV!R1C1:R2981C21", Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:=pvtSheetName & "!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion15
    Sheets(pvtSheetName).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SubSystem")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Avg. MTTR/Call (hrs)"), "#MTTR/Call (hrs)", _
        xlSum
    With ActiveSheet.PivotTables("PivotTable1")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A7").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Range("B7").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Range("C7").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("BuildingBlock")
        .PivotItems("Buildingblocks Aggregated").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("132205111168").Visible = False
        .PivotItems("132205111173").Visible = False
        .PivotItems("132205111189").Visible = False
        .PivotItems("132205411156").Visible = False
        .PivotItems("242202800229").Visible = False
        .PivotItems("242203000196").Visible = False
        .PivotItems("242208600166").Visible = False
        .PivotItems("242208600546").Visible = False
        .PivotItems("242212000636").Visible = False
        .PivotItems("242212916007").Visible = False
        .PivotItems("242254944424").Visible = False
        .PivotItems("243050000065").Visible = False
        .PivotItems("251278502015").Visible = False
        .PivotItems("251278502075").Visible = False
        .PivotItems("252204314008").Visible = False
        .PivotItems("252240109011").Visible = False
        .PivotItems("252272808005").Visible = False
        .PivotItems("262200130078").Visible = False
        .PivotItems("262285521091").Visible = False
        .PivotItems("282206502596").Visible = False
        .PivotItems("451000035931").Visible = False
        .PivotItems("451210045251").Visible = False
        .PivotItems("451210177141").Visible = False
        .PivotItems("451210498891").Visible = False
        .PivotItems("451210788004").Visible = False
        .PivotItems("451213056931").Visible = False
        .PivotItems("451213118611").Visible = False
        .PivotItems("451213406061").Visible = False
        .PivotItems("451213435751").Visible = False
        .PivotItems("451214843172").Visible = False
        .PivotItems("451220106902").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("451220109901").Visible = False
        .PivotItems("451291694535").Visible = False
        .PivotItems("451291768093").Visible = False
        .PivotItems("451291768123").Visible = False
        .PivotItems("451980042911").Visible = False
        .PivotItems("451980044391").Visible = False
        .PivotItems("452201101011").Visible = False
        .PivotItems("452205900151").Visible = False
        .PivotItems("452205901511").Visible = False
        .PivotItems("452209000971").Visible = False
        .PivotItems("452209004501").Visible = False
        .PivotItems("452209004841").Visible = False
        .PivotItems("452209008352").Visible = False
        .PivotItems("452209014208").Visible = False
        .PivotItems("452209017181").Visible = False
        .PivotItems("452209018231").Visible = False
        .PivotItems("452209018241").Visible = False
        .PivotItems("452209024951").Visible = False
        .PivotItems("452210280082").Visible = False
        .PivotItems("452210280102").Visible = False
        .PivotItems("452210280122").Visible = False
        .PivotItems("452210280142").Visible = False
        .PivotItems("452210355241").Visible = False
        .PivotItems("452210457286").Visible = False
        .PivotItems("452210459405").Visible = False
        .PivotItems("452210459421").Visible = False
        .PivotItems("452210459453").Visible = False
        .PivotItems("452210466113").Visible = False
        .PivotItems("452210601902").Visible = False
        .PivotItems("452212624336").Visible = False
        .PivotItems("452212650053").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452212650054").Visible = False
        .PivotItems("452212672141").Visible = False
        .PivotItems("452212702938").Visible = False
        .PivotItems("452212825601").Visible = False
        .PivotItems("452212857534").Visible = False
        .PivotItems("452212876313").Visible = False
        .PivotItems("452212876642").Visible = False
        .PivotItems("452212876681").Visible = False
        .PivotItems("452212876761").Visible = False
        .PivotItems("452212876781").Visible = False
        .PivotItems("452212876791").Visible = False
        .PivotItems("452212905322").Visible = False
        .PivotItems("452212905782").Visible = False
        .PivotItems("452212906782").Visible = False
        .PivotItems("452213171001").Visible = False
        .PivotItems("452214239252").Visible = False
        .PivotItems("452216422653").Visible = False
        .PivotItems("452216424613").Visible = False
        .PivotItems("452216424614").Visible = False
        .PivotItems("452216424624").Visible = False
        .PivotItems("452216424625").Visible = False
        .PivotItems("452216424661").Visible = False
        .PivotItems("452216424962").Visible = False
        .PivotItems("452216501133").Visible = False
        .PivotItems("452216501143").Visible = False
        .PivotItems("452216502181").Visible = False
        .PivotItems("452216503993").Visible = False
        .PivotItems("452216505961").Visible = False
        .PivotItems("452216506872").Visible = False
        .PivotItems("452216506882").Visible = False
        .PivotItems("452216507203").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452216507204").Visible = False
        .PivotItems("452216507249").Visible = False
        .PivotItems("452216508221").Visible = False
        .PivotItems("452216508231").Visible = False
        .PivotItems("452216508292").Visible = False
        .PivotItems("452216508451").Visible = False
        .PivotItems("452221042312").Visible = False
        .PivotItems("452230005332").Visible = False
        .PivotItems("452230007341").Visible = False
        .PivotItems("452230007351").Visible = False
        .PivotItems("452230007361").Visible = False
        .PivotItems("452230009052").Visible = False
        .PivotItems("452230009062").Visible = False
        .PivotItems("452230009181").Visible = False
        .PivotItems("452230014021").Visible = False
        .PivotItems("452230014031").Visible = False
        .PivotItems("452230014042").Visible = False
        .PivotItems("452230014043").Visible = False
        .PivotItems("452230014122").Visible = False
        .PivotItems("452230014123").Visible = False
        .PivotItems("452230014191").Visible = False
        .PivotItems("452230014292").Visible = False
        .PivotItems("452230014301").Visible = False
        .PivotItems("452230014342").Visible = False
        .PivotItems("452230014521").Visible = False
        .PivotItems("452230014661").Visible = False
        .PivotItems("452230014894").Visible = False
        .PivotItems("452230014906").Visible = False
        .PivotItems("452230014971").Visible = False
        .PivotItems("452230016714").Visible = False
        .PivotItems("452230019854").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452230019872").Visible = False
        .PivotItems("452230019881").Visible = False
        .PivotItems("452230019891").Visible = False
        .PivotItems("452230019921").Visible = False
        .PivotItems("452230019931").Visible = False
        .PivotItems("452230019941").Visible = False
        .PivotItems("452230019952").Visible = False
        .PivotItems("452230019971").Visible = False
        .PivotItems("452230019981").Visible = False
        .PivotItems("452230019991").Visible = False
        .PivotItems("452230026001").Visible = False
        .PivotItems("452230026021").Visible = False
        .PivotItems("452230026031").Visible = False
        .PivotItems("452230026041").Visible = False
        .PivotItems("452230026051").Visible = False
        .PivotItems("452230026061").Visible = False
        .PivotItems("452230026062").Visible = False
        .PivotItems("452230026101").Visible = False
        .PivotItems("452230026172").Visible = False
        .PivotItems("452230026173").Visible = False
        .PivotItems("452230026211").Visible = False
        .PivotItems("452230026221").Visible = False
        .PivotItems("452230026251").Visible = False
        .PivotItems("452230028542").Visible = False
        .PivotItems("452230028571").Visible = False
        .PivotItems("452230028581").Visible = False
        .PivotItems("452230028721").Visible = False
        .PivotItems("452230028921").Visible = False
        .PivotItems("452230028931").Visible = False
        .PivotItems("452230028941").Visible = False
        .PivotItems("452230029041").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("452230029081").Visible = False
        .PivotItems("452230029101").Visible = False
        .PivotItems("452230029111").Visible = False
        .PivotItems("452230029132").Visible = False
        .PivotItems("452230029153").Visible = False
        .PivotItems("452230029154").Visible = False
        .PivotItems("452230029161").Visible = False
        .PivotItems("452230029172").Visible = False
        .PivotItems("452250066171").Visible = False
        .PivotItems("452250066241").Visible = False
        .PivotItems("452298037862").Visible = False
        .PivotItems("452298037871").Visible = False
        .PivotItems("452298037881").Visible = False
        .PivotItems("452298037891").Visible = False
        .PivotItems("452298038081").Visible = False
        .PivotItems("452298038352").Visible = False
        .PivotItems("452298038422").Visible = False
        .PivotItems("452298038603").Visible = False
        .PivotItems("453560087111").Visible = False
        .PivotItems("453561153714").Visible = False
        .PivotItems("453561153724").Visible = False
        .PivotItems("453561158311").Visible = False
        .PivotItems("453561219221").Visible = False
        .PivotItems("453564260891").Visible = False
        .PivotItems("453564347391").Visible = False
        .PivotItems("453566440751").Visible = False
        .PivotItems("453567914251").Visible = False
        .PivotItems("453580488115").Visible = False
        .PivotItems("453580488116").Visible = False
        .PivotItems("455300002861").Visible = False
        .PivotItems("459800000561").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800000573").Visible = False
        .PivotItems("459800001091").Visible = False
        .PivotItems("459800001443").Visible = False
        .PivotItems("459800001444").Visible = False
        .PivotItems("459800001445").Visible = False
        .PivotItems("459800001681").Visible = False
        .PivotItems("459800002531").Visible = False
        .PivotItems("459800002541").Visible = False
        .PivotItems("459800002572").Visible = False
        .PivotItems("459800002573").Visible = False
        .PivotItems("459800002582").Visible = False
        .PivotItems("459800002591").Visible = False
        .PivotItems("459800002601").Visible = False
        .PivotItems("459800002612").Visible = False
        .PivotItems("459800002613").Visible = False
        .PivotItems("459800002623").Visible = False
        .PivotItems("459800002625").Visible = False
        .PivotItems("459800002631").Visible = False
        .PivotItems("459800003061").Visible = False
        .PivotItems("459800003142").Visible = False
        .PivotItems("459800011751").Visible = False
        .PivotItems("459800011761").Visible = False
        .PivotItems("459800011771").Visible = False
        .PivotItems("459800012171").Visible = False
        .PivotItems("459800013612").Visible = False
        .PivotItems("459800013631").Visible = False
        .PivotItems("459800020611").Visible = False
        .PivotItems("459800034242").Visible = False
        .PivotItems("459800037342").Visible = False
        .PivotItems("459800048691").Visible = False
        .PivotItems("459800061991").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800062051").Visible = False
        .PivotItems("459800062091").Visible = False
        .PivotItems("459800063241").Visible = False
        .PivotItems("459800066331").Visible = False
        .PivotItems("459800066333").Visible = False
        .PivotItems("459800066334").Visible = False
        .PivotItems("459800066335").Visible = False
        .PivotItems("459800066611").Visible = False
        .PivotItems("459800070281").Visible = False
        .PivotItems("459800070291").Visible = False
        .PivotItems("459800072661").Visible = False
        .PivotItems("459800072671").Visible = False
        .PivotItems("459800099201").Visible = False
        .PivotItems("459800108861").Visible = False
        .PivotItems("459800125171").Visible = False
        .PivotItems("459800125172").Visible = False
        .PivotItems("459800128372").Visible = False
        .PivotItems("459800148821").Visible = False
        .PivotItems("459800151382").Visible = False
        .PivotItems("459800151421").Visible = False
        .PivotItems("459800151422").Visible = False
        .PivotItems("459800151442").Visible = False
        .PivotItems("459800153441").Visible = False
        .PivotItems("459800155441").Visible = False
        .PivotItems("459800162341").Visible = False
        .PivotItems("459800164611").Visible = False
        .PivotItems("459800173431").Visible = False
        .PivotItems("459800173432").Visible = False
        .PivotItems("459800196181").Visible = False
        .PivotItems("459800219711").Visible = False
        .PivotItems("459800220541").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("459800238711").Visible = False
        .PivotItems("459800240611").Visible = False
        .PivotItems("459800240621").Visible = False
        .PivotItems("459800240681").Visible = False
        .PivotItems("459800240683").Visible = False
        .PivotItems("459800240691").Visible = False
        .PivotItems("459800240731").Visible = False
        .PivotItems("459800240741").Visible = False
        .PivotItems("459800240781").Visible = False
        .PivotItems("459800240801").Visible = False
        .PivotItems("459800240821").Visible = False
        .PivotItems("459800240841").Visible = False
        .PivotItems("459800240961").Visible = False
        .PivotItems("459800260121").Visible = False
        .PivotItems("459800267151").Visible = False
        .PivotItems("459800274261").Visible = False
        .PivotItems("459800295301").Visible = False
        .PivotItems("459800319211").Visible = False
        .PivotItems("459800319212").Visible = False
        .PivotItems("459800320161").Visible = False
        .PivotItems("459800359091").Visible = False
        .PivotItems("459800359092").Visible = False
        .PivotItems("459800372151").Visible = False
        .PivotItems("459800418511").Visible = False
        .PivotItems("459800440311").Visible = False
        .PivotItems("459800609732").Visible = False
        .PivotItems("459800671421").Visible = False
        .PivotItems("459800766581").Visible = False
        .PivotItems("867000053429").Visible = False
        .PivotItems("929900059707").Visible = False
        .PivotItems("989600007772").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC")
        .PivotItems("989600008501").Visible = False
        .PivotItems("989600180815").Visible = False
        .PivotItems("989600193001").Visible = False
        .PivotItems("989600204612").Visible = False
        .PivotItems("989600206924").Visible = False
        .PivotItems("989600216801").Visible = False
        .PivotItems("989601000652").Visible = False
        .PivotItems("989601023201").Visible = False
        .PivotItems("989601041312").Visible = False
        .PivotItems("989601041313").Visible = False
        .PivotItems("989601063621").Visible = False
        .PivotItems("989601065321").Visible = False
        .PivotItems("989670000011").Visible = False
        .PivotItems("989710002291").Visible = False
        .PivotItems("989710005263").Visible = False
        .PivotItems("989710006151").Visible = False
        .PivotItems("991920050193").Visible = False
        .PivotItems("991920050194").Visible = False
        .PivotItems("991920160462").Visible = False
        .PivotItems("991932050882").Visible = False
        .PivotItems("991932050883").Visible = False
        .PivotItems("991932050912").Visible = False
        .PivotItems("991932050913").Visible = False
        .PivotItems("991932050923").Visible = False
        .PivotItems("991932051114").Visible = False
        .PivotItems("991932212002").Visible = False
        .PivotItems("991932212011").Visible = False
        .PivotItems("991932472041").Visible = False
        .PivotItems("All Aggregated").Visible = False
        '.PivotItems("(blank)").Visible = False
    End With
    Range("C4").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"
    Range("D4").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
    Range("B6").Select
    Columns("A:D").EntireColumn.AutoFit
    Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
    ActiveWindow.SmallScroll ToRight:=16
    Windows("Veradius_Aug_2015_Jun_2013.xlsx").Activate
    ActiveSheet.PivotTables("PivotTable1").Location = _
        "'[KPI''s_NewVer_1.0_Change_R.xlsm]MTTR'!$AK$3"
        Windows("KPI's_NewVer_1.0_Change_R.xlsm").Activate
        Sheets("MTTR").Activate
    'ActiveSheet.PivotTables("PivotTable1").PivotSelect "Period", xlButton, True
    'ActiveSheet.PivotTables("PivotTable1").Location = "MTTR!$AK$3"
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[5]"
    Range("AF3").Select
    Selection.Copy
    
    Range("AF3,AF91").Select
    
    Range("AF3,AF3:AJ91").Select
    ActiveSheet.Paste
    
    Range("AE3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    Range("AE3").Select
    Selection.Copy
    
    Range("AE3:AE91").Select
    ActiveSheet.Paste
    
End Sub


'""""""""""""""""""""""""""""""""""""""""""""""""""""
'                       RRR Code
'""""""""""""""""""""""""""""""""""""""""""""""""""""

Public myWorkBook, wbName1 As String
Sub ListSubfoldersFile()
    Dim StrFile, myPath As String
    Dim objFSO, destRow As Long
    Dim mainFolder, mySubFolder
    Dim wbName As String
    Dim MyFiles(), DirArr() As String
    Dim FNum As Long
    Dim mybook As Workbook
    Dim wb As Workbook
    Dim BaseWks As Worksheet
    Dim ws As Worksheet
    Dim CalcMode, rowCount, baseItemCount As Long
    Dim cnt As Integer
    Dim lastRowCnt1 As Range
    Dim strDate As String
    Dim strDate1 As String
    Dim fstadd As String
    Dim lstadd1 As String
    Dim lstadd2 As String
    Dim firstAdd, lastAdd As String
    Dim outPutFileFrstAdd As String
    Dim outPutFileLstAdd As String
 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myPath = "D:\Com_cS\KPI Automation\CTS\Input Source"
    Set mainFolder = objFSO.GetFolder(myPath)
    StrFile = Dir(myPath & "*.xlsx*")
    Set BaseWks = ThisWorkbook.Worksheets(1)
    Workbooks.Add
    wbName = ActiveWorkbook.name
    Workbooks(wbName).Activate
    Sheets(1).Activate
    Sheets(1).Range("A1").value = "Period"
        
    cnt = 1
    For Each mySubFolder In mainFolder.SubFolders
        StrFile = Dir(mySubFolder & "\*.xls*")
        Do While Len(StrFile) > 0
            FNum = FNum + 1
            ReDim Preserve MyFiles(1 To FNum)
            ReDim Preserve DirArr(1 To FNum)
            MyFiles(FNum) = StrFile

            Set wb = Workbooks.Open(mySubFolder & "\" & StrFile)
            Set ws = wb.Sheets(1) 'uses first sheet or if all the same names then ws.Sheets("yoursheet")
            Workbooks(StrFile).Activate
            wb.Sheets(1).Activate
            Set lastRowCnt1 = Workbooks(StrFile).Sheets(1).Range("N" & Range("A" & rows.Count).End(xlUp).Row)
            lstadd1 = lastRowCnt1.Address
            
            ActiveSheet.Cells(3, 1).Select
    
            If cnt <= 1 Then
                fstadd = ActiveCell.Address
            Else
                fstadd = ActiveCell.Offset(1, 0).Address
            End If
                strDate = Cells(1, 1).value
                strDate1 = Mid(strDate, 43, 8)
                ActiveSheet.Range(fstadd, lstadd1).Select
                Selection.Copy
                Workbooks(wbName).Activate
                Sheets(1).Activate
                Range("B1").Select
    
           If Range("B1").value = "" Then
               Range("B1").PasteSpecial xlPasteValues
               
           Else
               ActiveCell.End(xlDown).Select
               lstadd2 = ActiveCell.Offset(1, 0).Address
               Range(lstadd2).PasteSpecial xlPasteAllExceptBorders
               Range("B1").Select
               ActiveCell.End(xlDown).Select
               ActiveCell.End(xlToLeft).Select

           End If

          If Range("A2").value = "" Then
              Range("A2").Select
              firstAdd = ActiveCell.Address
              Range("B1").Select
              ActiveCell.End(xlDown).Select
              ActiveCell.End(xlToLeft).Select
              lastAdd = ActiveCell.Address
              Range(firstAdd, lastAdd).value = strDate1
         Else
              Range("A2").Select
              ActiveCell.End(xlDown).Select
              firstAdd = ActiveCell.Address
              Range("B1").Select
              ActiveCell.End(xlDown).Select
              ActiveCell.End(xlToLeft).Select
              lastAdd = ActiveCell.Address
              Range(firstAdd, lastAdd) = strDate1

         End If
              Application.DisplayAlerts = False

              wb.Close True
              ActiveCell.Offset(1, 0).Select
            
              StrFile = Dir
              cnt = cnt + 1
        Loop
        
    Next mySubFolder
    
    Workbooks(wbName).Activate
    Sheets(1).Activate
    ActiveSheet.name = "MergedCCC1"
    ActiveSheet.Range("A1").Select
    outPutFileFrstAdd = ActiveCell.Address
    ActiveSheet.Range("A1").End(xlToRight).Select
    ActiveSheet.Range("A1").End(xlDown).Select
    outPutFileLstAdd = ActiveCell.Address
    ActiveSheet.Range(outPutFileFrstAdd, outPutFileLstAdd).RowHeight = 15

    Range("A2").Select
    outPutFileFrstAdd = ActiveCell.Address
    Range("A2").End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(0, 7).Select
    outPutFileLstAdd = ActiveCell.Address
    ActiveSheet.Range(outPutFileFrstAdd, outPutFileLstAdd).Select

    ActiveWorkbook.Worksheets("MergedCCC1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MergedCCC1").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ActiveWorkbook.Worksheets("MergedCCC1").Sort
        .SetRange Range(outPutFileFrstAdd, outPutFileLstAdd)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    Dim rows1 As Range, cell As Range, value As Double

    Range("A1").Select
    firstAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lastAdd = ActiveCell.Address

    Range(firstAdd, lastAdd).Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range(firstAdd, lastAdd).AutoFilter Field:=8, Criteria1:=Array( _
        "722400", "718094", "718095", "718400", "718074", _
        "714045", "712301", "712301", "714047", "714247", _
        "714048", "714248", "704301", "704310", "712310", _
        "718132", "718131", "718130", "718075"), Operator:=xlFilterValues
         
    ActiveSheet.Range(firstAdd, lastAdd).AutoFilter Field:=10, Criteria1:=Array( _
        "C", "W"), Operator:=xlFilterValues
    
    Dim add1
        Range(firstAdd, lastAdd).Select
        Range("A1").Select
        firstAdd = ActiveCell.Address
        ActiveCell.SpecialCells(xlCellTypeLastCell).Select
        lastAdd = ActiveCell.Address
        Range(firstAdd, lastAdd).Select
        Selection.Copy
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "RRR_Report"

        Sheets("RRR_Report").Activate
        ActiveSheet.Cells(1, 1).Select
        ActiveCell.PasteSpecial xlPasteAll
        'Sheets("Sheet2").Name = "MergedCCC"
        Selection.RowHeight = 15
        Sheets("RRR_Report").Activate
        Sheets("RRR_Report").UsedRange.Find(what:="RR", lookat:=xlWhole).Select
        thrdLstClmn = ActiveCell.Offset(0, 1).Address
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(0, 3).Select
        lstclmn = ActiveCell.Address
    ActiveSheet.Range(thrdLstClmn, lstclmn).Select
    Selection.EntireColumn.Delete

    Sheets("RRR_Report").Activate
    ActiveSheet.Range(thrdLstClmn, lstclmn).Select
    Selection.EntireColumn.Delete
    'Workbooks(wbName).Sheets(1).Name = "RR"
    Workbooks(wbName).SaveAs _
    fileName:=ThisWorkbook.Path & "\" & "CTS-Cost to Serve_RRR_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    Dim srcfile As String
    srcfile = ThisWorkbook.Path & "\" & "CTS-Cost to Serve_RRR_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    Application.Workbooks.Open(srcfile).Activate
    Worksheets("MergedCCC1").Activate
    wbName1 = ActiveWorkbook.name
    
    outPutFileVlookup
    
    Workbooks(wbName1).Activate
    ActiveWorkbook.Save
    MsgBox "RRR Data is Generated succesfully", vbOKOnly
    'Workbooks(wbName1).Activate
    'Sheets("Last3Clmns of MergedCCC").Activate
    'Range("A1:C1").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Application.CutCopyMode = False
    'Selection.Copy
    'Sheets("MergedCCC").Activate
    'Sheets("MergedCCC").UsedRange.Find(what:="RR", lookat:=xlWhole).Select
    'ActiveCell.Offset(0, 1).EntireColumn.Select
'    Columns("M:M").Select
    'Selection.Insert Shift:=xlToRight
    Application.ScreenUpdating = True
    
    Application.DisplayAlerts = True

End Sub

Sub outPutFileVlookup()
    Dim StrFile, myPath, myFile As String
    Dim objFSO, destRow As Long
    Dim wb As Workbook
    Dim inputItem As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myPath = "D:\Com_cS\KPI Automation\CTS\Input Source\"
    myFile = "BCTool_SWO_RawData_SingleVersionOfTheTruth"
    inputItem = myPath & "\" & Dir(myPath & "\" & "BCTool_SWO_RawData_SingleVersionOfTheTruth" & "*.xls*") 'input file path
    Application.Workbooks.Open (inputItem), False
    myWorkBook = ActiveWorkbook.name
    
    Workbooks(myWorkBook).Activate
    With ThisWorkbook
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "CombinedFDV"
    End With

    ActiveWorkbook.Sheets("SWO CV").Activate
    ActiveSheet.Cells(1, 2).Select
    firstAdd = ActiveCell.Address
    ActiveCell.SpecialCells(xlCellTypeLastCell).Select
    lastAdd = ActiveCell.Address
    Range(firstAdd, lastAdd).Select
    Selection.Copy
    Sheets("CombinedFDV").Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveCell.PasteSpecial xlPasteAll
        
    ActiveWorkbook.Sheets("SWO GXR").Activate
    ActiveSheet.Cells(2, 2).Select
    firstAdd = ActiveCell.Address
    ActiveCell.SpecialCells(xlCellTypeLastCell).Select
    lastAdd = ActiveCell.Address
    Range(firstAdd, lastAdd).Select
    Selection.Copy
    Sheets("CombinedFDV").Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveCell.End(xlDown).Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteAll

Sheets("CombinedFDV").Activate
Cells(1, 1).Select
Do Until ActiveCell.value = ""

If ActiveCell.value = "SWO" Or ActiveCell.value = "CallPeriodTECO" Or ActiveCell.value = "Entitlement" Or ActiveCell.value = "ContractMaterial" _
Or ActiveCell.value = "SystemCode" Or ActiveCell.value = "SystemName" Or ActiveCell.value = "Market" Or ActiveCell.value = "ETTRValue" _
Or ActiveCell.value = "SubSystem" Or ActiveCell.value = "BuildingBlock" Or ActiveCell.value = "Part12NC" Or ActiveCell.value = "MaterialDescription" _
Or ActiveCell.value = "CustomerComplaintSubject" Or ActiveCell.value = "CustomerComplaint" Or ActiveCell.value = "JobRepairText" _
Or ActiveCell.value = "JobCustomerRepairText" Or ActiveCell.value = "RR" Or ActiveCell.value = "FTF" Then
ActiveCell.Offset(0, 1).Select
Else
ActiveCell.EntireColumn.Delete
End If

Loop

'Workbooks(myWorkBook).Activate
'Sheets("CombinedFDV").Activate
Cells(1, 1).Select
fstClmnAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select
lstClmnAd = ActiveCell.Address
ActiveCell.End(xlDown).Select
lstRowAdd = ActiveCell.Address
ActiveSheet.Range(fstClmnAdd, lstRowAdd).RemoveDuplicates Columns:=1, Header:=xlNo
Cells(1, 1).Select

    Workbooks(wbName1).Activate
    Sheets("RRR_Report").Activate
    ActiveSheet.UsedRange.Find(what:="SWO", lookat:=xlWhole).Select
    
    Do Until ActiveCell.value = ""
    
    RRRSwoFstCellAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "" Then
    ActiveCell.EntireRow.Delete
    End If
    Loop
    
    ActiveSheet.UsedRange.Find(what:="SWO", lookat:=xlWhole).Select

    Dim findVal As String
    On Error Resume Next
    Do Until ActiveCell.value = ""
        ActiveCell.Offset(1, 0).Select
        findVal = ActiveCell.value
        findValAdd = ActiveCell.Address
        If Not Workbooks(myWorkBook).Sheets("CombinedFDV").UsedRange.Find(what:=findVal, lookat:=xlWhole) Is Nothing Then
            Workbooks(myWorkBook).Activate
            Sheets("CombinedFDV").Activate
            Workbooks(myWorkBook).Sheets("CombinedFDV").UsedRange.Find(what:=findVal, lookat:=xlWhole).Select
            fndValAdd = ActiveCell.Address
            ActiveSheet.Range(fndValAdd).Select
    
            Dim fstAddForCopy As String
            fstAddForCopy = ActiveCell.Address
    
            Dim lstAddForCopy As String
            ActiveCell.Offset(0, Columns.Count - 1).Select
            ActiveCell.End(xlToLeft).Select
            lstAddForCopy = ActiveCell.Address
            ActiveSheet.Range(fndValAdd, lstAddForCopy).Select
            Selection.Copy
            Workbooks(wbName1).Sheets("RRR_Report").Activate
            ActiveCell.End(xlToRight).Select
            ActiveCell.Offset(0, 1).Select
            Selection.PasteSpecial xlPasteAll
            ActiveCell.Offset(0, -10).Select
                
          Else

        End If
    Loop

Workbooks(myWorkBook).Activate
Sheets("CombinedFDV").Activate
Cells(1, 1).Select
fstClmnAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select
lstClmnAd = ActiveCell.Address
ActiveSheet.Range(fstClmnAdd, lstClmnAd).Select
Selection.Copy
Workbooks(wbName1).Activate
Sheets("RRR_Report").Activate
Cells(1, 1).Select
fstClmnAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.PasteSpecial xlPasteAll
ActiveWorkbook.Save

End Sub
Sub createPivotTableRRRData()

Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim rngData As String
Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim strtPt As String
Dim SrcData As String
Dim wsData As Worksheet
Dim wsPtTable As Worksheet
Dim pvtExcel As String
Dim wsptName  As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String

'pt.ManualUpdate = False
    myPath = "D:\Com_cS\KPI Automation\CTS\Input Source"
    pvtExcel = myPath & "\" & Dir(myPath & "\" & "CTS-Cost to Serve_RRR_" & "*.xls*")  'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    Sheets("RRR_Report").Activate
    Range("A1").Select
    pkAdd = ActiveCell.Address
    fstCellAdd = ActiveCell.Address(ReferenceStyle:=R1C1)
    Range("A1").End(xlToRight).Select
    lastClmnAdd = ActiveCell.Address
    mioflstcell = Left(lastClmnAdd, 4)
    Range("A1").Select
    ActiveCell.End(xlDown).Select
    lstRowAdd = ActiveCell.Address
    midoflstadd = Mid(lstRowAdd, 4)
    Add = mioflstcell & midoflstadd
    ActiveSheet.Range(Add).Select
    addofLstClmn = ActiveCell.Address(ReferenceStyle:=R1C1)
    Sheets.Add After:=Worksheets(Worksheets.Count)

    Set wsPtTable = Worksheets(Sheets.Count)

    Set wsPtTable = Worksheets(3)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets("RRR_Report").Activate

    Set wsData = Worksheets("RRR_Report")
    Worksheets("RRR_Report").Activate
    sourceSheet = ActiveSheet.name

    Sheets(wsptName).Activate
    rngData = fstCellAdd & ":" & addofLstClmn
    
    Workbooks(myPvtWorkBook).Connections.Add2 _
        "WorksheetConnection_" & "RRR_Report" & "!" & rngData, "", _
        "WORKSHEET;" & myPath & "\[myPvtWorkBook]RRR_Report" _
        , "RRR_Report!" & pkAdd & ":" & Add, 7, True, False

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("WorksheetConnection_RRR_Report!" & rngData), _
        Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
        wsptName & "!R3C1", TableName:="PivotTable1", DefaultVersion:= _
        xlPivotTableVersion15
       
        wsPtTable.Activate
        
     With ActiveSheet.PivotTables("PivotTable1").CubeFields("[Range].[Entitlement]")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").CubeFields("[Range].[RR]")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").CubeFields( _
        "[Range].[BuildingBlock]")
        .Orientation = xlPageField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").CubeFields("[Range].[SubSystem]")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").CubeFields( _
        "[Range].[MaterialDescription]")
        .Orientation = xlRowField
        .Position = 2
    End With
     
     ActiveSheet.PivotTables("PivotTable1").CubeFields.GetMeasure "[Range].[SWO]", _
        xlCount, "Count of SWO"
     ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").CubeFields("[Measures].[Count of SWO]"), "Count of SWO"
 
     ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[SubSystem].[SubSystem]").AutoSort xlDescending, _
        "[Measures].[Count of SWO]", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1
      
     ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[SubSystem].[SubSystem]").DrilledDown = False
  
     ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").AutoSort xlDescending, _
        "[Measures].[Count of SWO]", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1
     
    ActiveWorkbook.RefreshAll
    
    ActiveSheet.PivotTables("PivotTable1").CubeFields(12).EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[100]")
    
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlValue, True
    Selection.Copy

    ActiveSheet.Range("D6").PasteSpecial

    ActiveSheet.PivotTables("PivotTable1").ClearAllFilters

    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[0]")

    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlValue, True
    Selection.Copy

    ActiveSheet.Range("E6").PasteSpecial
    ActiveSheet.PivotTables("PivotTable1").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlValue, True
    Selection.Copy

    ActiveSheet.Range("C6").PasteSpecial
    ActiveSheet.Range("C5").value = "Total"
    ActiveSheet.Range("D5").value = "RR"
    ActiveSheet.Range("E5").value = "Non-RR"
   
   ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").AutoSort xlDescending, _
        "[Measures].[Count of SWO]", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1
              
        ActiveWorkbook.Save

    MsgBox "Pivot Table of RRR Data is created succesfully", vbOKOnly

End Sub

