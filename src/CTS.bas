Attribute VB_Name = "CTS"
Public myWorkBook As String
Public wsptName  As String
Public wbName1 As String
Public srcfile As String
Public Sub createPivotTableAggregatedDataKPIALL()

Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim pvtItm As PivotItem

Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim wsData As Worksheet
Dim wsPtTable As Worksheet

Dim rngData As String
Dim pvtExcel As String
Dim strtPt As String
Dim SrcData As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer

'Case select for sheet tab
    KPISheetName = Sheet1.comb6NC1.value

    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = Sheet1.comb6NC1.value
    dateValue = Sheet1.combYear.value
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    Exit Sub
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    myPath = ThisWorkbook.Path
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
     

'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
    Sheets("Aggr. SWO Data CV").Activate
    Dim l As Long
    l = Application.WorksheetFunction.Match("BuildingBlock", Range("2:2"), 0)
    Range("A2").Select
    fstCellAdd = ActiveCell.Address
    Range("A2").End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    Selection.AutoFilter
    ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=l, Criteria1:="=Buildingblocks Aggregated"
    Range("A2").Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range("A2").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
    Range(fstFiltCellAdd1).End(xlToRight).Select
    fstFiltCellAdd2 = ActiveCell.Address
   ' lastFiltCellAdd = ActiveCell.Address
    Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
    ActiveSheet.Range("G1").Select
    Selection.UnMerge
    ActiveSheet.Range("A1").Select
    Selection.UnMerge
            
       Dim Part12NCClmn As Range
       Set Part12NCClmn = Sheets("Aggr. SWO Data CV").rows(2).Find("Part12NC", , , xlWhole, , , , False)
    
      If Not Part12NCClmn Is Nothing Then
        Application.ScreenUpdating = False
        Part12NCClmn.Offset(1, 0).Select
        Part12NcClmnAdd = ActiveCell.Address(False, False)
      End If
        
        
       Dim ttlCalls As Range
       Set ttlCalls = Sheets("Aggr. SWO Data CV").rows(2).Find("Total Calls (#)", , , xlWhole, , , , False)
    
      If Not ttlCalls Is Nothing Then
        Application.ScreenUpdating = False
        ttlCalls.Offset(1, 0).Select
        ttlCallsAdd = ActiveCell.Address(False, False)
      End If
        
       Dim AvgMTTRprCallHrs As Range
       Set AvgMTTRprCallHrs = Sheets("Aggr. SWO Data CV").rows(2).Find("Avg. MTTR/Call (hrs)", , , xlWhole, , , , False)
    
      If Not AvgMTTRprCallHrs Is Nothing Then
        Application.ScreenUpdating = False
        AvgMTTRprCallHrs.Offset(1, 0).Select
        AvgMTTRprCallHrsAdd = ActiveCell.Address(False, False)
      End If
            
       Dim visitsprCallNP As Range
       Set visitsprCallNP = Sheets("Aggr. SWO Data CV").rows(2).Find("# of calls with 1 visit", , , xlWhole, , , , False)
    
      If Not visitsprCallNP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallNP.Offset(1, 0).Select
        visitsprCallNPAdd = ActiveCell.Address(False, False)
      End If
      
       Dim visitsprCallP As Range
       Set visitsprCallP = Sheets("Aggr. SWO Data CV").rows(2).Find("Calls = 0 Visit", , , xlWhole, , , , False)
    
      If Not visitsprCallP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallP.Offset(1, 0).Select
        visitsprCallPAdd = ActiveCell.Address(False, False)
      End If
      
      
      'Add one column for "Total Cost of Parts & Non-Parts"

  Dim found As Range
  Set found = Sheets("Aggr. SWO Data CV").rows(2).Find("Total Costs/part (EUR)", , , xlWhole, , , , False)
    
    If Not found Is Nothing Then
        Application.ScreenUpdating = False
        found.Offset(, 1).Resize(, 1).EntireColumn.Insert
  
  End If
  
        Workbooks(myPvtWorkBook).Sheets("Aggr. SWO Data CV").Activate
      ' ActiveSheet.ClearAllFilters

        found.End(xlDown).Select
        ActiveCell.Offset(0, 1).Select
        ttlCstLstAdd = ActiveCell.Address
        found.Offset(, 1).value = "Total Cost of Parts & Non-Parts"
        found.Offset(1, 1).Select
        ttlCstAdd = ActiveCell.Address
      
   '   Part12NcClmnAdd , ttlCallsAdd, AvgMTTRprCallHrsAdd, visitsprCallNPAdd, visitsprCallPAdd
      
       ActiveCell.Offset(, 0).Formula = "=IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallPAdd & "*" & 200 & ")," & "IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Non-Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallNPAdd & "*" & 200 & ")))"
      
      '=IF(E3="Parts Aggregated",(G3*H3*100)+(L3*200),IF(E3="Non-Parts Aggregated",(G3*H3*100)+(K3*200)))
      
      ' Range("N3", "N" & Cells(rows.Count, 1).End(xlUp).Row).FillDown
      Range(ttlCstAdd).Select
      Selection.Copy
      Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteFormulas
      Range(ttlCstAdd, ttlCstLstAdd).Select
      Selection.Copy
      Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteValues
      
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
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
    Range("A1").value = "Period"
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
         
     DataBrekUpFrPivotKPIALL
     
     'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)

        Set wsPtTable = Worksheets(Sheets.Count)

        'Set wsPtTable = Worksheets(3)
        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate

        Set wsData = Worksheets("Aggr. SWO Data CV")
        Worksheets("Aggr. SWO Data CV").Activate
        sourceSheet = ActiveSheet.name
        
        Cells(1, 1).Select
        
        fstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select

        lstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
        Sheets(wsptName).Activate
        rngData = fstadd & ":" & lstadd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
        
        Range(fstadd).Select
        ActiveCell.PivotTable.name = "pvtKPIALL"
        wsPtTable.Activate
        
        Set pt = wsPtTable.PivotTables("pvtKPIALL")
        Set pf = pt.PivotFields("Period")
        pf.Orientation = xlPageField
        pf.Position = 1
        
        Set pf = pt.PivotFields("SubSystem")
        pf.Orientation = xlRowField
        pf.Position = 1
        Set pf = pt.PivotFields("BuildingBlock")
        pf.Orientation = xlRowField
        pf.Position = 2
     
        Set pf = pt.PivotFields("Part12NC-Sub Parts")
        pf.Orientation = xlRowField
        pf.Position = 3
        
        Set pf = pt.PivotFields("PartDescription")
        pf.Orientation = xlRowField
        pf.Position = 4
        
        Set pf = pt.PivotFields("Part12NC")
        pf.Orientation = xlColumnField
        pf.Position = 1
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Total Calls (#)"), "# of Calls", xlSum
        
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Avg. MTTR/Call (hrs)"), "MTTR/Call (hrs)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Avg. ETTR (days)"), "ETTR (days)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Avg. Visits/call (#)"), "Visits/call (#)", xlSum
        
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
        
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Total Costs/part (EUR)"), "Costs/part (EUR)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Total Cost of Parts & Non-Parts"), _
        "#Total Cost of Parts & Non-Parts", xlSum
    
        With ActiveSheet.PivotTables("pvtKPIALL")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
        End With
    
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "", xlDataAndLabel, True
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("PartDescription"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 2
        End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtKPIALL")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtKPIALL")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
    Set pf = pvtTbl.PivotFields("Part12NC")

        pf.ClearAllFilters
        pf.EnableMultiplePageItems = True
    
    pf.PivotItems("Parts/Non-Parts Breakups").Visible = False
    ActiveSheet.PivotTables("pvtKPIALL").HasAutoFormat = False
    ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "", xlDataAndLabel, True
    Selection.ColumnWidth = 8
    ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "'Part12NC-Sub Parts'['-']" _
        , xlDataAndLabel, True
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts"). _
        ShowDetail = False
    Range("B4").Select
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail _
        = False
    
    With ActiveSheet.PivotTables("pvtKPIALL")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    pvtTbl.RefreshTable
' Add ConditionalFormatting of Data Bars on total calls of Parts and Non parts
    Columns("E:E").Select
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
    Columns("F:F").Select
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
'Add conditional formatting on MTTR and ETTR Calls
    Columns("G:G").Select
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
    Range("G23").Select
    
    Columns("H:H").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("H19").Select
    
    Columns("I:I").Select
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
    
    Columns("K:L").Select
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
    
    Columns("K:K").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("E:P").Select
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

    Columns("A:D").Select
    With Selection
        .ColumnWidth = 8
    End With
    Cells(1, 1).Select

    
    Worksheets(wsptName).PivotTables("pvtKPIALL").PreserveFormatting = False
    Sheets(wsptName).name = "PivotTableAggData"
  '  pt.ManualUpdate = True
ActiveWindow.Zoom = 85


        
    Range(fstadd, lstadd).Select
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
    Selection.NumberFormat = "0.000"
    Selection.FormatConditions(1).StopIfTrue = False
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
        Selection.EntireColumn.Select
        Selection.ColumnWidth = "8"
        
        
        Range("P5").Select
    Selection.Copy
    Range("Q5:S5").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S5").Select
    ActiveCell.FormulaR1C1 = "/Sys/Yr"
    Range("R5").Select
    Selection.Copy
    Range("R4:S4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("R4").Select
    Selection.Copy
    Range("Q3:Q4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Q4").Select
    Selection.Copy
    Range("R3:S3").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
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
    
    Range("Q3:Q4").Select
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
        
        
 ChDir "D:\Philips\CTS\Com_cS\KPI Automation\CTS\Input Source"
    Workbooks.Open fileName:= _
        "D:\Philips\CTS\Com_cS\KPI Automation\CTS\Input Source\CTS_KPI_Summary.xlsx"
    Sheets("KPI-All").Select
    
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("KPI-All").Select
    Cells.Select
    Selection.Delete
    Workbooks(AggPvtTableName).Activate
    pvtTbl.TableRange2.Copy
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("KPI-All").Select
    Range("a1").PasteSpecial
     
    'Selection.Paste
    'outPutFilePath = ThisWorkbook.Path & "\"
    'installFlName = outPutFilePath & "CTS_KPI_Summary.xlsx"
    'Application.Workbooks.Open (installFlName), False 'false to disable link update message
    'myWorkBook = ActiveWorkbook.name
   '
   ' Workbooks(AggPvtTableName).Activate
   ' Sheets("PivotTableAggData").Activate
    'pvtTbl.TableRange2.Copy
    
    'Workbooks(myWorkBook).Activate
    'Sheets("KPI-All").Activate
    'Range("A1").PasteSpecial
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtKPIALL"

    Workbooks(myWorkBook).Save
    
    
End Sub

Sub DataBrekUpFrPivotKPIALL()

Sheets("Aggr. SWO Data CV").Select

Cells(1, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(0, 1).Select
Selection.EntireColumn.Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(0, 1).Select
ActiveCell.value = "Part12NC-Sub Parts"

Application.CutCopyMode = False
ActiveCell.Offset(1, 0).Select
fstadd = ActiveCell.Address
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
lstadd = ActiveCell.Address
Cells(1, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(1, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""All Aggregated"",RC[-1]=""Parts Aggregated"",RC[-1]=""Non-Parts Aggregated""),""-"",RC[-1])"
    Selection.AutoFill Destination:=Range(fstadd, lstadd)
    Range(fstadd, lstadd).Select
    Calculate
    Range(fstadd, lstadd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(1, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    Selection.Offset(1, 0).Select
'    Range("F2").Select
   ' ActiveCell.Offset(0, -1).Select
    
'    ActiveCell.End(xlDown).Select
 '   ActiveCell.Offset(1, 0).Select
  '  lstCellAdd = ActiveCell.Address
   ' Cells(2, 1).Select
    'Selection.EntireRow.Select
    'Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    'Selection.Offset(1, 0).Select
    'fstCellAdd = ActiveCell.Address
    Do Until ActiveCell.value = ""
    If ActiveCell.value = "All Aggregated" Then
    ActiveCell.Offset(1, 0).Select
    End If
    If ActiveCell.value = "Parts Aggregated" Then
    Do Until ActiveCell.value = "Non-Parts Aggregated"
    ActiveCell.value = "Parts Aggregated"
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "" Then
    Exit Do
    End If
    
    Loop
    ElseIf ActiveCell.value = "Non-Parts Aggregated" Then
    Do Until ActiveCell.value = "Parts Aggregated"
    ActiveCell.value = "Non-Parts Aggregated"
    ActiveCell.Offset(1, 0).Select

    If ActiveCell.value = "" Then
    Exit Do
    End If
    Loop
    End If
   
    Loop
    ActiveCell.Offset(0, 1).Select
    If ActiveCell.value = 0 Then
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    End If
    Cells(1, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.value = ""
If ActiveCell.value = "-" Then
ActiveCell.Offset(1, 0).Select
Else
ActiveCell.Offset(0, -1).value = "Parts/Non-Parts Breakups"
ActiveCell.Offset(1, 0).Select

End If
Loop
End Sub
Public Sub CallRateCalculation()
Dim fixedDate, myPath, CTSExcel, CTSWorkBook, pvtExcel, myPvtWorkBook As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1, KPISheetName As String

'Case select for sheet tab
    KPISheetName = Sheet1.comb6NC1.value

    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = Sheet1.comb6NC1.value
    dateValue = Sheet1.combYear.value
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    Exit Sub
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    CTSProductName = KPISheetName
   
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
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
    outPutFilePath = ThisWorkbook.Path & "\"
    installFlName = outPutFilePath & "CTS_KPI_Summary.xlsx"
    Application.Workbooks.Open (installFlName), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Sheets("KPI-All").Select
    
    
    'Windows("CTS_KPI_Summary.xlsx").Activate
    'Cells.Select
    'Application.CutCopyMode = False
    'Selection.Delete Shift:=xlUp
    'Range("A1").Select
    'Windows("Veradius_2014-06.xlsx").Activate
    'Selection.Copy
    'Windows("CTS_KPI_Summary.xlsx").Activate
    'ActiveSheet.Paste
    fixedDate = Sheet1.combYear.value
endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters
With ActiveSheet.PivotTables("pvtKPIALL")
        .ColumnGrand = True
        .RowGrand = False
    End With
   ' pvtTbl.RefreshTable
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm
    
'        Workbooks("KPI's_NewVer_1.0_change_R").Activate
        Sheets("CR").Select
        Range("A:A").Select
        On Error Resume Next
        Selection.EntireRow.Select
        Selection.EntireRow.Delete
        Application.Columns.Ungroup

        Sheets("KPI-All").Select
        pvtTbl.TableRange1.Select
        pvtTbl.TableRange1.Copy
        
        Sheets("CR").Select
        Range("a1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("1:1").Select
        Selection.EntireRow.Delete
        Sheets("CR").UsedRange.Find(what:="MTTR/Call (hrs)", lookat:=xlWhole).Select
        ActiveCell.Offset(1, 0).Select
        fstclmn = ActiveCell.Address
        ActiveCell.End(xlToRight).Select
       lstclmnAdd = ActiveCell.Address
        Range(fstclmn, lstclmnAdd).Select
        Selection.EntireColumn.Select
        Selection.EntireColumn.Delete
        Sheets("CR").UsedRange.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
        delteteClmnfstAdd = ActiveCell.Address
        Sheets("CR").UsedRange.Find(what:="PartDescription", lookat:=xlWhole).Select
        delteteClmnlstAdd = ActiveCell.Address
        Range(delteteClmnfstAdd, delteteClmnlstAdd).Select
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
    
     myPath = ThisWorkbook.Path
    CTSExcel = myPath & "\" & Dir(myPath & "\" & "CTS_Guidelines.xlsx")  'input file path
    Application.Workbooks.Open (CTSExcel), False
    
    CTSWorkBook = ActiveWorkbook.name
    Workbooks(CTSWorkBook).Activate
    Sheets("Sheet2").Activate
    Sheets("Sheet2").UsedRange.Find(what:="CR / Sys / Yr", lookat:=xlWhole).Select

    Selection.EntireColumn.Select
    Selection.Copy
    
    Windows(myWorkBook).Activate
    Range("C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(CTSWorkBook).Close
    
    Windows(myWorkBook).Activate
    Range("A2").Select
    fstadd1 = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    lstadd2 = ActiveCell.Address
    Range(fstadd1, lstadd2).Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    
    Sheets("CR").Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C3").Select
    fstadd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    lstadd = ActiveCell.Address
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Selection.AutoFill Destination:=Range(fstadd, lstadd)
    Range(fstadd, lstadd).Select
    Calculate
    Cells(2, 3).value = "SS&BB"
    
    
    
  Sheets("CR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
  ActiveCell.Offset(1, 1).Select
    CRSysYrFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -4).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    CRSysYrLstAdd = ActiveCell.Address
    Range(CRSysYrFstAdd).Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]+RC[-2])/2047"
    Range(CRSysYrFstAdd).Select
    Selection.AutoFill Destination:=Range(CRSysYrFstAdd, CRSysYrLstAdd)
    Range(CRSysYrFstAdd, CRSysYrLstAdd).Select
    Range(CRSysYrFstAdd, CRSysYrLstAdd).NumberFormat = "0.00"

    Calculate
     
    
    
    
    
    
    'Columns("AE:AE").Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("AE3").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    'Range("AE3").Select
    'Selection.AutoFill Destination:=Range("AE3:AE91")
    'Range("AE3:AE91").Select
    'Calculate
    
        
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    Sheets("CR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    CRSysYrFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    CRSysYrLstAdd = ActiveCell.Address
    Range(CRSysYrFstAdd, CRSysYrLstAdd).Select
    
    'ActiveCell.Offset(1, 1).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "0.00"
    'Cells(3, 7).Select
    'ActiveCell.End(xlDown).Select
    'lstRowAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    'Range(lstRowAdd).Select
    'Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    'ActiveCell.EntireRow.Delete
    'Sheets("CR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    'ActiveCell.Offset(2, 0).Select
    'pkAdd = ActiveCell.Address
    'fstCellAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    'mioflstcell = Left(fstCellAdd, 3)
    'midoflstadd = Mid(lstRowAdd, 4)
    'Add = mioflstcell & midoflstadd
    'ActiveSheet.Range(fstCellAdd, Add).Select
        
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
    
    Sheets("CR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    CRSysYrFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    CRSysYrLstAdd = ActiveCell.Address
    Range(CRSysYrFstAdd, CRSysYrLstAdd).Select
    
    
        
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
    Sheets("CR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select

    ActiveCell.value = "MAT # of Calls profiles"
    ActiveCell.Offset(1, 0).Select

    ActiveCell.value = "Non-Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Parts"
    ActiveCell.Offset(0, 1).Select

    ActiveCell.value = "CR / Sys / ITM"
    ActiveCell.Offset(1, 0).Select

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
    ActiveCell.value = "Current Year Avg. CR / Sys"
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
    
    
    fixedDate1 = Sheet1.combYear.value
    frmtData = Format(DateAdd("m", 1, fixedDate1), "mmm" & "-" & "yyyy")
'currentdate = Format(Now(), "yyyymm")

   endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
   fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
'    j = 12
    Do Until frmEndDate = frmtData
    ActiveCell.value = frmEndDate
    ActiveCell.Offset(0, 1).Select
    frmEndDate = Format(DateAdd("m", 1, frmEndDate), "mmm" & "-" & "yyyy")
    Loop

    Range("A1").Select
    fstadd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstadd = ActiveCell.Address
    ActiveCell.Offset(-1, 0).Select
    upAdd = ActiveCell.Address
    Range(fstadd, lstadd).Select

        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15652757
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete

Sheets("CR").UsedRange.Find(what:="Crossover", lookat:=xlWhole).Select
ActiveCell.Offset(0, 1).Select
up1Add = ActiveCell.Address
    Range(up1Add, upAdd).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
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
Sheets("CR").UsedRange.Find(what:="CR / Sys / ITM", lookat:=xlWhole).Select
Sheets("CR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.EntireColumn.Select
Selection.ColumnWidth = 7
Call CRPivotTable
Dim visPvtItm As String

'Calculate ITM for the current Month
Set pvtTbl = Worksheets("CR").PivotTables("pvtKPIALL")
fixedDate = Sheet1.combYear.value

Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = fixedDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

      
    Dim X As Long
    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 8).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
    Cells(X, 8).NumberFormat = "0.00"
  Next X
             

'Calculate IMQ for the Quarter( from current date to the last 3 months) of the selected month and year

Set pvtTbl = Worksheets("CR").PivotTables("pvtKPIALL")
pvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, fixedDate), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, fixedDate), "yyyy" & "-" & "mm")

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = fixedDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 9).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / 3
    Cells(X, 9).NumberFormat = "0.00"
    
    'Application.WorksheetFunction.RoundUp (Cells(x, 8).Value)
    'Application.RoundUp (Cells(x, 9).Value)
    Next X

'calculate MAT for the last one year from the selected year_month
endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 11).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / 12
    Cells(X, 11).NumberFormat = "0.00"
    Next X

'Calculate ITM for the same month in the previous year
endDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = endDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 12).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
    Cells(X, 12).NumberFormat = "0.00"
    Next X
    
'Calculate IMQ for the quarter in the last (Previous) year

startDate = Format(fixedDate, "yyyy" & "-" & "mm")
prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = Worksheets("CR").PivotTables("pvtKPIALL")
pvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 13).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / 3
    Cells(X, 13).NumberFormat = "0.00"
    Next X

'Calculate MAT for the previous year of the current year

startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate1 = Format(DateAdd("yyyy", -2, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 15).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / 12
    Cells(X, 15).NumberFormat = "0.00"
    Next X

'calculate YTD from current month of the current year to the first motnh of current year i.e. from January to te current year

startDate = Format(fixedDate, "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 6, 2)

endDate = Format(DateAdd("m", -EndDateMonth, fixedDate), "yyyy" & "-" & "mm")

   
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 10).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / EndDateMonth
    Cells(X, 10).NumberFormat = "0.00"
    Next X


''calculate YTD from current month of the Previous year to the first motnh of Previous year i.e. from January to the current year


startDate = Format(fixedDate, "yyyy" & "-" & "mm")
StartDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 6, 2)

endDate = Format(DateAdd("m", -EndDateMonth, StartDate1), "yyyy" & "-" & "mm")

   
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > StartDate1 Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 14).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False) / EndDateMonth
    Cells(X, 14).NumberFormat = "0.00"
    Next X



fixedDate = Sheet1.combYear.value

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
Cells(3, 16).Select
i = 17
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
    Else
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
            rng = Range("AE3:AJ91")
            
            If i <= 28 Then
            For X = 2 To lr
            On Error Resume Next
            Cells(X, i).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
            'Round (Cells(x, i).Value)

            Next X
             
    End If
    i = i + 1
    End If
Next pvtItm
   

Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H3:O91").Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
     Range("E3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
    
     Range("F3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
     Range("g3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
    
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
    
    
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("X3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("AA3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("AB3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    
    
    Sheets("CR").Select
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
    
        
    Range("G4").Select
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
    Workbooks(myWorkBook).Save
   End Sub
Public Function CRPivotTable()
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim pvtItm As PivotItem

Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim wsData As Worksheet
Dim wsPtTable As Worksheet

Dim rngData As String
Dim pvtExcel As String
Dim strtPt As String
Dim SrcData As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer

'Case select for sheet tab
    KPISheetName = Sheet1.comb6NC1.value

    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = Sheet1.comb6NC1.value
    dateValue = Sheet1.combYear.value
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    myPath = ThisWorkbook.Path
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
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
    Range("A1").value = "Period"
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
     

'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
    Sheets("Aggr. SWO Data CV").Activate
    Dim l As Long
    l = Application.WorksheetFunction.Match("BuildingBlock", Range("1:1"), 0)
    Range("A1").Select
    fstCellAdd = ActiveCell.Address
    Range("A1").End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    Selection.AutoFilter
    ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=l, Criteria1:="=Buildingblocks Aggregated"
    Range("A1").Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range("A1").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
    Range(fstFiltCellAdd1).End(xlToRight).Select
    fstFiltCellAdd2 = ActiveCell.Address
   ' lastFiltCellAdd = ActiveCell.Address
    Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
    
   'ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=4, Criteria1:="=Non-Parts Aggregated"
'Remove the values which are less then 10% of the top value in the Total Calls(#) column
    
  
        
'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)

        Set wsPtTable = Worksheets(Sheets.Count)

        'Set wsPtTable = Worksheets(3)
        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate

        Set wsData = Worksheets("Aggr. SWO Data CV")
        Worksheets("Aggr. SWO Data CV").Activate
        sourceSheet = ActiveSheet.name

        ActiveSheet.Cells(1, 1).Select
        fstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select

        lstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
        Sheets(wsptName).Activate
        rngData = fstadd & ":" & lstadd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
                
           ActiveSheet.Range("A1").Select
        ActiveCell.PivotTable.name = "pvtCallRate"
    
        wsPtTable.Activate
        
        Set pt = wsPtTable.PivotTables("pvtCallRate")
        Set pf = pt.PivotFields("Period")
        pf.Orientation = xlPageField
        pf.Position = 1
        Set pf = pt.PivotFields("SubSystem")
        pf.Orientation = xlRowField
        pf.Position = 1
        Set pf = pt.PivotFields("BuildingBlock")
        pf.Orientation = xlRowField
        pf.Position = 2
        Set pf = pt.PivotFields("Part12NC")
        pf.Orientation = xlColumnField
        pf.Position = 1
        
       With ActiveSheet.PivotTables("pvtCallRate").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
       End With
        ActiveSheet.PivotTables("pvtCallRate").AddDataField ActiveSheet.PivotTables( _
        "pvtCallRate").PivotFields("Total Calls (#)"), "Count of Calls", xlSum
        
        ActiveSheet.PivotTables("pvtCallRate").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtCallRate").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
       
        With ActiveSheet.PivotTables("pvtCallRate")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
        End With
    
        ActiveSheet.PivotTables("pvtCallRate").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        ActiveSheet.PivotTables("pvtCallRate").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
        With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 1
        End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtCallRate")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtCallRate")
        .ColumnGrand = True
        .RowGrand = True
    End With
    pvtTbl.RefreshTable
    
    Columns("A:D").EntireColumn.AutoFit
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
    pvtTbl.TableRange2.Copy
    Windows("CTS_KPI_Summary.xlsx").Activate
    Range("AK1").PasteSpecial
    Range("AK1").Select
        ActiveCell.PivotTable.name = "pvtKPIALL"

        'ActiveSheet.PivotTables("pvtCallRate").Location =
        '"'[CTS_KPI_Summary.xlsx]CR'!$AK$3"
        Windows("CTS_KPI_Summary.xlsx").Activate
        Sheets("CR").Activate
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
    
End Function
Public Sub MTTRRateCalculationNew()
Dim fixedDate, myPath, CTSExcel, CTSWorkBook, pvtExcel, myPvtWorkBook As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1, KPISheetName As String

'Case select for sheet tab
    KPISheetName = Sheet1.comb6NC1.value

    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = Sheet1.comb6NC1.value
    dateValue = Sheet1.combYear.value
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    Exit Sub
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    CTSProductName = KPISheetName
   
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
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
    outPutFilePath = ThisWorkbook.Path & "\"
    installFlName = outPutFilePath & "CTS_KPI_Summary.xlsx"
    Application.Workbooks.Open (installFlName), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Sheets("KPI-All").Select
    
    
    'Windows("CTS_KPI_Summary.xlsx").Activate
    'Cells.Select
    'Application.CutCopyMode = False
    'Selection.Delete Shift:=xlUp
    'Range("A1").Select
    'Windows("Veradius_2014-06.xlsx").Activate
    'Selection.Copy
    'Windows("CTS_KPI_Summary.xlsx").Activate
    'ActiveSheet.Paste
    fixedDate = Sheet1.combYear.value
endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters
With ActiveSheet.PivotTables("pvtKPIALL")
        .ColumnGrand = True
        .RowGrand = False
    End With
   ' pvtTbl.RefreshTable
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm
    
'        Workbooks("KPI's_NewVer_1.0_change_R").Activate
        Sheets("MTTR").Select
        Range("A:A").Select
        On Error Resume Next
        Selection.EntireRow.Select
        Selection.EntireRow.Delete
        Application.Columns.Ungroup
        rows("1:1").Select
        'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets("KPI-All").Select
        pvtTbl.TableRange1.Select
        pvtTbl.TableRange1.Copy
        
        Sheets("MTTR").Select
        Range("a1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("1:1").Select
        Selection.EntireRow.Delete
        Sheets("MTTR").UsedRange.Find(what:="ETTR (days)", lookat:=xlWhole).Select
        ActiveCell.Offset(1, 0).Select
        fstclmn = ActiveCell.Address
        ActiveCell.End(xlToRight).Select
       lstclmnAdd = ActiveCell.Address
        Range(fstclmn, lstclmnAdd).Select
        Selection.EntireColumn.Select
        Selection.EntireColumn.Delete
        Cells(2, 1).Select
        Selection.EntireRow.Select
        Sheets("MTTR").UsedRange.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
        deleteClmnsAdd = ActiveCell.Address
        Sheets("MTTR").UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
        ActiveCell.Offset(1, 1).Select
        deleteLstClmnsAdd = ActiveCell.Address
        Range(deleteClmnsAdd, deleteLstClmnsAdd).Select
        Selection.EntireColumn.Select
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
    
     myPath = ThisWorkbook.Path
    CTSExcel = myPath & "\" & Dir(myPath & "\" & "CTS_Guidelines.xlsx")  'input file path
    Application.Workbooks.Open (CTSExcel), False
    
    CTSWorkBook = ActiveWorkbook.name
    Workbooks(CTSWorkBook).Activate
    Sheets("Shhet2").Activate
    Sheets("Sheet2").UsedRange.Find(what:="MTTR / Sys / Yr", lookat:=xlWhole).Select

    Selection.EntireColumn.Select
    Selection.Copy
    
    Windows(myWorkBook).Activate
    Range("C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(CTSWorkBook).Close
    Windows(myWorkBook).Activate
    Sheets("MTTR").Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C3").Select
    fstadd = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    
    lstadd = ActiveCell.Address
    
    Range(fstadd).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Selection.AutoFill Destination:=Range(fstadd, lstadd)
    Range(fstadd, lstadd).Select
    Calculate
    Cells(2, 3).value = "SS&BB"
    
    
    
    Sheets("MTTR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
ActiveCell.Offset(1, 1).Select
    CRSysYrFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -4).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    CRSysYrLstAdd = ActiveCell.Address
    Range(CRSysYrFstAdd).Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]+RC[-2])/2047"
    Range(CRSysYrFstAdd).Select
    Selection.AutoFill Destination:=Range(CRSysYrFstAdd, CRSysYrLstAdd)
    Range(CRSysYrFstAdd, CRSysYrLstAdd).Select
    Range(CRSysYrFstAdd, CRSysYrLstAdd).NumberFormat = "0.00"

    Calculate
    Windows(myWorkBook).Activate
    Range("A2").Select
    fstadd1 = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    lstadd2 = ActiveCell.Address
    Range(fstadd1, lstadd2).Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    
    
    'Columns("AE:AE").Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("AE3").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    'Range("AE3").Select
    'Selection.AutoFill Destination:=Range("AE3:AE91")
    'Range("AE3:AE91").Select
    'Calculate
    
        
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    Sheets("MTTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    nonPartsFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    nonPartslstAdd = ActiveCell.Address
    
    Range(nonPartsFstAdd, nonPartslstAdd).Select
    Selection.NumberFormat = "0.00"
    'Sheets("MTTR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
    'Cells(3, 7).Select
    'ActiveCell.End(xlDown).Select
    'lstRowAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    'Range(lstRowAdd).Select
    'Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    'ActiveCell.EntireRow.Delete
    'Sheets("MTTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    'ActiveCell.Offset(2, 0).Select
    'pkAdd = ActiveCell.Address
    'fstCellAdd = ActiveCell.Address(ReferenceStyle:=xlA1)
    'mioflstcell = Left(fstCellAdd, 3)
    'midoflstadd = Mid(lstRowAdd, 4)
    'Add = mioflstcell & midoflstadd
   ' ActiveSheet.Range(fstCellAdd, Add).Select
    Range(nonPartsFstAdd, nonPartslstAdd).Select
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
    
    Sheets("MTTR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
   ActiveCell.Offset(2, 0).Select
    nonPartsFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    nonPartslstAdd = ActiveCell.Address
    
    Range(nonPartsFstAdd, nonPartslstAdd).Select
    Selection.NumberFormat = "0.00"
    
    
        
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
        .NumberFormat = "0"

    End With
    
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    Sheets("MTTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select

    ActiveCell.value = "MAT # of Calls profiles"
    ActiveCell.Offset(1, 0).Select

    ActiveCell.value = "Non-Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Parts"
    ActiveCell.Offset(0, 1).Select

    ActiveCell.value = "CR / Sys / ITM"
    ActiveCell.Offset(1, 0).Select

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
    ActiveCell.value = "Current Year Avg. MTTR / Sys"
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
    
    
    fixedDate1 = Sheet1.combYear.value
    frmtData = Format(DateAdd("m", 1, fixedDate1), "mmm" & "-" & "yyyy")
'currentdate = Format(Now(), "yyyymm")

   endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
   fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
'    j = 12
    Do Until frmEndDate = frmtData
    ActiveCell.value = frmEndDate
    ActiveCell.Offset(0, 1).Select
    frmEndDate = Format(DateAdd("m", 1, frmEndDate), "mmm" & "-" & "yyyy")
    Loop

    Range("A1").Select
    fstadd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstadd = ActiveCell.Address
    ActiveCell.Offset(-1, 0).Select
    upAdd = ActiveCell.Address
    Range(fstadd, lstadd).Select

        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15652757
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Sheets("MTTR").UsedRange.Find(what:="Crossover", lookat:=xlWhole).Select
ActiveCell.Offset(0, 1).Select
up1Add = ActiveCell.Address
    Range(up1Add, upAdd).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
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
Sheets("MTTR").UsedRange.Find(what:="CR / Sys / ITM", lookat:=xlWhole).Select
Sheets("MTTR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.EntireColumn.Select
Selection.ColumnWidth = 7
Call MTTRPivotTableNew
Dim visPvtItm As String
Set pvtTbl = Worksheets("MTTR").PivotTables("pvtMTTR")
fixedDate = Sheet1.combYear.value
'currentdate = Format(Now(), "yyyymm")
endDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = ActiveSheet.PivotTables("pvtMTTR")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = fixedDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

      
    Dim X As Long
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 8).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
    Cells(X, 8).NumberFormat = "0.00"
  Next X
             


endDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = Worksheets("CR").PivotTables("PivotTable1")
pvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, fixedDate), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, fixedDate), "yyyy" & "-" & "mm")

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = fixedDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 9).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / 3
    Cells(X, 9).NumberFormat = "0.00"
    
    'Application.WorksheetFunction.RoundUp (Cells(x, 8).Value)
    'Application.RoundUp (Cells(x, 9).Value)
    Next X

'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 11).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / 12
    Cells(X, 11).NumberFormat = "0.00"
    Next X

'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm.value = endDate Then
    pf.CurrentPage = pvtItm.Caption
    End If
Next

    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 12).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
    Cells(X, 12).NumberFormat = "0.00"
    Next X
    
'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

Set pvtTbl = Worksheets("MTTR").PivotTables("PivotTable1")
pvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 13).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / 3
    Cells(X, 13).NumberFormat = "0.00"
    Next X

'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
endDate1 = Format(DateAdd("yyyy", -2, fixedDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")
pf.ClearAllFilters

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 15).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / 12
    Cells(X, 15).NumberFormat = "0.00"
    Next X

'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 6, 2)

endDate = Format(DateAdd("m", -EndDateMonth, fixedDate), "yyyy" & "-" & "mm")

   
Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / EndDateMonth
    Cells(X, 10).NumberFormat = "0.00"
    Next X


'fixedDate = 201406

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
StartDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
EndDateMonth = Mid(fixedDate, 6, 2)

endDate = Format(DateAdd("m", -EndDateMonth, StartDate1), "yyyy" & "-" & "mm")

   
Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
'2013-01
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm <= endDate Or pvtItm > StartDate1 Then
            pvtItm.Visible = False
    Else
            pvtItm.Visible = True
    End If
Next pvtItm

'   ActiveSheet.Range("N3").Select
    lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
    rng = Range("AE3:AJ91")
            
    For X = 3 To lr
    On Error Resume Next
    Cells(X, 14).value = (Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)) / EndDateMonth
    Cells(X, 14).NumberFormat = "0.00"
    Next X



fixedDate = Sheet1.combYear.value

currentdate = Format(Now(), "yyyymm")
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
Set pvtTbl = ActiveSheet.PivotTables("PivotTable1")
Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        pf.CurrentPage = "(All)"
        
Cells(3, 16).Select
i = 17
For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
    If pvtItm < endDate Or pvtItm > startDate Then
    Else
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
            rng = Range("AE3:AJ91")
            
            If i <= 28 Then
            For X = 2 To lr
            On Error Resume Next
            Cells(X, i).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
            'Round (Cells(x, i).Value)

            Next X
             
    End If
    i = i + 1
    End If
Next pvtItm
   

Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H3:O91").Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
     Range("E3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
    
     Range("F3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
     Range("g3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[40]C)"
    
    
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
    
    
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("X3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("AA3").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[88]C)"
    Range("AB3").Select
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
    
    
    
    Range("G4").Select
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
    Workbooks(myWorkBook).Save
   End Sub
 Public Function MTTRPivotTableNew()
 Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim pvtItm As PivotItem

Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim wsData As Worksheet
Dim wsPtTable As Worksheet

Dim rngData As String
Dim pvtExcel As String
Dim strtPt As String
Dim SrcData As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer

'Case select for sheet tab
    KPISheetName = Sheet1.comb6NC1.value

    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = Sheet1.comb6NC1.value
    dateValue = Sheet1.combYear.value
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    myPath = ThisWorkbook.Path
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
    ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
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
    Range("A1").value = "Period"
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
     

'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
    Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    fstCellAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    If ActiveSheet.AutoFilterMode = True Then
 ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter
 End If
    Dim l As Long
    l = Application.WorksheetFunction.Match("BuildingBlock", Range("1:1"), 0)
    ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=l, Criteria1:="=Buildingblocks Aggregated"
    Range("A1").Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range("A1").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
    Range(fstFiltCellAdd1).End(xlToRight).Select
    fstFiltCellAdd2 = ActiveCell.Address
   ' lastFiltCellAdd = ActiveCell.Address
    Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
    
   'ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=4, Criteria1:="=Non-Parts Aggregated"
'Remove the values which are less then 10% of the top value in the Total Calls(#) column
    
  
        
'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)

        Set wsPtTable = Worksheets(Sheets.Count)

        'Set wsPtTable = Worksheets(3)
        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate

        Set wsData = Worksheets("Aggr. SWO Data CV")
        Worksheets("Aggr. SWO Data CV").Activate
        sourceSheet = ActiveSheet.name

        ActiveSheet.Cells(1, 1).Select
        fstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select

        lstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
        Sheets(wsptName).Activate
        rngData = fstadd & ":" & lstadd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
             
        Range("A1").Select
        ActiveCell.PivotTable.name = "pvtMTTR"
             
        wsPtTable.Activate
        
        Set pt = wsPtTable.PivotTables("pvtMTTR")
        Set pf = pt.PivotFields("Period")
        pf.Orientation = xlPageField
        pf.Position = 1
        Set pf = pt.PivotFields("SubSystem")
        pf.Orientation = xlRowField
        pf.Position = 1
        Set pf = pt.PivotFields("BuildingBlock")
        pf.Orientation = xlRowField
        pf.Position = 2
        Set pf = pt.PivotFields("Part12NC")
        pf.Orientation = xlColumnField
        pf.Position = 1
        
       With ActiveSheet.PivotTables("pvtMTTR").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
       End With
        ActiveSheet.PivotTables("pvtMTTR").AddDataField ActiveSheet.PivotTables( _
        "pvtMTTR").PivotFields("Avg. MTTR/Call (hrs)"), "#MTTR/Call (hrs)", xlSum
        
        ActiveSheet.PivotTables("pvtMTTR").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtMTTR").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
       
        With ActiveSheet.PivotTables("pvtMTTR")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
        End With
    
        ActiveSheet.PivotTables("pvtMTTR").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        ActiveSheet.PivotTables("pvtMTTR").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
        With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 1
        End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtMTTR")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtMTTR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    pvtTbl.RefreshTable
    
    Columns("A:E").EntireColumn.AutoFit
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
        ActiveSheet.PivotTables("pvtMTTR").Location = _
        "'[CTS_KPI_Summary.xlsx]MTTR'!$AK$3"
        Windows("CTS_KPI_Summary.xlsx").Activate
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
    
End Function


'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'""""RRR Data-Merge All CCC Files and Create a RRR Report using FDV Raw Data

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


Public Sub ListSubfoldersFile()

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
    myPath = ThisWorkbook.Path & "\Input Source\"
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
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "RRR_Report"

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
    srcfile = ThisWorkbook.Path & "\" & "CTS-Cost to Serve_RRR_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    Application.Workbooks.Open(srcfile).Activate
    Worksheets("MergedCCC1").Activate
    wbName1 = ActiveWorkbook.name
    Columns("L:L").Select
    Selection.Replace what:="100", Replacement:="1", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Columns("M:O").Select
        Selection.EntireColumn.Delete
        ActiveSheet.Cells(1, 1).Select
    'createPivotTableRRRData
    'pivotChart
    Workbooks(wbName1).Activate
    ActiveWorkbook.Save
    'MsgBox "RRR Data is Generated succesfully", vbOKOnly
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


    Dim myFile As String

        Dim inputItem As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myPath = ThisWorkbook.Path & "\Input Source\"
    myFile = "BCTool_SWO_RawData_SingleVersionOfTheTruth"
    inputItem = myPath & "\" & Dir(myPath & "\" & "BCTool_SWO_RawData_SingleVersionOfTheTruth" & "*.xls*") 'input file path
    Application.Workbooks.Open (inputItem), False
    myWorkBook = ActiveWorkbook.name
    
    Workbooks(myWorkBook).Activate
    With ThisWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "CombinedFDV"
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
Cells(1, 1).Select
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
Columns("M:M").Select
Selection.EntireColumn.Delete
Columns("L:L").Select
Selection.Replace what:="100", Replacement:="1", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


ActiveWorkbook.Save

'Create Pivot Table and Pivot Chart
'

'
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim rngData As String
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


Dim lstadd As String

'pt.ManualUpdate = False
    myPath = ThisWorkbook.Path & "\OLD_Programs"
    pvtExcel = myPath & "\" & Dir(myPath & "\" & "CTS-Cost to Serve_RRR_" & "*.xls*")  'input file path
    Application.Workbooks.Open (pvtExcel), False
    
   ' Workbooks(wbName1).Activate

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
    Sheets.Add after:=Worksheets(Worksheets.Count)

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
        
    Sheets(wsptName).Select
    Cells(2, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(wsptName & "!$A$2:$C$1000")
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[Entitlement]")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[RR]")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[BuildingBlock]")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[Entitlement]")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[RR]")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[BuildingBlock]")
        .Orientation = xlPageField
        .Position = 3
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields("[Range].[SubSystem]")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.CubeFields( _
        "[Range].[MaterialDescription]")
        .Orientation = xlRowField
        .Position = 2

    End With
    ActiveSheet.PivotTables("PivotTable1").CubeFields.GetMeasure "[Range].[SWO]", _
        xlCount, "Count of SWO"
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").CubeFields("[Measures].[Count of SWO]"), "Count of SWO"
    ActiveChart.ChartArea.Select
    
   ' ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
    '    "PivotTable1").CubeFields("[Measures].[Count of SWO]"), "Count of SWO"
    'Range("A2").Select
    
    
   ' ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[SubSystem].[SubSystem]").AutoSort xlDescending, _
        "[Measures].[Count of SWO]", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1
      
        
    ActiveSheet.PivotTables("PivotTable1").CubeFields(12).EnableMultiplePageItems _
        = True
    
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.Shapes("Chart 1").IncrementLeft -210.75
    ActiveSheet.Shapes("Chart 1").IncrementTop -12.75
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("[Range].[MaterialDescription].&", "[Range].[MaterialDescription].&[]")
    
    ActiveSheet.PivotTables("PivotTable1").PivotSelect _
        "'[Range].[SubSystem].[SubSystem]'[All]", xlLabelOnly + xlFirstRow, True
    ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Range].[SubSystem].[SubSystem]").DrilledDown = False
    ActiveSheet.PivotTables("PivotTable1").PivotSelect _
        "'[Range].[SubSystem].[SubSystem]'[All]", xlLabelOnly + xlFirstRow, True
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("")
        
      Sheets(wsptName).name = "RRRPvtTablePvtChart"
    
End Sub
Public Sub PrtChartCalculation()
Dim StrFile, myPath, myFile As String
Dim objFSO, destRow As Long
Dim wb As Workbook
Dim inputItem As String
Dim rng As Range
Dim chtObject As ChartObject
Dim rownum As Integer
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myPath = ThisWorkbook.Path & "\Output_Files\"
    inputItem = myPath & "\" & Dir(myPath & "\" & "CTS-Cost to Serve_RRR_*" & "*.xls*") 'input file path
    Application.Workbooks.Open (inputItem), False
    myWorkBook = ActiveWorkbook.name
    
    Workbooks(myWorkBook).Activate

    Sheets("RRRPvtTablePvtChart").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("")
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("")
    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A6:B28").Select
    Selection.Copy
    With ThisWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "ParetoChart"
    
    End With
    Sheets("ParetoChart").Select
    Range("A2").Select
    ActiveSheet.Paste
    Cells(1, 1).value = "SS"
    Cells(1, 2).value = "Total"
    Sheets("RRRPvtTablePvtChart").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[0]")
    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A6:B28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ParetoChart").Select
    Range("C2").Select
    ActiveSheet.Paste
    Sheets("RRRPvtTablePvtChart").Select
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[1]")
    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A6:B22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ParetoChart").Activate
    Range("E2").Select
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("C2").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C1:R24C1,R2C5:R24C6,2,0)"
    'Range("C2").Select
    'Selection.AutoFill Destination:=Range("C2:C24")
    'Range("C2:C24").Select
    'Calculate
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range("B2") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range("A2:B24")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range("F2") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range("E2:F24")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C1:R24C1,R2C5:R24C6,2,0)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C24")
    Range("C2:C24").Select
    Calculate
    Selection.Copy
    'Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:="#N/A", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Selection.PasteSpecial , xlPasteValues
    'Range("C2").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2]:R[22]C[-2],RC[2]:R[22]C[3],2,0)"
    'Range("C2").Select
    'Selection.AutoFill Destination:=Range("C2:C24")
    'Range("C2:C24").Select
    'Calculate
'    ActiveWindow.ScrollRow = 1
    Range("C1").Select
    ActiveCell.value = "Non RR"
    Range("D1").Select
    ActiveCell.value = "RR"
    
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D24")
    Range("D2:D24").Select
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:="#N/A", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("E2:F24").Select
    ActiveWindow.ScrollRow = 1
    Range("C2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("E:F").EntireColumn.Hidden = True
    Application.CutCopyMode = False

    Application.ScreenUpdating = False
    Columns("C:D").Select
    Selection.Replace what:="#N/A", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(2, 3).Select
    fstCellAdd = ActiveCell.Address
    Cells(2, 3).Select
    ActiveCell.End(xlDown).Select
    
    lstCellAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    myrange = ActiveSheet.Range(fstCellAdd, lstCellAdd)
'    Columns("C:D").Select
    
    'ActiveCell.value = WorksheetFunction.Sum(myRange)
    'totalSum = ActiveCell.value
    
    Columns("D").Copy
    Cells(1, 7).PasteSpecial xlPasteValues
    Cells(1, 7).value = 0
    Cells(2, 7).value = Cells(2, 3).value
    Cells(2, 3).Select
    Set rng = Range("G3:G" & Range("C2").End(xlDown).Row)

    rng.FormulaR1C1 = "=R[-1]C+RC[-4]"
    rng.value = rng.FormulaR1C1
    
    Cells(1, 7).value = "CUM NRR"
    Cells(1, 8).value = "% NON RR"

    Cells(2, 7).Select
    ActiveCell.Offset(0, 1).Select
    percentFstCellAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    percentlstCellAdd = ActiveCell.Address
    ActiveWorkbook.Names.Add name:="NRRTtl", RefersToR1C1:= _
        "=SUM(ParetoChart!R2C3:R24C3)"
    ActiveWorkbook.Names("NRRTtl").Comment = ""
 
    Cells(1, 8).Select
    Do Until ActiveCell.Address = percentlstCellAdd
        ActiveCell.Offset(1, 0).FormulaR1C1 = "=RC[-1]/NRRTtl*100"
        ActiveCell.Offset(1, 0).Select
    Loop

    Application.ScreenUpdating = True
    Range("A2").Select
    

    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart(250, xlColumnStacked).Select
    
    ActiveChart.SetSourceData Source:=Range("ParetoChart!$A$2:$A$24")
    ActiveChart.ChartTitle.Select
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).name = "=""Total"""
    ActiveChart.FullSeriesCollection(1).Values = "=ParetoChart!$B$1:$B$24"
    ActiveChart.FullSeriesCollection(1).XValues = "=ParetoChart!$A$2:$A$24"
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).name = "=""RR"""
    ActiveChart.FullSeriesCollection(2).Values = "=ParetoChart!$I$2:$I$24"
    ActiveChart.FullSeriesCollection(2).XValues = "=ParetoChart!$A$2:$A$24"
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).name = "=""% NON RR"""
    ActiveChart.FullSeriesCollection(3).Values = "=ParetoChart!$H$1:$H$24"
    ActiveChart.FullSeriesCollection(3).XValues = "=ParetoChart!$A$2:$A$24"
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnStacked
    ActiveChart.FullSeriesCollection(3).ChartType = xlLineMarkers
    ActiveChart.FullSeriesCollection(3).AxisGroup = 2
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 328
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "RRR 2015"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "RRR 2015"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 8).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Shadow.Type = msoShadow22
        .Shadow.Visible = msoTrue
        .Shadow.Style = msoShadowStyleOuterShadow
        .Shadow.Blur = 4
        .Shadow.OffsetX = 1.8369701987E-16
        .Shadow.OffsetY = 3
        .Shadow.RotateWithShape = msoFalse
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.599999994
        .Shadow.Size = 100
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 16
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(4, 5).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Shadow.Type = msoShadow22
        .Shadow.Visible = msoTrue
        .Shadow.Style = msoShadowStyleOuterShadow
        .Shadow.Blur = 4
        .Shadow.OffsetX = 1.8369701987E-16
        .Shadow.OffsetY = 3
        .Shadow.RotateWithShape = msoFalse
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.599999994
        .Shadow.Size = 100
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 16
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(146, 208, 80)
        .Transparency = 0
    End With
    
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Solid
    End With
    
    Dim Ch As ChartObject
    Set Ch = Worksheets("ParetoChart").ChartObjects(1)
    With Ch
        .Top = Range("A30").Top
        .Width = Range("A30:K30").Width
        .Height = Range("A30:K46").Height
    End With
       ActiveSheet.ChartObjects(1).Left = ActiveSheet.Columns(1).Left
   ActiveSheet.ChartObjects(1).Top = ActiveSheet.rows(30).Top

    Sheets("RRRPvtTablePvtChart").Select
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("")
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("")
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[0]")
    
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:B32").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("ParetoChart").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("J2:K28").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range( _
        "K2:K28"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range("J2:K28")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("RRRPvtTablePvtChart").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("[Range].[MaterialDescription].&", "[Range].[MaterialDescription].&[]")
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:B24").Select
    Selection.Copy
    Sheets("ParetoChart").Select
    Range("L1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range( _
        "M1:M20"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range("L1:M20")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(C[-2],C[-4],1,0),RC[-4])"
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N28")
    Range("N2:N28").Select
    Calculate
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-1],C[-5]:C[-4],2,0)"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O28")
    Range("O2:O28").Select
    Calculate
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-3]"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P28")
    Range("P2:P28").Select
   
    Range("Q2:R2").Select
    Selection.ClearContents
    Columns("O:O").Select
'    Sheets("Sheet5").Select
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "SS"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Count of SWO"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "SS"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "NON Parts-NON RR"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "SS"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Count of SWO"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Parts-NON RR"

    Cells(1, 2).Select
    ActiveCell.Offset(1, 0).Select
    percentFstCellAdd = ActiveCell.Address
    'ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 7).Select
    percentlstCellAdd = ActiveCell.Address
    ActiveWorkbook.Names.Add name:="NRRTtl", RefersToR1C1:= _
        "=SUM(ParetoChart!R2C2:R24C2)"
    ActiveWorkbook.Names("NRRTtl").Comment = ""

    Cells(1, 9).Select
    Do Until ActiveCell.Address = percentlstCellAdd
        ActiveCell.Offset(1, 0).FormulaR1C1 = "=RC[-5]/NRRTtl*100"
        ActiveCell.Offset(1, 0).Select
    Loop
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "% RR"
 
    Sheets("ParetoChart").Select
    Range("J2").Select
    Selection.End(xlToRight).Select
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("J2:K24").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("ParetoChart!$J$2:$K$28")
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).name = "=""NON-RR Total"""
    ActiveChart.FullSeriesCollection(1).Values = "=ParetoChart!$K$1:$K$28"
    ActiveChart.FullSeriesCollection(1).XValues = "=ParetoChart!$J$2:$J$28"
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).name = "=""NON Parts-NON RR"""
    ActiveChart.FullSeriesCollection(2).Values = "=ParetoChart!$M$1:$M$20"
    ActiveChart.FullSeriesCollection(2).XValues = "=ParetoChart!$J$2:$J$28"
    
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).name = "=""Part Replaced NON RR"""
    ActiveChart.FullSeriesCollection(3).Values = "=ParetoChart!$P$1:$P$28"
    ActiveChart.FullSeriesCollection(3).XValues = "=ParetoChart!$J$2:$J$28"
         
    ActiveChart.ApplyLayout (4)
    ActiveChart.ChartArea.Select
    With Selection
    .Width = 500
    .Height = 300
    
    End With
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Delete
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Delete
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Parts Vs Non Parts 2015"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Parts Vs Non Parts 2015"
    
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, O, O)
        .Transparency = 0
        .Solid
    End With
   
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.ChartArea.Select
    Application.CommandBars("Format Object").Visible = False
'End With

    Set Ch = Worksheets("ParetoChart").ChartObjects(2)
    With Ch
        .Top = Range("M30").Top
        .Width = Range("M30:V30").Width
        .Height = Range("M30:V46").Height
    End With

   ActiveSheet.ChartObjects(2).Left = ActiveSheet.Columns(13).Left
   ActiveSheet.ChartObjects(2).Top = ActiveSheet.rows(30).Top
 
 End Sub

Public Sub createPivotTableAggregatedKPIMaster()

Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim ptcache As PivotCache
Dim ptname As String
Dim pvtItm As PivotItem

Dim ws As Worksheet
Dim sht As Worksheet
Dim sht1 As Worksheet
Dim wsData As Worksheet
Dim wsPtTable As Worksheet

Dim rngData As String
Dim pvtExcel As String
Dim strtPt As String
Dim SrcData As String
Dim fstadd1 As String
Dim sourceSheet As String
Dim myPath As String
Dim fstadd As String
Dim lstadd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer

'Case select for sheet tab
    KPISheetName = "Veradius"
    Select Case KPISheetName

        Case "IXR-MOS Pulsera-Y"
        KPISheetName = "Pulsera"
        selectSheet = 1

        Case "IXR-MOS BV Vectra-N"
        KPISheetName = "BV Vectra"
        selectSheet = 1

        Case "IXR-MOS Endura-Y"
        KPISheetName = "Endura"
        selectSheet = 1

        Case "IXR-MOS Veradius-Y"
        KPISheetName = "Veradius"
        selectSheet = 1

        Case "IXR-CV Allura FC-Y"
        KPISheetName = "Allura FC"
        selectSheet = 1

        Case "IXR-MOS Libra-N"
        KPISheetName = "Libra"
        selectSheet = 1

        Case "DXR-PrimaryDiagnost Digital-N"
        KPISheetName = "PrimaryDiagnost Digital"
        selectSheet = 1

        Case "DXR-MicroDose Mammography-Y"
        KPISheetName = "MicroDose Mammography"
        selectSheet = 1

        Case "DXR-MobileDiagnost Opta-N"
        KPISheetName = "MobileDiagnost Opta"
        selectSheet = 1

    End Select

    CTSProductName = "Veradius"
    dateValue = "2014-06"
    prdNameFile = KPISheetName & "_" & dateValue

'check if file is present
    filePresent = ""
    filePresent = Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")
    If filePresent = "" Then
        MsgBox "The file " & prdNameFile & " is not available", vbOKOnly
    Exit Sub
    End If

'pt.ManualUpdate = False
'Open Aggregated Data File
    myPath = ThisWorkbook.Path
    pvtExcel = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & prdNameFile & "*.xls*")   'input file path
    Application.Workbooks.Open (pvtExcel), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
     
DataBrekUpFrPivot
'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
    Sheets("Aggr. SWO Data CV").Activate
    Cells(1, 1).Select
    Selection.UnMerge
    Cells(2, 1).Select
    Dim l As Long
    l = Application.WorksheetFunction.Match("BuildingBlock", Range("2:2"), 0)
    Range("A2").Select
    fstCellAdd = ActiveCell.Address
    Range("A2").End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=l, Criteria1:="=Buildingblocks Aggregated"
    Range("A2").Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range("A2").Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address
    Range(fstFiltCellAdd1).End(xlToRight).Select
    fstFiltCellAdd2 = ActiveCell.Address
   ' lastFiltCellAdd = ActiveCell.Address
    Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
    ActiveSheet.Range("H1").Select
    Selection.UnMerge
    ActiveSheet.Range("A1").Select
    Selection.UnMerge
    ActiveSheet.UsedRange.Find(what:="Actual Parts (#)", lookat:=xlWhole).Select
    fstadd = ActiveCell.Offset(1, 0).Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstadd = ActiveCell.Address
    Range(fstadd).Select
  Selection.EntireColumn.Insert
    
    lastRow = ActiveSheet.Range(ActiveCell.Offset(0, -1) & ActiveSheet.rows.Count).End(xlUp).Row
        Range(fstadd).Offset(-1, 0).value = "RRR %"
        Range(fstadd).FormulaR1C1 = _
        "=(RC[-1]/RC[-6]*100)"
        Range(fstadd, lstadd).NumberFormat = "0.00"
      ' Range("N3", "N" & Cells(rows.Count, 1).End(xlUp).Row).FillDown
        Range(fstadd).AutoFill Destination:=ActiveSheet.Range(fstadd, lstadd)
        Range(fstadd, lstadd).Select
        Calculate
        Selection.Copy
        Range(fstadd, lstadd).PasteSpecial xlPasteValues
    'Apply Icon Set Conditional formatting on RRR Column Values
    Range(fstadd).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Triangles)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValuePercent
        .value = 33
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValuePercent
        .value = 67
        .Operator = 7
    End With
    
   'ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter Field:=4, Criteria1:="=Non-Parts Aggregated"
'Remove the values which are less then 10% of the top value in the Total Calls(#) column
ActiveSheet.UsedRange.Find(what:="Total Calls (#)", lookat:=xlWhole).Select
fstadd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range(fstadd).Offset(1, 0).End(xlDown).Select
    fstFiltCellAdd1 = ActiveCell.Address

    ActiveSheet.Range(fstFiltCellAdd, fstFiltCellAdd1).Select
    Max = Application.WorksheetFunction.Max(ActiveSheet.Range(fstFiltCellAdd, fstFiltCellAdd1))
    tenPercentofMax = Max / 10
    
    Set cell = Range(fstFiltCellAdd)
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
    
'Add one column for "Total Cost of Parts & Non-Parts"
ActiveSheet.UsedRange.Find(what:="MPCF", lookat:=xlWhole).Select
        Selection.EntireColumn.Insert
  fstadd = ActiveCell.Offset(1, 0).Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstadd = ActiveCell.Address
    Range(fstadd).Select
    
    lastRow = ActiveSheet.Range(ActiveCell.Offset(0, -1) & ActiveSheet.rows.Count).End(xlUp).Row
        Range(fstadd).Offset(-1, 0).value = "Total Cost of Parts & Non-Parts"
        'Workbooks(myPvtWorkBook).Sheets("Aggr. SWO Data CV").Activate
       ActiveSheet.autofilters

     '   lastRow = ActiveSheet.Range("P" & ActiveSheet.rows.Count).End(xlUp).Row
        Range(fstadd).FormulaR1C1 = _
        "=IF(OR(RC[-12]=""Non-Parts Aggregated"",RC[-12]=""Parts Aggregated""),(RC[-9]*RC[-8]*100)+RC[-6]*200,0)"
      ' Range("N3", "N" & Cells(rows.Count, 1).End(xlUp).Row).FillDown
        Range(fstadd).AutoFill Destination:=Range(fstadd, lstadd)
        Range(fstadd, lstadd).Select
        Calculate
        Selection.Copy
        Range(fstadd, lstadd).PasteSpecial xlPasteValues
       Cells(2, 1).Select
        ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate
'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)

        Set wsPtTable = Worksheets(Sheets.Count)

        'Set wsPtTable = Worksheets(3)
        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets("Aggr. SWO Data CV").Activate

        Set wsData = Worksheets("Aggr. SWO Data CV")
        Worksheets("Aggr. SWO Data CV").Activate
        sourceSheet = ActiveSheet.name

        ActiveSheet.Cells(2, 1).Select
        
        Selection.EntireColumn.Select
        
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ActiveCell.Offset(1, 0).Select
        ActiveCell.value = "Period"
        ActiveCell.Offset(0, 1).value = "Period1"
fstadd = ActiveCell.Offset(1, 0).Address
ActiveCell.Offset(0, 1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
lstadd = ActiveCell.Address
Cells(3, 1).Select

    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Selection.AutoFill Destination:=Range(fstadd, lstadd)
    Range(fstadd, lstadd).Select
    Calculate
    Range(fstadd, lstadd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(2, 1).Select
    
        fstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select

        lstadd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
        Sheets(wsptName).Activate
        rngData = fstadd & ":" & lstadd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
               
        ActiveSheet.Range("A1").Select
        ActiveCell.PivotTable.name = "pvtKPIMASTER"
        wsPtTable.Activate
               
        Set pt = wsPtTable.PivotTables("pvtKPIMASTER")
        Set pf = pt.PivotFields("Period")
        pf.Orientation = xlPageField
        pf.Position = 1
        
        Set pf = pt.PivotFields("SubSystem")
        pf.Orientation = xlRowField
        pf.Position = 1
        Set pf = pt.PivotFields("BuildingBlock")
        pf.Orientation = xlRowField
        pf.Position = 2
        
        
        Set pf = pt.PivotFields("Part12NC-Sub Parts")
        pf.Orientation = xlRowField
        pf.Position = 3
        
        Set pf = pt.PivotFields("PartDescription")
        pf.Orientation = xlRowField
        pf.Position = 4
        
        Set pf = pt.PivotFields("Part12NC")
        pf.Orientation = xlColumnField
        pf.Position = 1
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Total Calls (#)"), "# of Calls", xlSum
        
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Avg. MTTR/Call (hrs)"), "MTTR/Call (hrs)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Avg. ETTR (days)"), "ETTR (days)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Avg. Visits/call (#)"), "Visits/call (#)", xlSum
        
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
        
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Total Costs/part (EUR)"), "Costs/part (EUR)", xlSum
    
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Total Cost of Parts & Non-Parts"), _
        "#Total Cost of Parts & Non-Parts", xlSum
    
        With ActiveSheet.PivotTables("pvtKPIMASTER")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
        End With
    
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("SubSystem").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "", xlDataAndLabel, True
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("PartDescription"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
        With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 2
        End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtKPIMASTER")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtKPIMASTER")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIMASTER")
    Set pf = pvtTbl.PivotFields("Part12NC")

        pf.ClearAllFilters
        pf.EnableMultiplePageItems = True
    
    pf.PivotItems("Parts/Non-Parts Breakups").Visible = False
    ActiveSheet.PivotTables("pvtKPIMASTER").HasAutoFormat = False
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "", xlDataAndLabel, True
    Selection.ColumnWidth = 8
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "'Part12NC-Sub Parts'['-']" _
        , xlDataAndLabel, True
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts"). _
        ShowDetail = False
    Range("B4").Select
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail _
        = False
    
    With ActiveSheet.PivotTables("pvtKPIMASTER")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    pvtTbl.RefreshTable
' Add ConditionalFormatting of Data Bars on total calls of Parts and Non parts
    Columns("E:E").Select
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
    Columns("F:F").Select
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
'Add conditional formatting on MTTR and ETTR Calls
    Columns("G:G").Select
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
    Range("G23").Select
    
    Columns("H:H").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("H19").Select
    
    Columns("I:I").Select
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
    
    Columns("K:L").Select
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
    
    Columns("K:K").Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("E:P").Select
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

    Columns("A:D").Select
    With Selection
        .ColumnWidth = 15
    End With
    Cells(1, 1).Select

    
    Worksheets(wsptName).PivotTables("PivotTable1").PreserveFormatting = False
    Sheets(wsptName).name = "PivotTableAggData"
  '  pt.ManualUpdate = True
ActiveWindow.Zoom = 85


'Add RRR% and CallRate Columns
Sheets("Aggr. SWO Data CV").Select

fixedDate = Sheet1.combYear.value
startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
startDate = Format(startDate, "yyyy" & "-" & "mm")
endDate = Format(fixedDate, "yyyy" & "-" & "mm")

Cells(2, 1).Select
Selection.EntireRow.Select
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.Range("$A$2:$X$391").AutoFilter Field:=1, Criteria1:=endDate

    Sheets("PivotTableAggData").Select

    
    Range("Q6").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-15],'Aggr. SWO Data CV'!R[12]C[-12]:R[2372]C[-2],11,0)"
    Range("Q6").Select
    Selection.AutoFill Destination:=Range("Q6:Q92")
    Range("Q6:Q92").Select
    
    Columns("Q:Q").Select
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
    Selection.NumberFormat = "0"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = "RRR%"
    Columns("Q:Q").Select
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
    Columns("Q:Q").EntireColumn.AutoFit
    Range("Q5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Triangles)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValuePercent
        .value = 33
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValuePercent
        .value = 67
        .Operator = 7
    End With
    Columns("Q:Q").ColumnWidth = 10.14
    
    Range("Q6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("R3").value = "CallRate"
        Range("R3:S3").Select
        Selection.MergeCells = True
        Range("Q3:Q4").Select
        Selection.MergeCells = True
        Range("R4").value = "IW"
        Range("R5").value = "/Sys/Yr"
        Range("S4").value = "OoW"
        Range("S5").value = "/Sys/Yr"
        Range("R5").Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.FormulaR1C1 = "=(R[6]C[-13]+R[6]C[-12])"
        fstadd = ActiveCell.Address
        Selection.Copy
        ActiveCell.Offset(0, -1).Select
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(0, 1).Select
        lstadd = ActiveCell.Address
        Range(fstadd, lstadd).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    Range(fstadd, lstadd).Select
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
    Selection.NumberFormat = "0.000"
    Selection.FormatConditions(1).StopIfTrue = False
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
        Selection.EntireColumn.Select
        Selection.ColumnWidth = "8"
        
        
        Range("P5").Select
    Selection.Copy
    Range("Q5:S5").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S5").Select
    ActiveCell.FormulaR1C1 = "/Sys/Yr"
    Range("R5").Select
    Selection.Copy
    Range("R4:S4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("R4").Select
    Selection.Copy
    Range("Q3:Q4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Q4").Select
    Selection.Copy
    Range("R3:S3").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
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
    
    Range("Q3:Q4").Select
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
   ' AggPvtTableName = ActiveWorkbook.name
    Workbooks(myPvtWorkBook).Activate
Set pvtTbl = ActiveSheet.PivotTables("pvtKPIMASTER")
    outPutFilePath = ThisWorkbook.Path & "\"
    installFlName = outPutFilePath & "CTS_KPI_Summary.xlsx"
    Application.Workbooks.Open (installFlName), False 'false to disable link update message
    myWorkBook = ActiveWorkbook.name
    Workbooks(myWorkBook).Activate
    Sheets("KPI-Master").Select
    Cells.Select
    Selection.Delete
    Workbooks(myPvtWorkBook).Activate
    pvtTbl.TableRange2.Copy
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("KPI-Master").Select
    Range("a1").PasteSpecial
    
        
   
        Workbooks(myPvtWorkBook).Activate
        Range("Q1").Select
        Range("Q1:S1").Select
        Selection.EntireColumn.Select
        Selection.Copy
        Workbooks(myWorkBook).Activate
        Sheets("KPI-Master").Activate
        Range("Q1").PasteSpecial
    Workbooks(myWorkBook).Save
    Workbooks(myPvtWorkBook).SaveAs _
    fileName:=ThisWorkbook.Path & "\" & CTSProductName & "_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    AggPvtTableName = ActiveWorkbook.name
    
End Sub

Sub DataBrekUpFrPivot()

Sheets("Aggr. SWO Data CV").Select
Cells(1, 1).Select
ActiveCell.UnMerge
Cells(2, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(0, 1).Select
Selection.EntireColumn.Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(2, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(0, 1).Select
ActiveCell.value = "Part12NC-Sub Parts"

Application.CutCopyMode = False
ActiveCell.Offset(1, 0).Select
fstadd = ActiveCell.Address
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
lstadd = ActiveCell.Address
Cells(2, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
Selection.Offset(1, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""All Aggregated"",RC[-1]=""Parts Aggregated"",RC[-1]=""Non-Parts Aggregated""),""-"",RC[-1])"
    Selection.AutoFill Destination:=Range(fstadd, lstadd)
    Range(fstadd, lstadd).Select
    Calculate
    Range(fstadd, lstadd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    Selection.Offset(1, 0).Select
'    Range("F2").Select
   ' ActiveCell.Offset(0, -1).Select
    
'    ActiveCell.End(xlDown).Select
 '   ActiveCell.Offset(1, 0).Select
  '  lstCellAdd = ActiveCell.Address
   ' Cells(2, 1).Select
    'Selection.EntireRow.Select
    'Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    'Selection.Offset(1, 0).Select
    'fstCellAdd = ActiveCell.Address
    Do Until ActiveCell.value = ""
    If ActiveCell.value = "All Aggregated" Then
    ActiveCell.Offset(1, 0).Select
    End If
    If ActiveCell.value = "Parts Aggregated" Then
    Do Until ActiveCell.value = "Non-Parts Aggregated"
    ActiveCell.value = "Parts Aggregated"
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "" Then
    Exit Do
    End If
    
    Loop
    ElseIf ActiveCell.value = "Non-Parts Aggregated" Then
    Do Until ActiveCell.value = "Parts Aggregated"
    ActiveCell.value = "Non-Parts Aggregated"
    ActiveCell.Offset(1, 0).Select

    If ActiveCell.value = "" Then
    Exit Do
    End If
    Loop
    End If
   
    Loop
    ActiveCell.Offset(0, 1).Select
    If ActiveCell.value = 0 Then
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    End If
    Cells(2, 1).Select
Selection.EntireRow.Select
Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.value = ""
If ActiveCell.value = "-" Then
ActiveCell.Offset(1, 0).Select
Else
ActiveCell.Offset(0, -1).value = "Parts/Non-Parts Breakups"
ActiveCell.Offset(1, 0).Select

End If
Loop
End Sub
