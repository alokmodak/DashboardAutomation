Attribute VB_Name = "CTS"
'================================================================
'Who     When     What

'================================================================


Public myworkbook As String
Public wsptName  As String
Public wbName1 As String
Public srcfile As String
Public inputFileGlobal As String
Public outputFileGlobal As String
Public myPvtWorkBook As String

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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer
 Application.ScreenUpdating = False
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
    prdNameFile = KPISheetName

'Open input file-Aggregated Data File

inputFileGlobal = prdNameFile & ".xlsx"
If Sheet1.rdbLocalDrive.value = True Then
inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
inputFlName = inputFileGlobal
End If

If Sheet1.rdbSharedDrive.value = True Then
SharedDrive_Path inputFileGlobal
inputPath = sharedDrivePath
inputFlName = inputFileGlobal
End If

Application.Workbooks.Open (inputPath), False
Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(inputFlName).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
   ' On Error Resume Next
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh

'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
ActiveWorkbook.Sheets(2).Activate
AggrDataShtName = ActiveSheet.name
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
   DataBrekUpFrPivotKPIALL
       Dim Part12NCClmn As Range
       Set Part12NCClmn = Sheets(AggrDataShtName).rows(2).Find("Part12NC", , , xlWhole, , , , False)
    
      If Not Part12NCClmn Is Nothing Then
        Application.ScreenUpdating = False
        Part12NCClmn.Offset(1, 0).Select
        Part12NcClmnAdd = ActiveCell.Address(False, False)
      End If
        
       Dim ttlCalls As Range
       Set ttlCalls = Sheets(AggrDataShtName).rows(2).Find("Total Calls (#)", , , xlWhole, , , , False)
    
      If Not ttlCalls Is Nothing Then
        Application.ScreenUpdating = False
        ttlCalls.Offset(1, 0).Select
        ttlCallsAdd = ActiveCell.Address(False, False)
      End If
        
       Dim AvgMTTRprCallHrs As Range
       Set AvgMTTRprCallHrs = Sheets(AggrDataShtName).rows(2).Find("Avg. MTTR/Call (hrs)", , , xlWhole, , , , False)
    
      If Not AvgMTTRprCallHrs Is Nothing Then
        Application.ScreenUpdating = False
        AvgMTTRprCallHrs.Offset(1, 0).Select
        AvgMTTRprCallHrsAdd = ActiveCell.Address(False, False)
      End If
            
       Dim visitsprCallNP As Range
       Set visitsprCallNP = Sheets(AggrDataShtName).rows(2).Find("# of calls with 1 visit", , , xlWhole, , , , False)
    
      If Not visitsprCallNP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallNP.Offset(1, 0).Select
        visitsprCallNPAdd = ActiveCell.Address(False, False)
      End If
      
       Dim visitsprCallP As Range
       Set visitsprCallP = Sheets(AggrDataShtName).rows(2).Find("Calls = 0 Visit", , , xlWhole, , , , False)
    
      If Not visitsprCallP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallP.Offset(1, 0).Select
        visitsprCallPAdd = ActiveCell.Address(False, False)
      End If
      
      Dim avgVisitsprCallNP As Range
       Set avgVisitsprCallNP = Sheets(AggrDataShtName).rows(2).Find("Avg. Visits/call (#)", , , xlWhole, , , , False)
    
      If Not avgVisitsprCallNP Is Nothing Then
        Application.ScreenUpdating = False
        avgVisitsprCallNP.Offset(1, 0).Select
        avgVisitsprCallAdd = ActiveCell.Address(False, False)
      End If
      
       Dim ttlCostprPartNP As Range
       Set ttlCostprPartNP = Sheets(AggrDataShtName).rows(2).Find("Total Costs/part (EUR)", , , xlWhole, , , , False)
    
      If Not ttlCostprPartNP Is Nothing Then
        Application.ScreenUpdating = False
        ttlCostprPartNP.Offset(1, 0).Select
        ttlCostprPartAdd = ActiveCell.Address(False, False)
      End If
      
'Add one column for "Total Cost of Parts & Non-Parts"

  Dim found As Range
  Set found = Sheets(AggrDataShtName).rows(2).Find("Total Costs/part (EUR)", , , xlWhole, , , , False)
    
    If Not found Is Nothing Then
        Application.ScreenUpdating = False
        found.Offset(, 1).Resize(, 1).EntireColumn.Insert
  
    End If
  
        Workbooks(myPvtWorkBook).Sheets(AggrDataShtName).Activate

        found.End(xlDown).Select
        ActiveCell.Offset(0, 1).Select
        ttlCstLstAdd = ActiveCell.Address
        found.Offset(, 1).value = "Total Cost of Parts & Non-Parts"
        found.Offset(1, 1).Select
        ttlCstAdd = ActiveCell.Address
   
        'ActiveCell.Offset(, 0).Formula = "=IFERROR(IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallPAdd & "*" & 200 & ")," & "IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Non-Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallNPAdd & "*" & 200 & "))),0)"
        ActiveCell.Offset(, 0).Formula = "=(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 105 & ")+(" & avgVisitsprCallAdd & "*" & 200 & ")+(" & ttlCostprPartAdd & "*" & ttlCallsAdd & ")"

        Range(ttlCstAdd).Select
        Selection.Copy
        Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteFormulas
        Range(ttlCstAdd, ttlCstLstAdd).Select
        Selection.Copy
        Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteValues
      
        ActiveWorkbook.Sheets(AggrDataShtName).Activate
        Cells(1, 1).Select
        ActiveCell.EntireRow.Select
        Selection.Delete
        Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A2").Select
        ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
        Range("A2").Select
        fstAdd = ActiveCell.Address
        ActiveCell.Offset(0, 1).Select
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(0, -1).Select
        lstAdd = ActiveCell.Address
        Range("A2").Select
        Selection.Copy
        Range(fstAdd, lstAdd).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Calculate
        Range(fstAdd, lstAdd).Select
        Selection.Copy
        Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
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
    
        
'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)
        Set wsPtTable = Worksheets(Sheets.Count)
        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets(AggrDataShtName).Activate
        Set wsData = Worksheets(AggrDataShtName)
        Worksheets(AggrDataShtName).Activate
        sourceSheet = ActiveSheet.name
        Cells(1, 1).Select
        fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select
        lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        Sheets(wsptName).Activate
        rngData = fstAdd & ":" & lstAdd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
'        Range(fstAdd).Select
        ActiveCell.PivotTable.name = "pvtKPIALL"
        wsPtTable.Activate
        Set pt = wsPtTable.PivotTables("pvtKPIALL")
       ' Set pf = pt.PivotFields("Period")
        'pf.Orientation = xlPageField
        'pf.Position = 1
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
        ActiveSheet.PivotTables("pvtKPIALL").CalculatedFields.Add _
        "Avg. MTTR/Call (hrs)/12", "='Avg. MTTR/Call (hrs)' /12", True
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Avg. MTTR/Call (hrs)/12"). _
        Orientation = xlDataField
        ActiveSheet.PivotTables("pvtKPIALL").DataPivotField.PivotItems( _
        "Sum of Avg. MTTR/Call (hrs)/12").Caption = "#Avg. MTTR/Call (hrs)/12"
        ActiveSheet.PivotTables("pvtKPIALL").CalculatedFields.Add "Avg. ETTR (days)/12" _
        , "='Avg. ETTR (days)' /12", True
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Avg. ETTR (days)/12"). _
        Orientation = xlDataField
        ActiveSheet.PivotTables("pvtKPIALL").DataPivotField.PivotItems( _
        "Sum of Avg. ETTR (days)/12").Caption = "#Avg. ETTR (days)/12"
        ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "'#Avg. MTTR/Call (hrs)/12'", _
        xlDataAndLabel, True
        With ActiveSheet.PivotTables("pvtKPIALL").PivotFields( _
            "#Avg. MTTR/Call (hrs)/12")
            .NumberFormat = "0.00"
        End With
    
        ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "'#Avg. ETTR (days)/12'", _
        xlDataAndLabel, True
        With ActiveSheet.PivotTables("pvtKPIALL").PivotFields("#Avg. ETTR (days)/12")
            .NumberFormat = "0.00"
        End With
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Avg. Visits/call (#)"), "Visits/call (#)", xlAverage
        With ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Visits/call (#)")
            .NumberFormat = "0.00"
        End With
        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
    
        ActiveSheet.PivotTables("pvtKPIALL").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIALL").PivotFields("Total Cost of Parts & Non-Parts"), _
        "#Total Costs", xlSum
    
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
    
    ActiveSheet.PivotTables("pvtKPIALL").HasAutoFormat = False
    ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "", xlDataAndLabel, True
    Selection.ColumnWidth = 8
    ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "'Part12NC-Sub Parts'['-']" _
        , xlDataAndLabel, True
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts"). _
        ShowDetail = False
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail _
        = False
    
    With ActiveSheet.PivotTables("pvtKPIALL")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    pvtTbl.RefreshTable
    
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("SubSystem").ShowDetail = True
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        True

' Add ConditionalFormatting of Data Bars on total calls of Parts and Non parts
     ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        True
     ActiveSheet.PivotTables("pvtKPIALL").PivotFields("SubSystem").RepeatLabels = _
        True
     ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").RepeatLabels _
        = True
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 4).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    
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
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 1).Select
    
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 5).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    
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
    

    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        False
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
'
    'Range("E4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 4).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    'Range("F4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 5).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
  
   ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        True
   
'Add conditional formatting on MTTR and ETTR Calls
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

     'Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 6).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    lstadd1 = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & "," & 20 & ")"
        
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. MTTR/Call (hrs)/12"",$G$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    On Error Resume Next
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    
    'Range("H4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 7).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. ETTR (days)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    'Range("I4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 8).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstJAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & "," & 10 & ")"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",$A$3,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
     Range(fstAdd, lstJAdd).Select
     Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstJAdd & "," & 20 & ")"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",$A$3,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""#Avg. ETTR (days)/12"",$A$3,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="Visits/call (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    'Range("K4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 10).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstLAdd = ActiveCell.Address
    
    Range(fstAdd, lstLAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""Visits/call (#)"",$K$4,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""Visits/call (#)"",$L$4,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Total Costs", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    PfstAdd = Range(fstAdd).Offset(0, 1).Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 12).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    PlstAdd = Range(lstAdd).Offset(0, 1).Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$M$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range(PfstAdd, PlstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$N$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        False
   ActiveSheet.UsedRange.Find(what:="#Total Costs", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    PfstAdd = Range(fstAdd).Offset(0, 1).Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 12).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    PlstAdd = Range(lstAdd).Offset(0, 1).Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$M$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range(PfstAdd, PlstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$N$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 6).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & "," & 20 & ")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    'Range("H4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 7).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. ETTR (days)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
   ' Range("I4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 8).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstJAdd = ActiveCell.Address
    
    Range(fstAdd, lstAdd).Select
     Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & "," & 10 & ")"
   ' Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. MTTR/Call (hrs)/12"",R3C1,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
     Range(fstAdd, lstJAdd).Select
      Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstJAdd & "," & 10 & ")"
   ' Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. MTTR/Call (hrs)/12"",R3C1,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="Visits/call (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

    'Range("K4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 10).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstLAdd = ActiveCell.Address
    
    Range(fstAdd, lstLAdd).Select
    Range(fstAdd, lstLAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""Visits/call (#)"",$K$3,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""Visits/call (#)"",$L$3,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range(fstAdd, lstAdd).Select
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
    
    Columns("A:D").Select
    With Selection
        .ColumnWidth = 8
    End With
    Cells(1, 1).Select
    
    Worksheets(wsptName).PivotTables("pvtKPIALL").PreserveFormatting = False
    Sheets(wsptName).name = "PivotTableAggData"
    
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
        True
    
     ActiveWindow.Zoom = 85
     Workbooks(myPvtWorkBook).Activate
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
    
'Open Output file CTS_KPI_Summary.xlsx
    outputFileGlobal = "CTS_KPI_Summary.xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        outputPath = ThisWorkbook.Path & "\" & outputFileGlobal
        outputFlName = outputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path outputFileGlobal
        outputPath = sharedDrivePath
        outputFlName = outputFileGlobal
    End If

    Application.Workbooks.Open (outputPath), False
    Application.Workbooks(outputFileGlobal).Windows(1).Visible = True
    
    myCTSWorkBook = ActiveWorkbook.name
    
    Workbooks(myCTSWorkBook).Activate
    Sheets("KPI-All").Select
    'ActiveSheet.Shapes.Range(Array("Period 1")).Select
    'Selection.Delete

'Delete Slicers
    Dim slCaches As SlicerCaches
    Dim slCache As SlicerCache

    Set slCaches = ActiveWorkbook.SlicerCaches

    For Each slCache In slCaches

        If slCache.PivotTables.Item(1).name = "pvtKPIALL" Then
           slCache.Delete
        End If

    Next slCache
    
    Cells.Select
    Selection.Delete
    Workbooks(myPvtWorkBook).Activate
    pvtTbl.TableRange2.Copy
    Workbooks(myCTSWorkBook).Activate
    Sheets("KPI-All").Select
    Range("a1").PasteSpecial
     
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtKPIALL"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
'Add Headings to DashBoard
    
    Range("A1").Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ActiveCell.FormulaR1C1 = "KPI-All Dash Board for " & KPISheetName
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Offset(-3, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
        .Font.Bold = True
        .Font.Italic = True
        .Font.name = "Calibri"
        .Font.Size = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.399975585192419
        .Interior.PatternTintAndShade = 0
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("A1").Select
    ActiveCell.EntireRow.Select
    Selection.RowHeight = 25
    Range("A2").Select
    Workbooks(outputFlName).Activate
    ActiveWorkbook.Save
    Workbooks(myPvtWorkBook).Close
  
'IB value for 12 months from IB sheet
    Workbooks(outputFlName).Activate
    Sheets("IB").Select
    Range("N1").Select
    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")

    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    pvtTbl.PivotFields("ProductGroup").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    pvtTbl.PivotFields("ProductGroup").PivotFilters.Add Type:=xlCaptionContains, Value1:=KPISheetName
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Add RRR% and CallRate Columns
    'Workbooks(outputFlName).Activate
    Sheets("KPI-All").Select
    ActiveSheet.UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "RRR%"
    ActiveCell.Offset(0, 1).value = "CallRate/Sys/Yr"
    Range(ActiveCell.Address, ActiveCell.Offset(0, 1).Address).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 10.14
    
    Application.ScreenUpdating = False
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("pvtKPIALL"), "Period" _
        ).Slicers.Add ActiveSheet, , "Period 1", "Period", 69.75, 950, 144, 700
    ActiveSheet.Shapes.Range(Array("Period 1")).Select
   ' ActiveSheet.Shapes("Period 1").IncrementLeft 490
    ActiveSheet.Shapes("Period 1").IncrementTop -49.25
        
        'With ActiveWorkbook.SlicerCaches("Slicer_Period1")
         '   For i = 1 To .SlicerItems.Count
          '      .SlicerItems(i).selected = True
          '  Next i

        'End With
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
           
    ActiveSheet.UsedRange.Find(what:="CallRate/Sys/Yr", lookat:=xlWhole).Select
        'Range("P4").Select
    ActiveCell.Offset(1, 0).Select
        ActiveCell.FormulaR1C1 = "=(R[]C[-11]+R[]C[-10])/" & IBVal
        fstAdd = ActiveCell.Address
        Selection.Copy
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select

        ActiveCell.Offset(0, 15).Select
        lstAdd = ActiveCell.Address
        Range(fstAdd, lstAdd).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range(fstAdd, lstAdd).Select
    
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
    
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 20
        .Percent = False
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
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
   Application.CutCopyMode = False

   Range("N3").Select
   Selection.Copy
   ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
   Range(ActiveCell.Address, ActiveCell.Offset(0, 1).Offset(-2, 0).Address).Select
'   Range("O3:P3").Select
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

'Create RRR% column and calculate RRR%
     Range("A2").Select
    pvtmstName = ActiveCell.PivotTable.name
    Set pvtMstTbl = ActiveSheet.PivotTables(pvtmstName)
    ActiveSheet.PivotTables(pvtmstName).PivotSelect "", xlDataAndLabel, True
    Selection.Copy

    Range("AD1").Select
    ActiveSheet.Paste
    pvtName = ActiveCell.PivotTable.name
    
    Set pvtTbl = ActiveSheet.PivotTables(pvtName)

    pvtTbl.DataPivotField.Orientation = xlHidden
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC").Orientation = _
        xlHidden

    ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Total Calls (#)"), "# of Calls", xlSum
    ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Calls = 0 Visit"), "Sum of Calls = 0 Visit", xlSum
    ActiveSheet.PivotTables(pvtName).PivotFields("Sum of Calls = 0 Visit"). _
        Caption = "#Calls = 0 Visit"
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").ShowAllItems = _
        True
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").ClearAllFilters
'Add RRR% calculated field in Pivot table

        pvtTbl.CalculatedFields.Add "RRR", _
        "='Calls = 0 Visit' /'Total Calls (#)' *100", True
        pvtTbl.PivotFields("RRR").Orientation = _
        xlDataField
        
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").ClearAllFilters
    ActiveSheet.PivotTables(pvtName).PivotSelect "", xlDataAndLabel, True
    Selection.Cut
    Range("AD3").Select
    ActiveSheet.Paste
    ActiveSheet.PivotTables(pvtName).PivotSelect "Subsystem['-']", xlDataAndLabel, True
    ActiveCell.Offset(0, -1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(-1, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],RC[2],RC[3])"
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd).Select
    ActiveCell.Offset(0, -1).Select
    fstpvtAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstPvtAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Sum of RRR", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    RRRLstRAdd = ActiveCell.Offset(-1, 0).Address
    Range(fstpvtAdd).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-27],RC[-26],RC[-25])"
    Selection.AutoFill Destination:=Range(fstpvtAdd, lstPvtAdd)
    ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    vlkpRRRFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    vlkpRRRLstAdd = ActiveCell.Address
    
    fstpvtAdd = Range(fstpvtAdd).Address(False, False)
    Range(vlkpRRRFstAdd).Select
    Range(vlkpRRRFstAdd).Formula = "=IFERROR(VLOOKUP(" & fstpvtAdd & ":" & lstPvtAdd & "," & fstAdd & ":" & RRRLstRAdd & ",8,0)" & "," & 0 & ")"
    
    Range(vlkpRRRFstAdd).Select
    Selection.AutoFill Destination:=Range(vlkpRRRFstAdd, vlkpRRRLstAdd)

    Application.CutCopyMode = False
    
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
'Apply Icon Set Conditional formatting on RRR Column Values
    ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select

    fstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
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
    Range(fstAdd, lstAdd).Select
    Selection.NumberFormat = "0"
    
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
  
'Filter the KPI-All data for previous 12 months (month selected by the user will be considered as start of the month)
    ActiveSheet.Shapes.Range(Array("Period 1")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Period1").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables(pvtName))
        
    Dim slSlicer As Slicer
    'Dim slCache As SlicerCache
    Dim slItem As slicerItem

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    ActiveSheet.Shapes.Range(Array("Period 1")).Select
   
    Application.ScreenUpdating = False
    With ActiveWorkbook.SlicerCaches("Slicer_Period1")
        '--deselect all items except the 12 months
        For Each slItem In .VisibleSlicerItems
            If slItem.name < endDate Or slItem.name > fixedDate Then
                slItem.selected = False
                Else
                slItem.selected = True
                End If
        Next slItem

End With
    Cells.Select
    Selection.ColumnWidth = 8.5
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("AA1:AP1").Select
    Selection.EntireColumn.Select
    Selection.EntireColumn.Hidden = True
    
    ActiveWindow.Zoom = 100

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Application.ActiveWorkbook.Save
  
End Sub
Sub DataBrekUpFrPivotKPIALL()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Sheets(AggrDataShtName).Select
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
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstAdd = ActiveCell.Address
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    Selection.Offset(1, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""All Aggregated"",RC[-1]=""Parts Aggregated"",RC[-1]=""Non-Parts Aggregated""),""-"",RC[-1])"
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
    Selection.Offset(1, 0).Select
    Do Until ActiveCell.value = ""
        If ActiveCell.value = "-" Then
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.Offset(0, -1).value = "Parts Aggregated"
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
    ActiveCell.Offset(0, 1).Select
    If ActiveCell.value = 0 Then
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
    End If
    
    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
End Sub
Public Sub CRRateCalculationNew()
Dim fixedDate, mypath, CTSExcel, CTSWorkBook, pvtExcel, myPvtWorkBook As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1, KPISheetName As String
Call IBPivotTable
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    
'Delete Pivot tables from aggregated Data file if any
    For Each xWs In Application.ActiveWorkbook.Worksheets
        For Each xpvt In xWs.PivotTables
            xWs.Range(xpvt.TableRange2.Address).Delete Shift:=xlUp
        Next
    Next
    
    fixedDate = Sheet1.combYear.value
    
    Workbooks("CTS_KPI_Summary.xlsx").Activate
    endDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    Sheets("IB").Select
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
            Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = fixedDate
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            IBVal = ActiveCell.Offset(0, 1).value
       
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    CTSProductName = KPISheetName
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    'Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
    Workbooks("CTS_KPI_Summary.xlsx").Activate
    myworkbook = ActiveWorkbook.name
    Sheets("KPI-All").Select
    
'    fixedDate = Sheet1.combYear.value
'    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
'    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
'    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIALL")
'    Set pf = pvtTbl.PivotFields("Period")
'    pf.ClearAllFilters
'    With ActiveSheet.PivotTables("pvtKPIALL")
'        .ColumnGrand = True
'        .RowGrand = False
'    End With
'    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
'        If pvtItm < endDate Or pvtItm > fixedDate Then
'            pvtItm.Visible = False
'        Else
'            pvtItm.Visible = True
'        End If
'    Next pvtItm
    
    
    
    
     ActiveSheet.Shapes.Range(Array("Period 1")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Period1").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("pvtKPIALL"))
        
    Dim slSlicer As Slicer
    Dim slCache As SlicerCache
    Dim slItem As slicerItem

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    ActiveSheet.Shapes.Range(Array("Period 1")).Select
   
    Application.ScreenUpdating = False
    With ActiveWorkbook.SlicerCaches("Slicer_Period1")
        '--deselect all items except the 12 months
        For Each slItem In .VisibleSlicerItems
            If slItem.name < endDate Or slItem.name > fixedDate Then
                slItem.selected = False
                Else
                slItem.selected = True
                End If
        Next slItem

End With
    
    Sheets("CR").Select
    Range("A:A").Select
    'On Error Resume Next
    Selection.EntireRow.Select
    Selection.EntireRow.Delete
    Application.Columns.Ungroup
    rows("1:1").Select
        
    Sheets("KPI-All").Select
    Range("A2").Select
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("SubSystem").RepeatLabels = _
    False
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").RepeatLabels _
    = False
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    False
    ActiveSheet.PivotTables("pvtKPIALL").PivotSelect "", xlDataAndLabel, True
    'pvtTbl.TableRange2.Select
    'pvtTbl.TableRange1.Copy
    Selection.Copy
    Sheets("CR").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Sheets("KPI-All").Select
    Range("A2").Select
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("SubSystem").RepeatLabels = _
    True
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").RepeatLabels _
    = True
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    True
    Sheets("CR").Select
    Range("1:1").Select
    Selection.EntireRow.Delete
    Sheets("CR").UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstclmn = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lstclmnAdd = ActiveCell.Address
    Range(fstclmn, lstclmnAdd).Select
    Selection.EntireColumn.Select
    Selection.EntireColumn.Delete
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Sheets("CR").UsedRange.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
    deleteClmnsAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="PartDescription", lookat:=xlWhole).Select
    deleteLstClmnsAdd = ActiveCell.Address
    Range(deleteClmnsAdd, deleteLstClmnsAdd).Select
    Selection.EntireColumn.Select
    Selection.EntireColumn.Delete
       
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "Designed"
    ActiveCell.Offset(1, 0).FormulaR1C1 = "CR / Sys / Yr"
    rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "All Systems"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "All Blocks"
    'Sheets("Designed Data").Activate
    'Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select

    'Selection.EntireColumn.Select
    'Selection.Copy
    'Sheets("CR").Activate
    'Range("C1").Select
    'ActiveSheet.Paste
    'Application.CutCopyMode = False
    Sheets("CR").Activate
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    Selection.value = "DataFill"
    Range("A3").Select
    
    ActiveCell.FormulaR1C1 = "=IF(RC[1]="""",R[-1]C,RC[1])"
    fstAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("CR").UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.Replace what:="", Replacement:="0", lookat:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Sheets("CR").UsedRange.Find(what:="Designed", lookat:=xlWhole).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(2, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 3).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[-3]&RC[-1]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Cells(2, 4).value = "SS&BB"
    Sheets("CR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 1).Select
    CRSysYrFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -4).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    CRSysYrLstAdd = ActiveCell.Address
    Range(CRSysYrFstAdd).Select
    ActiveCell.Formula = "=" & ActiveCell.Offset(0, 4).Address(False, False)
    'ActiveCell.FormulaR1C1 = "=(RC[-1]+RC[-2])/" & IBVal
    Range(CRSysYrFstAdd).Select
    Selection.AutoFill Destination:=Range(CRSysYrFstAdd, CRSysYrLstAdd)
    Range(CRSysYrFstAdd, CRSysYrLstAdd).Select
    Range(CRSysYrFstAdd, CRSysYrLstAdd).NumberFormat = "0.0000"
    Calculate
    Application.CutCopyMode = False
    Sheets("CR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    nonPartsFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    nonPartslstAdd = ActiveCell.Address
    Range(nonPartsFstAdd, nonPartslstAdd).Select
    Selection.NumberFormat = "0.00"
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
    
    Sheets("CR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
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
        
'    Range(fstCellAdd, add1).Select
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
        .NumberFormat = "0.00"

    End With
    
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    Sheets("CR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.value = "MAT # of Calls profiles"
    Range("F1:H1").Select
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

    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = "Non-Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "CR/Sys/Yr"
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
    
    Range("I2").Select
    ActiveCell.value = "ITM"
    Range("J2").Select
    ActiveCell.value = "IMQ"
    Range("K2").Select
    ActiveCell.value = "YTD"
    Range("L2").Select
    ActiveCell.value = "MAT"
    Range("I1:L1").Select
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
   
    Range("M2").Select
    ActiveCell.value = "ITM"
    Range("N2").Select
    ActiveCell.value = "IMQ"
    Range("O2").Select
    ActiveCell.value = "YTD"
    Range("P2").Select
    ActiveCell.value = "MAT"
    Range("M1:P1").Select
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
    Range("Q1").Select
    ActiveCell.value = "Crossover"
    Range("Q2").Select
    ActiveCell.value = "Trigger"
    Range("R2").Select
      
    fixedDate1 = Sheet1.combYear.value
    fixDte = Format(fixedDate1, "mmm" & "-" & "yyyy")
    fixDate2 = Format(DateAdd("yyyy", -1, fixedDate1), "mmm" & "-" & "yyyy")
    frmtData = Format(DateAdd("m", 1, fixedDate1), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("m", -12, frmtData), "mmm" & "-" & "yyyy")
    endDate2 = Format(DateAdd("m", -24, frmtData), "mmm" & "-" & "yyyy")

    fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    fnlEndDate1 = Format(endDate2, "mmm" & "-" & "yyyy")
    frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
    Range("M1").Select
    ActiveCell.value = "Last Year"
    Range("I1").Select
    ActiveCell.value = "Current Year CR/Sys"
    Range("R2").Select
    Do Until endDate1 = frmtData
        ActiveCell.value = endDate1
        ActiveCell.NumberFormat = "[$-409]mmm-yy;@"
        ActiveCell.Offset(0, 1).Select
        endDate1 = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    Loop

    Range("A1").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(-1, 0).Select
    upAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15652757
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("CR").UsedRange.Find(what:="Crossover", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 1).Select
    up1Add = ActiveCell.Address
    Range(up1Add).Select
    ActiveCell.value = "CR/Sys/Month"
    Range(up1Add, upAdd).Select
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
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    
    ActiveSheet.UsedRange.Select
    Selection.RowHeight = 15
    Range("I1:Q2").Select
    Selection.Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlRight
    End With
    
    Cells(2, 1).Select
    Sheets("CR").UsedRange.Find(what:="CR/Sys/Yr", lookat:=xlWhole).Select
    Sheets("CR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 7
    Application.ActiveWorkbook.Save
    Call fourthRowCalcCR
    Call CRPivotTableNew
    
'Enter Designed values from Designed Data sheet into CR sheet

    Sheets("Designed Data").Select
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(2, -1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Offset(2, 0).Select
    CRValFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    CRValLstAdd = ActiveCell.Address
    Sheets("CR").Select
    Sheets("CR").UsedRange.Find(what:="CR / Sys / Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    DVVlkpResultAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    DVVlkpFstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
'    ActiveCell.Offset(0, 1).Select
    DVVlkpLstAdd = ActiveCell.Address
    fstAddRef = Range(DVVlkpFstAdd).Address(False, False)

    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'Designed Data'!R4C19:R98C22,4,0)"
    Range(DVVlkpResultAdd).Formula = "=IFERROR(VLOOKUP(" & fstAddRef & "," & "'Designed Data'!" & fstAdd & ":" & CRValLstAdd & ",4,0)" & "," & 0 & ")"
    Range(DVVlkpResultAdd).Select
    
    Selection.AutoFill Destination:=Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address)
    Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address).Select
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.NumberFormat = "0.0000"
    Dim visPvtItm As String
    
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    fixedDate = Sheet1.combYear.value
    
'======================================================================
    
'Enter 12 month's data in column "Call/Sys/Month" after "crossover Trigger"
    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Sheets("CR").Select
    Set pvtTbl = ActiveSheet.PivotTables("pvtCR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 18).Select
    i = 18
       
    Sheets("CR").Select
    
    
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("CR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(2, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            abc = pf.CurrentPage
            lr = Worksheets("CR").Cells(rows.Count, "D").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(True, True)
            Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
            ActiveCell.Offset(1, -1).Select
            ActiveCell.End(xlDown).Select
'            ActiveCell.Offset(0, 6).Select
            lstAdd = ActiveCell.Address(True, True)
            rng = Range(fstAdd, lstAdd)
            Range("D3").Select
            vlkupAFstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            vlkupALstAdd = ActiveCell.Address(True, True)
            '#Calls (#)
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
            
            Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            IBVal = ActiveCell.Offset(0, 1).value
    
            Sheets("CR").Select
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & "/" & 12 & ")" & "," & 0 & ")"
                    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
                    Cells(X, myCol).Copy
                    Cells(X, myCol).PasteSpecial xlValues
                    Cells(X, myCol).NumberFormat = "0.0000"
                Next X
             
            End If
                i = i + 1
        
        End If
      
    Next pvtItm
    
    Range("D3").Select
    ActiveCell.Offset(1, 2).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("" & "R" & sumMidAdd - 1 & ":" & "AC" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    
'===========================================================================
'Calculate ITM i.e. CR Value for the same motnh in the same year as inPut given from user
    ActiveSheet.UsedRange.Find(what:="CR/Sys/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Formula = "=(" & "AC" & sumMidAdd & ")"
    Range(fstAdd).Select
    Selection.Copy
    Range("" & "I" & sumMidAdd & ":" & "I" & sumMidAdd1 & "").PasteSpecial xlPasteFormulas
    Selection.NumberFormat = "0.0000"
       
'Calculate IMQ for the quarter in the last (Previous) year
'IB for IMQ
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")
    pvtTbl.PivotFields("Period").ClearAllFilters

    previousMonth = Format(DateAdd("m", -1, startDate), "yyyy" & "-" & "mm")
    qMnth = Format(DateAdd("m", -2, startDate), "yyyy" & "-" & "mm")

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
 
        End If
 
    Next
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate IMQ
startDate = Format(fixedDate, "yyyy" & "-" & "mm")
'prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

previousMonth = Format(DateAdd("m", -1, startDate), "yyyy" & "-" & "mm")
qMnth = Format(DateAdd("m", -2, startDate), "yyyy" & "-" & "mm")

For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
 If pvtItm.value = startDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
 pvtItm.Visible = True
 Else
 pvtItm.Visible = False
 
End If
 
Next

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 1).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    'On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X

'Calculate YTD i.e. CR data for the last 12 months from the input date provided by user
'IB for YTD
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, fixedDate), "yyyy" & "-" & "mm")

   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate YTD
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, fixedDate), "yyyy" & "-" & "mm")

Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 2).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    'On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X
    
'Calculate MAT i.e. CR data for the data till Januray of the selected year till the month of the year selected by the user as an input date
'IB for MAT
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    'startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate MAT
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")


Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 3).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X

'calculate ITM for the same month in the previous year of of the input year provided by the user
'IB for ITM
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    'startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    'endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm = endDate1 Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate ITM
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm = endDate1 Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 0).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X
     
'Calculate IMQ i.e. previous 3 months in the previous year
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
    qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

'IB for IMQ
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
      
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = prvsIMQ Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
            pvtItmName = pvtItm.value
        Else
            pvtItm.Visible = False
 
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate IMQ
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    prvsIMQ = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    previousMonth = Format(DateAdd("m", -1, prvsIMQ), "yyyy" & "-" & "mm")
    qMnth = Format(DateAdd("m", -2, prvsIMQ), "yyyy" & "-" & "mm")

    
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = prvsIMQ Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
            pvtItmName = pvtItm.value
        Else
            pvtItm.Visible = False
 
        End If
    Next pvtItm

    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 1).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X
    
'"""""""""""""""""""""""""""""""""""""""""""""
'Calculate MAT for the Previous year
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -2, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")

'IB for MAT
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
      
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
     For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
            pvtItmName = pvtItm.value
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate MAT
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -2, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
            pvtItmName = pvtItm.value
        End If
    Next pvtItm
    
    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 3).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X
'"""""""""""""""""""""""""""""""""""""""""""""""""


'Calculate YTD for the previous year i.e. CR data from January to the month in the previous year
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, startDate1), "yyyy" & "-" & "mm")

'IB for YTD
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
      
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
     For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate1 Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
            pvtItmName = pvtItm.Caption
        End If
    Next pvtItm
            
            'Sheets("IB").PivotTables("pvtIB").PivotFields("Period").CurrentPage = abc
            ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        
            abc = ActiveCell.Offset(0, 1).value
            IBVal = abc / 12
'Calculate YTD
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, startDate1), "yyyy" & "-" & "mm")

Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

     For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate1 Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
            pvtItmName = pvtItm.Caption
        End If
    Next pvtItm
    
    lr = Worksheets("CR").Cells(rows.Count, "C").End(xlUp).Row
    'rng = Range("AE3:AJ91")
    Sheets("CR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(2, 2).Select
    myRow = ActiveCell.Row
    myCol = ActiveCell.Column
    Range("AE3").Select
    fstAdd = ActiveCell.Address(True, True)
    Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    ActiveCell.End(xlDown).Select
'   ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(True, True)
    rng = Range(fstAdd, lstAdd)
    Range("D3").Select
    vlkupAFstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    vlkupALstAdd = ActiveCell.Address(True, True)
    For X = myRow To lr
    On Error Resume Next
    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & ")" & "," & 0 & ")"
    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
    Cells(X, myCol).Copy
    Cells(X, myCol).PasteSpecial xlValues
    Cells(X, myCol).NumberFormat = "0.0000"

    'Cells(X, 10).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False)) / IBVal
    'Cells(X, 10).NumberFormat = "0.00"
    Next X
   
'""""""""""""""""""""""""""""""""""""""""""""""""""
    Range("D3").Select
    ActiveCell.Offset(1, 2).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("F3").Select
    ActiveCell.Formula = "=SUM(" & sumAdd & ":" & sumAdd1 & ")"
    Selection.NumberFormat = "0.00"
    Range("G3").Select
    ActiveCell.Formula = "=SUM(" & "G" & sumMidAdd & ":" & "G" & sumMidAdd1 & ")"
    Selection.NumberFormat = "0.00"

    'Range("H3").Select
    'ActiveCell.Formula = "=SUM(" & "H" & sumMidAdd & ":" & "H" & sumMidAdd1 & ")"
    'Range("I3").Select
    'ActiveCell.Formula = "=SUM(" & "I" & sumMidAdd & ":" & "I" & sumMidAdd1 & ")"
    'Range("J3").Select
    'ActiveCell.Formula = "=SUM(" & "J" & sumMidAdd & ":" & "J" & sumMidAdd1 & ")"
    'Range("K3").Select
    'ActiveCell.Formula = "=SUM(" & "K" & sumMidAdd & ":" & "K" & sumMidAdd1 & ")"
    'Range("L3").Select
    'ActiveCell.Formula = "=SUM(" & "L" & sumMidAdd & ":" & "L" & sumMidAdd1 & ")"
    'Range("M3").Select
    'ActiveCell.Formula = "=SUM(" & "M" & sumMidAdd & ":" & "M" & sumMidAdd1 & ")"
    'Range("N3").Select
    'ActiveCell.Formula = "=SUM(" & "N" & sumMidAdd & ":" & "N" & sumMidAdd1 & ")"
    'Range("O3").Select
    'ActiveCell.Formula = "=SUM(" & "O" & sumMidAdd & ":" & "O" & sumMidAdd1 & ")"
    'Range("P3").Select
    'ActiveCell.Formula = "=SUM(" & "P" & sumMidAdd & ":" & "P" & sumMidAdd1 & ")"
    'Range("R3").Select
    'ActiveCell.Formula = "=SUM(" & "R" & sumMidAdd & ":" & "R" & sumMidAdd1 & ")"
    'Range("S3").Select
    'ActiveCell.Formula = "=SUM(" & "S" & sumMidAdd & ":" & "S" & sumMidAdd1 & ")"
    'Range("T3").Select
    'ActiveCell.Formula = "=SUM(" & "T" & sumMidAdd & ":" & "T" & sumMidAdd1 & ")"
    'Range("U3").Select
    'ActiveCell.Formula = "=SUM(" & "U" & sumMidAdd & ":" & "U" & sumMidAdd1 & ")"
    'Range("V3").Select
    'ActiveCell.Formula = "=SUM(" & "V" & sumMidAdd & ":" & "V" & sumMidAdd1 & ")"
    'Range("W3").Select
    'ActiveCell.Formula = "=SUM(" & "W" & sumMidAdd & ":" & "W" & sumMidAdd1 & ")"
    'Range("X3").Select
    'ActiveCell.Formula = "=SUM(" & "X" & sumMidAdd & ":" & "X" & sumMidAdd1 & ")"
    'Range("Y3").Select
    'ActiveCell.Formula = "=SUM(" & "Y" & sumMidAdd & ":" & "Y" & sumMidAdd1 & ")"
    'Range("Z3").Select
    'ActiveCell.Formula = "=SUM(" & "Z" & sumMidAdd & ":" & "Z" & sumMidAdd1 & ")"
    'Range("AA3").Select
    'ActiveCell.Formula = "=SUM(" & "AA" & sumMidAdd & ":" & "AA" & sumMidAdd1 & ")"
    'Range("AB3").Select
    'ActiveCell.Formula = "=SUM(" & "AB" & sumMidAdd & ":" & "AB" & sumMidAdd1 & ")"
    'Range("AC3").Select
    'ActiveCell.Formula = "=SUM(" & "AC" & sumMidAdd & ":" & "AC" & sumMidAdd1 & ")"
    Sheets("CR").Select
    Range("Q3").Select
   ActiveCell.FormulaR1C1 = "=AND(RC[-7]>RC[-6],RC[-12]<RC[-9])"

    Range("Q3").Select
    Selection.AutoFill Destination:=Range("" & "Q" & sumMidAdd - 1 & ":" & "Q" & sumMidAdd1 & "")
    Range("" & "Q" & sumMidAdd - 1 & ":" & "Q" & sumMidAdd1 & "").Select
    Calculate
    Range("" & "R" & sumMidAdd - 1 & ":" & "AC" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    Application.CutCopyMode = False
    Range("AD1").Select
    Selection.EntireColumn.Select
     Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     Range("" & "AD" & sumMidAdd - 1 & ":" & "AD" & sumMidAdd1 & "").Select
    Selection.SparklineGroups.Add Type:=xlSparkLine, SourceData:= _
        "" & "R" & sumMidAdd - 1 & ":" & "AC" & sumMidAdd1 & ""
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
    
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Application.CutCopyMode = False
    Range("H4").Select
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
    
'Highlight crossover Trigger in Red or Orange for true values depending upon 20% largest values
    Sheets("CR").Select
    Sheets("CR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 3).Select
    fstAdd = ActiveCell.Address(False, False)
    fstMidAdd = Mid(fstAdd, 2)
    
    Sheets("CR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    lstAdd = ActiveCell.Address(False, False)
    lstMidAdd = Mid(lstAdd, 2)
    ActiveCell.Offset(1, 0).Select
   
    ActiveCell.Formula = "=LARGE(" & "F" & fstMidAdd & ":" & "G" & lstMidAdd & ",20)"
    topTwentyVal = ActiveCell.value
    Sheets("CR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "True" Then
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
    End If
    
    ActiveCell.Offset(1, 0).Select
    Do While ActiveCell.value <> ""
    If ActiveCell.value = "True" Then
        If Range("F" & fstMidAdd).value >= topTwentyVal Or Range("G" & fstMidAdd).value >= topTwentyVal Then
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        Else
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 1484526
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        End If
        
    End If
    fstMidAdd = fstMidAdd + 1
    ActiveCell.Offset(1, 0).Select
    
    Loop
    Range(lstAdd).Offset(1, 0).Select
    Selection.ClearContents

    Range("A1").Select
    Selection.EntireColumn.Delete
    Range("C1").Select
    Selection.EntireColumn.Delete
    Sheets("CR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 1).Select
    fstAdd = ActiveCell.Address
    
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 8).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 1).Select
    deleteFnlAdd = ActiveCell.Address
    ActiveCell.Offset(0, 13).Select
    deleteFnlCAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    deleteFnlRAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(0, -1).Select
    fstRwAdd = ActiveCell.Address
    Range(fstAdd & ":" & fstRwAdd, lstAdd).Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    Range(deleteFnlAdd & ":" & deleteFnlCAdd, deleteFnlRAdd).Select
    Selection.ClearContents
   ' ActiveWindow.Zoom = 85
    Sheets("CR").UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.Replace what:="0", Replacement:="", lookat:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Sheets("CR").UsedRange.Find(what:="DataFill", lookat:=xlWhole).Select
    fstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 6).Select
    lstAdd = ActiveCell.Address(False, False)
    Range(fstAdd, lstAdd).Select
    Selection.ClearContents
    Sheets("CR").Select
    Set pvtTbl = ActiveSheet.PivotTables("pvtCR")
    pvtTbl.TableRange1.Select
    pvtTbl.TableRange2.Clear
    Cells(1, 1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    With Selection.Font
        .name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.EntireColumn.Select
    Sheets("CR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 10.57
'Add Heading to DashBoard
    Sheets("CR").Select
    rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Call Rate Dash Board for " & KPISheetName
    Range("A1").Select
    fstAdd = ActiveCell.Address
    Sheets("CR").UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    ActiveCell.Offset(-2, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
        .Font.Bold = True
        .Font.Italic = True
        .Font.name = "Calibri"
        .Font.Size = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.399975585192419
        .Interior.PatternTintAndShade = 0
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Merge
    Selection.Font.Bold = True
     rows("1:1").Select
    Selection.RowHeight = 25
    Range("A1").Select
    ActiveWindow.Zoom = 100
    Workbooks(myPvtWorkBook).Close
    Application.Workbooks(myworkbook).Save
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
End Sub
Public Function CRPivotTableNew()
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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
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
    Sheets(AggrDataShtName).Activate
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
    Range(fstFiltCellAdd, fstFiltCellAdd2).Select
    Range(fstFiltCellAdd, fstFiltCellAdd2).EntireRow.Delete
    ActiveSheet.ShowAllData
   
'Add a new sheet to create a Pivot Table
    Sheets.Add after:=Worksheets(Worksheets.Count)
    Set wsPtTable = Worksheets(Sheets.Count)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets(AggrDataShtName).Activate
    Set wsData = Worksheets(AggrDataShtName)
    Worksheets(AggrDataShtName).Activate
    sourceSheet = ActiveSheet.name
    ActiveSheet.Cells(1, 1).Select
    fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    Sheets(wsptName).Activate
    rngData = fstAdd & ":" & lstAdd
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
    TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion15
             
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtCR"
    wsPtTable.Activate
    Set pt = wsPtTable.PivotTables("pvtCR")
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
        
    With ActiveSheet.PivotTables("pvtCR").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("pvtCR").AddDataField ActiveSheet.PivotTables( _
    "pvtCR").PivotFields("Total Calls (#)"), "#Calls (#)", xlSum
        
    ActiveSheet.PivotTables("pvtCR").PivotFields("Part12NC").PivotItems( _
    "Non-Parts Aggregated").Caption = "Non-Parts"

    ActiveSheet.PivotTables("pvtCR").PivotFields("Part12NC").PivotItems( _
    "Parts Aggregated").Caption = "Parts"
       
    With ActiveSheet.PivotTables("pvtCR")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
    End With
    
    ActiveSheet.PivotTables("pvtCR").PivotFields("SubSystem").Subtotals = _
    Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
    ActiveSheet.PivotTables("pvtCR").PivotFields("BuildingBlock").Subtotals _
    = Array(False, False, False, False, False, False, False, False, False, False, False, False _
    )
    With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 1
    End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtCR")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtCR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    pvtTbl.RefreshTable
    
    Columns("A:E").EntireColumn.AutoFit
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
    ActiveSheet.PivotTables("pvtCR").Location = _
        "'[CTS_KPI_Summary.xlsx]CR'!$AK$3"
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("CR").Activate
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[5]"
    Range("AF3").Select
    Selection.Copy
    Range("AF3,AF91").Select
    Range("AF3,AF3:AJ91").Select
    ActiveSheet.Paste
    Range("AD2").Select
    ActiveCell.value = "DataFill"
    Range("AD3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[2]=0,R[-1]C,RC[2])"
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("AD3:AD91")
    Range("AE2").Select
    ActiveCell.value = "SS&BB"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],RC[2])"
    Range("AE3").Select
    Selection.AutoFill Destination:=Range("AE3:AE91")
   Application.ScreenUpdating = False
   
End Function
Sub fourthRowCalcCR()
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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
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


   
'Add a new sheet to create a Pivot Table
    Sheets.Add after:=Worksheets(Worksheets.Count)
    Set wsPtTable = Worksheets(Sheets.Count)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets(AggrDataShtName).Activate
    Set wsData = Worksheets(AggrDataShtName)
    Worksheets(AggrDataShtName).Activate
    sourceSheet = ActiveSheet.name
    ActiveSheet.Cells(1, 1).Select
    fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    Sheets(wsptName).Activate
    rngData = fstAdd & ":" & lstAdd
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
    TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion15
             
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtCR"
    wsPtTable.Activate
    Set pt = wsPtTable.PivotTables("pvtCR")
    Set pf = pt.PivotFields("System")
    pf.Orientation = xlPageField
    pf.Position = 1
    Set pf = pt.PivotFields("SubSystem")
    pf.Orientation = xlRowField
    pf.Position = 1
    Set pf = pt.PivotFields("BuildingBlock")
    pf.Orientation = xlRowField
    pf.Position = 2
    Set pf = pt.PivotFields("Period")
    pf.Orientation = xlRowField
    pf.Position = 3
    Set pf = pt.PivotFields("Part12NC")
    pf.Orientation = xlColumnField
    pf.Position = 1
        
    
    ActiveSheet.PivotTables("pvtCR").AddDataField ActiveSheet.PivotTables( _
    "pvtCR").PivotFields("Total Calls (#)"), "#Calls (#)", xlSum
        
    ActiveSheet.PivotTables("pvtCR").PivotFields("Part12NC").PivotItems( _
    "Non-Parts Aggregated").Caption = "Non-Parts"

    ActiveSheet.PivotTables("pvtCR").PivotFields("Part12NC").PivotItems( _
    "Parts Aggregated").Caption = "Parts"
       
    With ActiveSheet.PivotTables("pvtCR")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
    End With
    
    ActiveSheet.PivotTables("pvtCR").PivotFields("SubSystem").Subtotals = _
    Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
    ActiveSheet.PivotTables("pvtCR").PivotFields("BuildingBlock").Subtotals _
    = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("pvtCR").PivotFields("Period").Subtotals _
    = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtCR")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts", Value2:="All"
    With ActiveSheet.PivotTables("pvtCR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    
    Set pf = pvtTbl.PivotFields("BuildingBlock")
    pf.ClearAllFilters
    ActiveSheet.PivotTables("pvtCR").PivotFields("BuildingBlock").ClearAllFilters
    
    For Each pvtItm In pvtTbl.PivotFields("BuildingBlock").PivotItems
        If pvtItm = "Buildingblocks Aggregated" Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm
    
    ActiveSheet.PivotTables("pvtCR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtCR").PivotFields("System"). _
        EnableMultiplePageItems = False
    ActiveSheet.PivotTables("pvtCR").PivotFields("System").CurrentPage = "(All)"
    ActiveSheet.PivotTables("pvtCR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtCR").PivotFields("System").CurrentPage = _
        "System level"
    pvtTbl.RefreshTable
    
    Columns("A:F").EntireColumn.AutoFit
    
    shtname = ActiveSheet.name
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("All SS-BB").Select
    Cells.Select
    Selection.Delete
    
    Workbooks(myPvtWorkBook).Activate
    Sheets(shtname).Select
    Set pt = ActiveCell.PivotTable
    pt.PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("All SS-BB").Select
    Worksheets("All SS-BB").Range("AF3").PasteSpecial
    Workbooks(myPvtWorkBook).Activate
    Sheets(shtname).Select
    
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
    ActiveSheet.PivotTables("pvtCR").Location = _
        "'[CTS_KPI_Summary.xlsx]CR'!$AK$3"
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("CR").Activate
    
 'Enter 4th Rows values for 12 months individually
    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")

    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    With ActiveSheet.PivotTables("pvtCR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    Sheets("CR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select

'IB for 12 months
    Do While endDate <= fixedDate
        Sheets("IB").Select
        Range("N1").Select
        ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
        ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
        ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
        ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
      
        Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
        Set pf = pvtTbl.PivotFields("Period")

        pf.ClearAllFilters
        For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
            If pvtItm = endDate Then
                pvtItm.Visible = True
            Else
                pvtItm.Visible = False
                pvtItmName = pvtItm.Caption
            End If
        Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12
    
        Sheets("CR").Activate
        Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
        Set pf = pvtTbl.PivotFields("Period")
        pf.ClearAllFilters
        With ActiveSheet.PivotTables("pvtCR")
            .ColumnGrand = True
            .RowGrand = True
        End With
    
        For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
           If pvtItm = endDate Then
               pvtItm.Visible = True
           Else
               pvtItm.Visible = False
               pvtItmName = pvtItm.Caption
           End If
    
        Next pvtItm
    
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        mnthVal = ActiveCell.Offset(1, 0).value
        mnthValDvdByIBVal = mnthVal / IBVal
        On Error Resume Next

        startDate = Format(endDate, "m" & "/" & "d" & "/" & "yyyy")
        'Range("A2").Select
        'Range(Selection, Selection.End(xlToRight)).Select
        ActiveSheet.UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
        ActiveCell.Offset(1, 0).Select
        fstAdd = ActiveCell.Address
        Range(fstAdd).value = mnthValDvdByIBVal
        fstAdd = Range(fstAdd).Offset(0, 1).Address
        startDate = Format(DateAdd("m", 1, startDate), "yyyy" & "-" & "mm")
        endDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    
    Loop
    
'Enter 4th Row values of Current Year CR/Sys and Last Year for the ITM, IMQ, YTD and MAT
'ITM values
    ActiveSheet.UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.value = mnthValDvdByIBVal
    
'IMQ Values
    Sheets("IB").Select
    Range("N1").Select
    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")

    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"

    crntIMQ = Format(fixedDate, "yyyy" & "-" & "mm")
    previousMonth = Format(DateAdd("m", -1, crntIMQ), "yyyy" & "-" & "mm")
    qMnth = Format(DateAdd("m", -2, crntIMQ), "yyyy" & "-" & "mm")

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = crntIMQ Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
 
        End If
 
    Next
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate IMQ
    Sheets("CR").Select
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = crntIMQ Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
 
    Next
    
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdIMQ = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 1).Select
    ActiveCell.value = clcltdIMQ

'Calculate YTD
'IB for YTD
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, fixedDate), "yyyy" & "-" & "mm")
   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate YTD
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdYTD = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 2).Select
    ActiveCell.value = clcltdYTD

'Calculate MAT
'IB for MAT
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate MAT
Sheets("CR").Select
Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
pvtTbl.PivotFields("Period").ClearAllFilters

   For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdMAT = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Current Year CR/Sys", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 3).Select
    ActiveCell.value = clcltdMAT

'Calculate ITM for Last Year
'IB for ITM
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    PrvsITMDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm = PrvsITMDate Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate ITM
    Sheets("CR").Select
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm = PrvsITMDate Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm

    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdPrvsITM = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = clcltdPrvsITM

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'IMQ Values for Last Year
    Sheets("IB").Select
    Range("N1").Select
    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")

    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    PrvsIMQDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")

    previousMonth = Format(DateAdd("m", -1, PrvsIMQDate), "yyyy" & "-" & "mm")
    qMnth = Format(DateAdd("m", -2, PrvsIMQDate), "yyyy" & "-" & "mm")

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = PrvsIMQDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
 
        End If
 
    Next
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate IMQ
    Sheets("CR").Select
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm.value = PrvsIMQDate Or pvtItm.value = previousMonth Or pvtItm.value = qMnth Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
 
    Next
    
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdIMQ = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 1).Select
    ActiveCell.value = clcltdIMQ

'Calculate YTD
'IB for YTD
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    PrvsYTDDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")

    EndDateMonth = Mid(fixedDate, 6, 2)
    endDate = Format(DateAdd("m", -EndDateMonth, PrvsYTDDate), "yyyy" & "-" & "mm")
   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > PrvsYTDDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate YTD
    Sheets("CR").Select
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= endDate Or pvtItm > PrvsYTDDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdPrvsYTD = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 2).Select
    ActiveCell.value = clcltdPrvsYTD

'Calculate MAT
'IB for MAT
    Sheets("IB").Select
    Range("N1").Select
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    PrvsMATDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    PrvsYrMATDate = Format(DateAdd("m", -12, PrvsMATDate), "yyyy" & "-" & "mm")

    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= PrvsYrMATDate Or pvtItm > PrvsMATDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12

'Calculate MAT
    Sheets("CR").Select
    Set pvtTbl = Worksheets("CR").PivotTables("pvtCR")
    pvtTbl.PivotFields("Period").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm <= PrvsYrMATDate Or pvtItm > PrvsMATDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm

    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    clcltdPrvsMAT = ActiveCell.value / IBVal
    ActiveSheet.UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 3).Select
    ActiveCell.value = clcltdPrvsMAT

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Sheets("CR").Select
    ActiveSheet.UsedRange.Find(what:="CR/Sys/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).value = ActiveCell.Offset(1, 4).value
    pvtTbl.TableRange2.Select
    pvtTbl.TableRange2.Clear
    
    Range("A1").Select
    Application.ActiveWorkbook.Save
    Workbooks(myPvtWorkBook).Activate
    Application.Workbooks(myPvtWorkBook).Close
End Sub


Public Sub MTTRRateCalculationNew()
Dim fixedDate, mypath, CTSExcel, CTSWorkBook, pvtExcel, myPvtWorkBook As String
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
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
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Period1"
    Workbooks("CTS_KPI_Summary.xlsx").Activate
    myworkbook = ActiveWorkbook.name
    Sheets("KPI-All").Select
    
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
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    
    Sheets("MTTR").Select
    Range("A:A").Select
    On Error Resume Next
    Selection.EntireRow.Select
    Selection.EntireRow.Delete
    Application.Columns.Ungroup
    rows("1:1").Select
        
    Sheets("KPI-All").Select
    ActiveSheet.PivotTables("pvtKPIALL").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    False
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    pvtTbl.TableRange1.Select
    pvtTbl.TableRange1.Copy
    Sheets("MTTR").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Sheets("KPI-All").Select
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    True
    ActiveSheet.PivotTables("pvtKPIALL").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    Sheets("MTTR").Select
    Range("1:1").Select
    Selection.EntireRow.Delete
    Sheets("MTTR").UsedRange.Find(what:="#Avg. ETTR (days)/12", lookat:=xlWhole).Select
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
   
    'Sheets("Designed Data").Select
    'Range("A1").Select
    'Sheets("Designed Data").UsedRange.Find(what:="MTTR/ Sys / Yr", lookat:=xlWhole).Select

    'Selection.EntireColumn.Select
    'Selection.Copy
    
    'Sheets("MTTR").Select
    'Range("C1").Select
    'ActiveSheet.Paste
    'Application.CutCopyMode = False
    Sheets("MTTR").Activate
    Range("C1").Select
   ' Sheets("MTTR").UsedRange.Find(what:="Designed", lookat:=xlWhole).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(2, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Cells(2, 3).value = "SS&BB"

'Enter Designed values from Designed Data sheet into CR sheet

    Sheets("Designed Data").Select
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(2, -1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(0, 4).Select
    ActiveCell.Offset(2, 0).Select
    CRValFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -5).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 5).Select
    CRValLstAdd = ActiveCell.Address
    Sheets("MTTR").Select
    Range("D1").value = "Designed"
    Range("D2").value = "MTTR / Sys / Yr"
    Sheets("MTTR").UsedRange.Find(what:="MTTR / Sys / Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    DVVlkpResultAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    DVVlkpFstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
'    ActiveCell.Offset(0, 1).Select
    DVVlkpLstAdd = ActiveCell.Address
    fstAddRef = Range(DVVlkpFstAdd).Address(False, False)

    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'Designed Data'!R4C19:R98C22,4,0)"
    Range(DVVlkpResultAdd).Formula = "=IFERROR(VLOOKUP(" & fstAddRef & "," & "'Designed Data'!" & fstAdd & ":" & CRValLstAdd & ",6,0)" & "," & 0 & ")"
    Range(DVVlkpResultAdd).Select
    
    Selection.AutoFill Destination:=Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address)
    Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address).Select
    Calculate
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.NumberFormat = "0.00"
    
    Sheets("MTTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    nonPartsFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    nonPartslstAdd = ActiveCell.Address
    Range(nonPartsFstAdd, nonPartslstAdd).Select
    Selection.NumberFormat = "0.00"
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
        .NumberFormat = "0.00"

    End With
    
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    Sheets("MTTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.value = "MAT # of MTTR profiles"
    Range("E1:G1").Select
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

    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = "Non-Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Avg. MTTR/Yr"
    
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
    fixDte = Format(fixedDate1, "mmm" & "-" & "yyyy")
    fixDate2 = Format(DateAdd("yyyy", -1, fixedDate1), "mmm" & "-" & "yyyy")
    frmtData = Format(DateAdd("m", 1, fixedDate1), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
    endDate2 = Format(DateAdd("m", -24, frmtData), "mmm" & "-" & "yyyy")

    fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    fnlEndDate1 = Format(endDate2, "mmm" & "-" & "yyyy")
    frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
    Range("L1").Select
    ActiveCell.value = "Last Year"
    Range("H1").Select
    ActiveCell.value = "Avg. MTTR / Call (Current Year)"
    Range("Q2").Select
    Do Until frmEndDate = frmtData
        ActiveCell.value = frmEndDate
        ActiveCell.Offset(0, 1).Select
        frmEndDate = Format(DateAdd("m", 1, frmEndDate), "mmm" & "-" & "yyyy")
    Loop

    Range("A1").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(-1, 0).Select
    upAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

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
    Range(up1Add).Select
    ActiveCell.value = "Avg of Avg MTTR/Month"
    Range(up1Add, upAdd).Select
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
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    
    ActiveSheet.UsedRange.Select
    Selection.RowHeight = 15
    Range("H1:P2").Select
    Selection.Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlRight
    End With
    
    Cells(2, 1).Select
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 7
    Call MTTRPivotTableNew
    Dim visPvtItm As String
    Set pvtTbl = Worksheets("MTTR").PivotTables("pvtMTTR")
    fixedDate = Sheet1.combYear.value

'enter values in 4th Row for All sub systems and All Buildingblocks

    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("System")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("Period")
        .Orientation = xlColumnField
        .Position = 2
    End With
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("Part12NC").Orientation = _
        xlHidden
   
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System"). _
        EnableMultiplePageItems = False
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").CurrentPage = "(All)"
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").CurrentPage = _
        "System level"

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    With ActiveSheet.PivotTables("pvtMTTR")
        .ColumnGrand = True
        .RowGrand = False
    End With
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    startDate = Format(endDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    fstAdd = ActiveCell.Offset(1, 0).Address
    Sheets("MTTR").UsedRange.Find(what:="System Level", lookat:=xlWhole).Select
    ActiveCell.Offset(4, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(fstAdd).PasteSpecial
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").Orientation = xlHidden
    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("Part12NC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtMTTR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    
'Enter 12 months Data in the column "Avg of Avg MTTR/Month" after Crossover Trigger
    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtMTTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 18).Select
    i = 18
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("MTTR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(1, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            ActiveCell.Offset(0, 5).Select
            lstAdd = ActiveCell.Address(False, False)
            rng = Range(fstAdd, lstAdd)
            'rng = Range("AE3:AJ91")
            
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
                    Cells(X, myCol).NumberFormat = "0.00"
                Next X
             
            End If
                i = i + 1
        End If
    Next pvtItm

    Range("C3").Select
    ActiveCell.Offset(1, 3).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    
'==================================================================================

'Enter 12 months data for the last year from the selected date for Last year YTD and MTD calculations

    fixedDate1 = Sheet1.combYear.value
    frmtfixedDate1 = Format(fixedDate1, "mmm" & "-" & "yyyy")
    prvsYrDate1 = Format(DateAdd("yyyy", -1, frmtfixedDate1), "mmm" & "-" & "yyyy")
    prvsYrDate = Format(DateAdd("m", 1, prvsYrDate1), "mmm" & "-" & "yyyy")

    prvs2YrDate = Format(DateAdd("yyyy", -1, prvsYrDate), "mmm" & "-" & "yyyy")

    Range("AQ2").Select
    Do Until prvs2YrDate = prvsYrDate
        ActiveCell.value = prvs2YrDate
        ActiveCell.Offset(0, 1).Select
        prvs2YrDate = Format(DateAdd("m", 1, prvs2YrDate), "mmm" & "-" & "yyyy")
    Loop
    
'Enter 12 months Data for the last year to calculate last years ITM, IMQ, YTD and MAT
    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    prvsYrDate1 = Format(DateAdd("yyyy", -1, startDate), "mmm" & "-" & "yyyy")
    startDate = Format(prvsYrDate1, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtMTTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 43).Select
    i = 18
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("MTTR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(1, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("MTTR").Cells(rows.Count, "C").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            ActiveCell.Offset(0, 6).Select
            lstAdd = ActiveCell.Address(False, False)
            rng = Range(fstAdd, lstAdd)
            'rng = Range("AE3:AJ91")
            
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
                    Cells(X, myCol).NumberFormat = "0.00"
                Next X
             
            End If
                i = i + 1
        End If
    Next pvtItm

    Range("C3").Select
    ActiveCell.Offset(1, 3).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("" & "AQ" & sumMidAdd - 1 & ":" & "BB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    
'Enter ITM values for current year month
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ITMDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    endDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:=endDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    ITMMonth = ActiveCell.Address
    Do While ActiveCell.value <> ""
        Range(ITMDatacellAdd).value = ActiveCell.value
        ITMDatacellAdd = Range(ITMDatacellAdd).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop

'Enter IMQ values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="IMQ", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

        Range(IMQDatacellAdd).value = "=IFERROR(AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")," & "NA" & ")"
        IMQDatacellAdd = Range(IMQDatacellAdd).Offset(1, 0).Address
        IMQStrtMonth = Range(IMQStrtMonth).Offset(1, 0).Address
        IMQEndMonth = Range(IMQEndMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop
    
 'Enter YTD values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="YTD", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

        Range(YTDDatacellAdd).value = "=IFERROR(AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")," & "NA" & ")"
        YTDDatacellAdd = Range(YTDDatacellAdd).Offset(1, 0).Address
        YDTStrtMonth = Range(YDTStrtMonth).Offset(1, 0).Address
        YDTEndMonth = Range(YDTEndMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop

'Enter MAT values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="MAT", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

        Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"
        MATDatacellAdd = Range(MATDatacellAdd).Offset(1, 0).Address
        MATStrtMonth = Range(MATStrtMonth).Offset(1, 0).Address
        MATEndMonth = Range(MATEndMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop
    
'=============================================================================================

'==================================================================================================

'Enter ITM values for Last year month
    fixedDate = Sheet1.combYear.value
    findDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(findDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    
    ActiveCell.Offset(3, 0).Select
    ITMLstYrAdd = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:=endDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    ITMLstYrMonth = ActiveCell.Address
    Do While ActiveCell.value <> ""
        Range(ITMLstYrAdd).value = ActiveCell.value
        ITMLstYrAdd = Range(ITMLstYrAdd).Offset(1, 0).Address
        ITMLstYrMonth = Range(ITMLstYrMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop

'Enter IMQ values for Last year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    IMDDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    startDate = Format(IMDDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 1).Select
    IMQDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

        Range(IMQDatacellAdd).value = "=IFERROR(AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")," & "NA" & ")"
        IMQDatacellAdd = Range(IMQDatacellAdd).Offset(1, 0).Address
        IMQStrtMonth = Range(IMQStrtMonth).Offset(1, 0).Address
        IMQEndMonth = Range(IMQEndMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop
    
 'Enter YTD values for Last year
    fixedDate = Sheet1.combYear.value
    YTDDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate = Format(YTDDate, "yyyy" & "-" & "mm")
    startDate = Format(YTDDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(YTDDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 2).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

        Range(YTDDatacellAdd).value = "=IFERROR(AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")," & "NA" & ")"
        YTDDatacellAdd = Range(YTDDatacellAdd).Offset(1, 0).Address
        YDTStrtMonth = Range(YDTStrtMonth).Offset(1, 0).Address
        YDTEndMonth = Range(YDTEndMonth).Offset(1, 0).Address
        ActiveCell.Offset(1, 0).Select
    Loop

'Enter MAT values for Last year
    fixedDate = Sheet1.combYear.value
    MATDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "yyyy" & "-" & "mm")
    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 3).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""

         Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"
         MATDatacellAdd = Range(MATDatacellAdd).Offset(1, 0).Address
         MATStrtMonth = Range(MATStrtMonth).Offset(1, 0).Address
         MATEndMonth = Range(MATEndMonth).Offset(1, 0).Address
         ActiveCell.Offset(1, 0).Select
    Loop

'Enter values of Avg. MTTR/Yr
    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtMTTR")
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
    Range("AE3").Select
    fstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 5).Select
    lstAdd = ActiveCell.Address(False, False)
    Range(fstAdd, lstAdd).Select
    rng = Range(fstAdd, lstAdd)
            
    For X = 3 To lr
        On Error Resume Next
        Cells(X, 7).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
        Cells(X, 7).NumberFormat = "0.00"
    Next X

'Enter values in 4th row for All Sub systems and All Building blocks

    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("System")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("System")
        .PivotItems("").Visible = False
        .PivotItems(" ").Visible = False
    End With
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("System"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("pvtMTTR").ColumnGrand = False
    
    Sheets("MTTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    MTTRperYrVal = ActiveCell.Offset(1, 0).value
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = MTTRperYrVal
    
'==================================================================================================
    Range("C3").Select
    ActiveCell.Offset(1, 2).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
   
    Sheets("MTTR").Select
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=AND(RC[-7]>RC[-6],RC[-12]<RC[-9])"
    Selection.AutoFill Destination:=Range("" & "P" & sumMidAdd - 1 & ":" & "P" & sumMidAdd1 & "")
    Range("" & "P" & sumMidAdd - 1 & ":" & "P" & sumMidAdd1 & "").Select
    Calculate
    Range("" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.00"
    Application.CutCopyMode = False
    Range("AC1").Select
    Selection.EntireColumn.Select
     Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     Range("" & "AC" & sumMidAdd - 1 & ":" & "AC" & sumMidAdd1 & "").Select
    Selection.SparklineGroups.Add Type:=xlSparkLine, SourceData:= _
        "" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & ""
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
    Range("" & "G" & sumMidAdd - 1 & ":" & "O" & sumMidAdd1 & "").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.00"
    Application.CutCopyMode = False
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Application.CutCopyMode = False
    Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & ",10)"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    'Formula1:="=G4/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    

'Highlight MTTR Crossover Trigger in Red or Orange for true values depending upon 20% largest values
    Sheets("MTTR").Select
    Sheets("MTTR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 3).Select
    fstAdd = ActiveCell.Address(False, False)
    fstMidAdd = Mid(fstAdd, 2)
    
    Sheets("MTTR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    lstAdd = ActiveCell.Address(False, False)
    lstMidAdd = Mid(lstAdd, 2)
    ActiveCell.Offset(1, 0).Select
   
    ActiveCell.Formula = "=LARGE(" & "E" & fstMidAdd & ":" & "F" & lstMidAdd & ",20)"
    topTwentyVal = ActiveCell.value
    
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    Do While ActiveCell.value <> ""
    If ActiveCell.value = "True" Then
        If Range("E" & fstMidAdd).value >= topTwentyVal Or Range("F" & fstMidAdd).value >= topTwentyVal Then
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        Else
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 1484526
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        End If
        
    End If
    fstMidAdd = fstMidAdd + 1
    ActiveCell.Offset(1, 0).Select
    
    Loop
    Range(lstAdd).Offset(1, 0).Select
    Selection.ClearContents

    Range("C1").Select
    Selection.EntireColumn.Delete
    Range("AE10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    
    Sheets("MTTR").UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    ITMfrtRowVal = ActiveCell.value
    ITMfrtRowAdd = ActiveCell.Address
    IMQfrtRowAdd = ActiveCell.Offset(0, -2).Address

    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 1).value = ITMfrtRowVal
    ActiveCell.Offset(1, 2) = "=IFERROR(AVERAGE(" & ITMfrtRowAdd & ":" & IMQfrtRowAdd & ")," & "NA" & ")"

'Enter YTD value in 4th Row
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    Sheets("MTTR").UsedRange.Find(what:="YTD", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    'Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"

    Range(YTDDatacellAdd).value = "=IFERROR(AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")," & "NA" & ")"

'Enter MAT value in 4th Row
    fixedDate = Sheet1.combYear.value
    MATDate = Format(fixedDate, "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "yyyy" & "-" & "mm")
    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR / Call (Current Year)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 3).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"
   
'Enter values in 4th row for last years, ITM,IMQ, YTD and MAT
    
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("Part12NC").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("pvtMTTR").PivotFields("Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtMTTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    With ActiveSheet.PivotTables("pvtMTTR")
        .ColumnGrand = True
        .RowGrand = False
    End With
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    
    fixedDate1 = Sheet1.combYear.value
    fixDte = Format(fixedDate1, "mmm" & "-" & "yyyy")
    fixDate2 = Format(DateAdd("yyyy", -1, fixedDate1), "mmm" & "-" & "yyyy")
    frmtData = Format(DateAdd("m", 1, fixDate2), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
    endDate2 = Format(DateAdd("m", -24, frmtData), "mmm" & "-" & "yyyy")

    fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    fnlEndDate1 = Format(endDate2, "mmm" & "-" & "yyyy")
    frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
   
    Range("AE10").Select
    fstadd1 = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Buildingblocks Aggregated", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Copy
    Range(fstadd1).PasteSpecial xlPasteAll
    Do While ActiveCell.value <> ""
    fixedDate = ActiveCell.value
    frmtDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    startDate = Format(frmtDate, "mmm" & "-" & "yyyy")
    ActiveCell.value = startDate
    ActiveCell.Offset(0, 1).Select
    
    Loop
    pvtTbl.TableRange2.Select
    pvtTbl.TableRange2.Clear
    ActiveCell.End(xlDown).Select
    fstAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Offset(1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Copy
    Range("AE2").PasteSpecial xlPasteAll
    Range(fstAdd, lstAdd).ClearContents
    
'Enter ITM values for current year month
    fixedDate = Sheet1.combYear.value
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ITMDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "mmm" & "-" & "yyyy")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    endDate = Format(DateAdd("yyyy", -2, fixedDate), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("m", 1, endDate), "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    ITMMonth = ActiveCell.Address
    Range(ITMDatacellAdd).value = Range(ITMMonth).value

'Enter IMQ values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(startDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    IMQDatacellAdd = Range(ITMDatacellAdd).Offset(0, 1).Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Sheets("MTTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Range("A2").Select
   ' Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"

    Range(IMQDatacellAdd).value = "=IFERROR(AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")," & "" & "NA" & "" & ")"
 
 'Enter YTD value in 4th Row
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(startDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("MTTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    YTDDatacellAdd = Range(ITMDatacellAdd).Offset(0, 2).Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("MTTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Range(YTDDatacellAdd).value = "=IFERROR(AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")," & "" & "NA" & "" & ")"

'Enter MAT value in 4th Row
    fixedDate = Sheet1.combYear.value
    MATDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    MATDatacellAdd = Range(ITMDatacellAdd).Offset(0, 3).Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("MTTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    If Not Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole) Is Nothing Then
    ActiveCell.Offset(1, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"
    
    Else
    Range(MATDatacellAdd).value = "NA"
    
    End If
    Sheets("MTTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Range(MATDatacellAdd).value = "=IFERROR(AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")," & "NA" & ")"
    Sheets("MTTR").UsedRange.Find(what:="Avg. MTTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues

    Range(MATEndMonth).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    Range("A1").Select
    'ActiveWindow.Zoom = 85
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "True" Then
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
    End If
    Cells(1, 1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    With Selection.Font
        .name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.EntireColumn.Select
    
    Sheets("MTTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 10.57
'Add Heading to DashBoard
    Sheets("MTTR").Select
    rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "MTTR Dash Board for " & KPISheetName
    Range("A1:AB1").Select
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
        .Font.Bold = True
        .Font.Italic = True
        .Font.name = "Calibri"
        .Font.Size = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.399975585192419
        .Interior.PatternTintAndShade = 0
    End With
    Selection.Merge
    Selection.Font.Bold = True
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
     ActiveWindow.Zoom = 100
     rows("1:1").Select
    Selection.RowHeight = 25
    Range("A2").Select
    Workbooks(myPvtWorkBook).Close
    Application.Workbooks(myworkbook).Save
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
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
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Sheets(AggrDataShtName).Activate
    Cells(1, 1).Select
    fstCellAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter
        End If
     
'Add a new sheet to create a Pivot Table
    Sheets.Add after:=Worksheets(Worksheets.Count)
    Set wsPtTable = Worksheets(Sheets.Count)
    'Set wsPtTable = Worksheets(3)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets(AggrDataShtName).Activate
    Set wsData = Worksheets(AggrDataShtName)
    Worksheets(AggrDataShtName).Activate
    sourceSheet = ActiveSheet.name
    ActiveSheet.Cells(1, 1).Select
    fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    Sheets(wsptName).Activate
    rngData = fstAdd & ":" & lstAdd
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
    "pvtMTTR").PivotFields("Avg. MTTR/Call (hrs)"), "#MTTR/Call (hrs)", xlAverage
        
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
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("SubSystem").RepeatLabels = _
    True
    ActiveSheet.PivotTables("pvtMTTR").PivotFields("BuildingBlock").RepeatLabels _
    = True
    pvtTbl.RefreshTable
    
    Columns("A:E").EntireColumn.AutoFit
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
    ActiveSheet.PivotTables("pvtMTTR").Location = _
        "'[CTS_KPI_Summary.xlsx]MTTR'!$AK$3"
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("MTTR").Activate
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[5]"
    Range("AF3").Select
    Selection.Copy
    Range("AF3,AF91").Select
    Range("AF3,AF3:AJ91").Select
    ActiveSheet.Paste
    Range("AE2").Select
    ActiveCell.value = "SS&BB"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],RC[2])"
    Range("AE3").Select
    Selection.AutoFill Destination:=Range("AE3:AE91")
    
End Function
Public Sub ETTRRateCalculationNew()
Dim fixedDate, mypath, CTSExcel, CTSWorkBook, pvtExcel, myPvtWorkBook As String
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
    prdFileName = KPISheetName

'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
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
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
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
    Workbooks("CTS_KPI_Summary.xlsx").Activate
    myworkbook = ActiveWorkbook.name
    Sheets("KPI-All").Select
    
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
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    
    Sheets("ETTR").Select
    Range("A:A").Select
    On Error Resume Next
    Selection.EntireRow.Select
    Selection.EntireRow.Delete
    Application.Columns.Ungroup
    rows("1:1").Select
        
    Sheets("KPI-All").Select
    ActiveSheet.PivotTables("pvtKPIALL").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    False
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    pvtTbl.TableRange1.Select
    pvtTbl.TableRange1.Copy
    Sheets("ETTR").Select
    Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Sheets("KPI-All").Select
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("BuildingBlock").ShowDetail = _
    True
    ActiveSheet.PivotTables("pvtKPIALL").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIALL").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    Sheets("ETTR").Select
    Range("1:1").Select
    Selection.EntireRow.Delete
    Sheets("ETTR").UsedRange.Find(what:="Visits/call (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstclmn = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lstclmnAdd = ActiveCell.Address
    Range(fstclmn, lstclmnAdd).Select
    Selection.EntireColumn.Select
    Selection.EntireColumn.Delete
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Sheets("ETTR").UsedRange.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
    deleteClmnsAdd = ActiveCell.Address
    Sheets("ETTR").UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
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
    'Sheets("Designed Data").Select
    'Range("A1").Select
    'Sheets("Designed Data").UsedRange.Find(what:="ETTR / Sys / Yr", lookat:=xlWhole).Select

    'Selection.EntireColumn.Select
    'Selection.Copy
    
    'Sheets("ETTR").Select
    'Range("C1").Select
    'ActiveSheet.Paste
    'Application.CutCopyMode = False
    'Sheets("ETTR").Activate
     
    Range("C2").Select
    Range("C2").value = "SS&BB"
    ActiveCell.Offset(1, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("ETTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 2).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    
    'Enter Designed values from Designed Data sheet into CR sheet

    Range("D1").Select
    Selection.EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").value = "Designed"
    Range("D2").value = "ETTR / Sys / Yr"

    Sheets("Designed Data").Select
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(2, -1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=RC[1]&RC[2]"
    Selection.AutoFill Destination:=ActiveSheet.Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Sheets("Designed Data").UsedRange.Find(what:=KPISheetName, lookat:=xlWhole).Select
    ActiveCell.Offset(0, 3).Select
    ActiveCell.Offset(2, 0).Select
    CRValFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -4).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    CRValLstAdd = ActiveCell.Address
    Sheets("ETTR").Select
    Sheets("ETTR").UsedRange.Find(what:="ETTR / Sys / Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    DVVlkpResultAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    DVVlkpFstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
'    ActiveCell.Offset(0, 1).Select
    DVVlkpLstAdd = ActiveCell.Address
    fstAddRef = Range(DVVlkpFstAdd).Address(False, False)

    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'Designed Data'!R4C19:R98C22,4,0)"
    Range(DVVlkpResultAdd).Formula = "=IFERROR(VLOOKUP(" & fstAddRef & "," & "'Designed Data'!" & fstAdd & ":" & CRValLstAdd & ",5,0)" & "," & 0 & ")"
    Range(DVVlkpResultAdd).Select
    
    Selection.AutoFill Destination:=Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address)
    Range(DVVlkpResultAdd, Range(DVVlkpLstAdd).Offset(0, 1).Address).Select
    Calculate
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.NumberFormat = "0.00"
    
    Sheets("ETTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    nonPartsFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    nonPartslstAdd = ActiveCell.Address
    Range(nonPartsFstAdd, nonPartslstAdd).Select
    Selection.NumberFormat = "0.00"
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
    
    Sheets("ETTR").UsedRange.Find(what:="Parts", lookat:=xlWhole).Select
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
        .NumberFormat = "0.00"

    End With
    
    Sheets("ETTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.EntireRow.Delete
    Sheets("ETTR").UsedRange.Find(what:="Non-Parts", lookat:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.value = "MAT # of ETTR profiles"
    Range("E1:G1").Select
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

    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = "Non-Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Parts"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "Avg. ETTR/Yr"
    
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
    fixDte = Format(fixedDate1, "mmm" & "-" & "yyyy")
    fixDate2 = Format(DateAdd("yyyy", -1, fixedDate1), "mmm" & "-" & "yyyy")
    frmtData = Format(DateAdd("m", 1, fixedDate1), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
    endDate2 = Format(DateAdd("m", -24, frmtData), "mmm" & "-" & "yyyy")

    fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    fnlEndDate1 = Format(endDate2, "mmm" & "-" & "yyyy")
    frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
    Range("L1").Select
    ActiveCell.value = "Last Year"
    Range("H1").Select
    ActiveCell.value = "Avg. ETTR (days) (Current Year)"
    Range("Q2").Select
    Do Until frmEndDate = frmtData
        ActiveCell.value = frmEndDate
        ActiveCell.Offset(0, 1).Select
        frmEndDate = Format(DateAdd("m", 1, frmEndDate), "mmm" & "-" & "yyyy")
    Loop

    Range("A1").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(-1, 0).Select
    upAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15652757
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("ETTR").UsedRange.Find(what:="Crossover", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 1).Select
    up1Add = ActiveCell.Address
    Range(up1Add).Select
    ActiveCell.value = "Avg of Avg ETTR/Month"
    Range(up1Add, upAdd).Select
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
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    
    ActiveSheet.UsedRange.Select
    Selection.RowHeight = 15
    Range("H1:P2").Select
    Selection.Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlRight
    End With
    
    Cells(2, 1).Select
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 7
    Call ETTRPivotTableNew
    Dim visPvtItm As String
    Set pvtTbl = Worksheets("ETTR").PivotTables("pvtETTR")
    fixedDate = Sheet1.combYear.value

'enter values in 4th Row for All sub systems and All Buildingblocks

    With ActiveSheet.PivotTables("pvtETTR").PivotFields("System")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("Period")
        .Orientation = xlColumnField
        .Position = 2
    End With
    ActiveSheet.PivotTables("pvtETTR").PivotFields("Part12NC").Orientation = _
        xlHidden
   
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System"). _
        EnableMultiplePageItems = False
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").CurrentPage = "(All)"
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").CurrentPage = _
        "System level"

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    With ActiveSheet.PivotTables("pvtETTR")
        .ColumnGrand = True
        .RowGrand = False
    End With
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    startDate = Format(endDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    fstAdd = ActiveCell.Offset(1, 0).Address
    Sheets("ETTR").UsedRange.Find(what:="System Level", lookat:=xlWhole).Select
    ActiveCell.Offset(4, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(fstAdd).PasteSpecial
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").Orientation = xlHidden
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("Part12NC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("pvtETTR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    
'Enter 12 months Data in the column "Avg of Avg ETTR/Month" after Crossover Trigger
    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtETTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 18).Select
    i = 18
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("ETTR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(1, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("ETTR").Cells(rows.Count, "C").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            ActiveCell.Offset(0, 5).Select
            lstAdd = ActiveCell.Address(False, False)
            rng = Range(fstAdd, lstAdd)
            'rng = Range("AE3:AJ91")
            
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
                    Cells(X, myCol).NumberFormat = "0.00"
                Next X
             
            End If
                i = i + 1
        End If
    Next pvtItm

    Range("C3").Select
    ActiveCell.Offset(1, 3).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    
'==================================================================================

'Enter 12 months data for the last year from the selected date for Last year YTD and MTD calculations

    fixedDate1 = Sheet1.combYear.value
    frmtfixedDate1 = Format(fixedDate1, "mmm" & "-" & "yyyy")
    prvsYrDate1 = Format(DateAdd("yyyy", -1, frmtfixedDate1), "mmm" & "-" & "yyyy")
    prvsYrDate = Format(DateAdd("m", 1, prvsYrDate1), "mmm" & "-" & "yyyy")

    prvs2YrDate = Format(DateAdd("yyyy", -1, prvsYrDate), "mmm" & "-" & "yyyy")

    Range("AQ2").Select
    Do Until prvs2YrDate = prvsYrDate
        ActiveCell.value = prvs2YrDate
        ActiveCell.Offset(0, 1).Select
        prvs2YrDate = Format(DateAdd("m", 1, prvs2YrDate), "mmm" & "-" & "yyyy")
    Loop
    
'Enter 12 months Data for the last year to calculate last years ITM, IMQ, YTD and MAT
    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    prvsYrDate1 = Format(DateAdd("yyyy", -1, startDate), "mmm" & "-" & "yyyy")
    startDate = Format(prvsYrDate1, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtETTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 43).Select
    i = 18
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("ETTR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(1, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            lr = Worksheets("ETTR").Cells(rows.Count, "C").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            ActiveCell.Offset(0, 6).Select
            lstAdd = ActiveCell.Address(False, False)
            rng = Range(fstAdd, lstAdd)
            'rng = Range("AE3:AJ91")
            
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
                    Cells(X, myCol).NumberFormat = "0.00"
                Next X
             
            End If
                i = i + 1
        End If
    Next pvtItm

    Range("C3").Select
    ActiveCell.Offset(1, 3).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -3).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 3).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
    Range("" & "AQ" & sumMidAdd - 1 & ":" & "BB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.0000"
    
'Enter ITM values for current year month
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="ITM", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ITMDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    endDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:=endDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    ITMMonth = ActiveCell.Address
    Do While ActiveCell.value <> ""
    Range(ITMDatacellAdd).value = ActiveCell.value
    ITMDatacellAdd = Range(ITMDatacellAdd).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop

'Enter IMQ values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="IMQ", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""
    Range(IMQDatacellAdd).value = "=AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")"

    IMQDatacellAdd = Range(IMQDatacellAdd).Offset(1, 0).Address
    IMQStrtMonth = Range(IMQStrtMonth).Offset(1, 0).Address
    IMQEndMonth = Range(IMQEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop
    
 'Enter YTD values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="YTD", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""
    Range(YTDDatacellAdd).value = "=AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")"

    YTDDatacellAdd = Range(YTDDatacellAdd).Offset(1, 0).Address
    YDTStrtMonth = Range(YDTStrtMonth).Offset(1, 0).Address
    YDTEndMonth = Range(YDTEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop

'Enter MAT values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="MAT", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
   ' Range(YTDDatacellAdd).Select
    Do While ActiveCell.value <> ""
    Range(MATDatacellAdd).value = "=AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")"

    MATDatacellAdd = Range(MATDatacellAdd).Offset(1, 0).Address
    MATStrtMonth = Range(MATStrtMonth).Offset(1, 0).Address
    MATEndMonth = Range(MATEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop
    
'=============================================================================================

'==================================================================================================

'Enter ITM values for Last year month
    fixedDate = Sheet1.combYear.value
    findDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(findDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    
    ActiveCell.Offset(3, 0).Select
    ITMLstYrAdd = ActiveCell.Address
    Sheets("ETTR").UsedRange.Find(what:=endDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    ITMLstYrMonth = ActiveCell.Address
    Do While ActiveCell.value <> ""
    Range(ITMLstYrAdd).value = ActiveCell.value
    ITMLstYrAdd = Range(ITMLstYrAdd).Offset(1, 0).Address
    ITMLstYrMonth = Range(ITMLstYrMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop

'Enter IMQ values for Last year
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    IMDDate = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    startDate = Format(IMDDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 1).Select
    IMQDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""
    Range(IMQDatacellAdd).value = "=AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")"

    IMQDatacellAdd = Range(IMQDatacellAdd).Offset(1, 0).Address
    IMQStrtMonth = Range(IMQStrtMonth).Offset(1, 0).Address
    IMQEndMonth = Range(IMQEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop
    
 'Enter YTD values for Last year
    fixedDate = Sheet1.combYear.value
    YTDDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate = Format(YTDDate, "yyyy" & "-" & "mm")
    startDate = Format(YTDDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(YTDDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 2).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""
    Range(YTDDatacellAdd).value = "=AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")"

    YTDDatacellAdd = Range(YTDDatacellAdd).Offset(1, 0).Address
    YDTStrtMonth = Range(YDTStrtMonth).Offset(1, 0).Address
    YDTEndMonth = Range(YDTEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop

'Enter MAT values for Last year
    fixedDate = Sheet1.combYear.value
    MATDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "yyyy" & "-" & "mm")
    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(3, 0).Select
    ActiveCell.Offset(0, 3).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(2, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Do While ActiveCell.value <> ""
    Range(MATDatacellAdd).value = "=AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")"

    MATDatacellAdd = Range(MATDatacellAdd).Offset(1, 0).Address
    MATStrtMonth = Range(MATStrtMonth).Offset(1, 0).Address
    MATEndMonth = Range(MATEndMonth).Offset(1, 0).Address
    ActiveCell.Offset(1, 0).Select
    Loop

'Enter values of Avg. ETTR/Yr
    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtETTR")
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
    Range("AE3").Select
    fstAdd = ActiveCell.Address(False, False)
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 5).Select
    lstAdd = ActiveCell.Address(False, False)
    Range(fstAdd, lstAdd).Select
    rng = Range(fstAdd, lstAdd)
            
    For X = 3 To lr
        On Error Resume Next
        Cells(X, 7).value = Application.WorksheetFunction.VLookup(Cells(X, 3).value, rng, 6, False)
        Cells(X, 7).NumberFormat = "0.00"
    Next X

'Enter values in 4th row for All Sub systems and All Building blocks

    With ActiveSheet.PivotTables("pvtETTR").PivotFields("System")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("System")
        .PivotItems("").Visible = False
        .PivotItems(" ").Visible = False
    End With
    ActiveSheet.PivotTables("pvtETTR").PivotFields("System"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("pvtETTR").ColumnGrand = False
    
    Sheets("ETTR").UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ETTRperYrVal = ActiveCell.Offset(1, 0).value
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = ETTRperYrVal
    
'==================================================================================================
    Range("C3").Select
    ActiveCell.Offset(1, 2).Select
    sumAdd = ActiveCell.Address(False, False)
    sumMidAdd = Mid(sumAdd, 2)
    ActiveCell.Offset(0, -2).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    sumAdd1 = ActiveCell.Address(False, False)
    sumMidAdd1 = Mid(sumAdd1, 2)
    Range("F3").Select
    sumAdd2 = ActiveCell.Address(False, False)
    sumMidAdd2 = Mid(sumAdd2, 2)
   
    Sheets("ETTR").Select
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=AND(RC[-7]>RC[-6],RC[-12]<RC[-9])"
    Selection.AutoFill Destination:=Range("" & "P" & sumMidAdd - 1 & ":" & "P" & sumMidAdd1 & "")
    Range("" & "P" & sumMidAdd - 1 & ":" & "P" & sumMidAdd1 & "").Select
    Calculate
    Range("" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & "").Select
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.00"
    Application.CutCopyMode = False
    Range("AC1").Select
    Selection.EntireColumn.Select
     Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     Range("" & "AC" & sumMidAdd - 1 & ":" & "AC" & sumMidAdd1 & "").Select
    Selection.SparklineGroups.Add Type:=xlSparkLine, SourceData:= _
        "" & "Q" & sumMidAdd - 1 & ":" & "AB" & sumMidAdd1 & ""
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
    Range("" & "G" & sumMidAdd - 1 & ":" & "O" & sumMidAdd1 & "").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:="", Replacement:="0.00", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "0.00"
    Application.CutCopyMode = False
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Trend"
    Application.CutCopyMode = False
    Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=LARGE(" & fstAdd & ":" & lstAdd & ",10)"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    'Formula1:="=G4/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'Highlight ETTR Crossover Trigger in Red or Orange for true values depending upon 20% largest values
    Sheets("ETTR").Select
    Sheets("ETTR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 3).Select
    fstAdd = ActiveCell.Address(False, False)
    fstMidAdd = Mid(fstAdd, 2)
    
    Sheets("ETTR").UsedRange.Find(what:="BuildingBlock", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 4).Select
    lstAdd = ActiveCell.Address(False, False)
    lstMidAdd = Mid(lstAdd, 2)
    ActiveCell.Offset(1, 0).Select
   
    ActiveCell.Formula = "=LARGE(" & "E" & fstMidAdd & ":" & "F" & lstMidAdd & ",20)"
    topTwentyVal = ActiveCell.value
    
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    Do While ActiveCell.value <> ""
    If ActiveCell.value = "True" Then
        If Range("E" & fstMidAdd).value >= topTwentyVal Or Range("F" & fstMidAdd).value >= topTwentyVal Then
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        Else
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 1484526
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        End If
        
    End If
    fstMidAdd = fstMidAdd + 1
    ActiveCell.Offset(1, 0).Select
    
    Loop
    Range(lstAdd).Offset(1, 0).Select
    Selection.ClearContents

    Range("C1").Select
    Selection.EntireColumn.Delete
    Range("AE10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    
    Sheets("ETTR").UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    ITMfrtRowVal = ActiveCell.value
    ITMfrtRowAdd = ActiveCell.Address
    IMQfrtRowAdd = ActiveCell.Offset(0, -2).Address

    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 1).value = ITMfrtRowVal
    ActiveCell.Offset(1, 2) = "=AVERAGE(" & ITMfrtRowAdd & ":" & IMQfrtRowAdd & ")"

'Enter YTD value in 4th Row
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    startDate = Format(fixedDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(fixedDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    Sheets("ETTR").UsedRange.Find(what:="YTD", after:=ActiveCell, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    YTDDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Range(YTDDatacellAdd).value = "=AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")"

'Enter MAT value in 4th Row
    fixedDate = Sheet1.combYear.value
    MATDate = Format(fixedDate, "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "yyyy" & "-" & "mm")
    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR (days) (Current Year)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 3).Select
    MATDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Range(MATDatacellAdd).value = "=AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")"
   
'Enter values in 4th row for last years, ITM,IMQ, YTD and MAT
    
    ActiveSheet.PivotTables("pvtETTR").PivotFields("Part12NC").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("Period")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtETTR")
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    With ActiveSheet.PivotTables("pvtETTR")
        .ColumnGrand = True
        .RowGrand = False
    End With
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
    
    fixedDate1 = Sheet1.combYear.value
    fixDte = Format(fixedDate1, "mmm" & "-" & "yyyy")
    fixDate2 = Format(DateAdd("yyyy", -1, fixedDate1), "mmm" & "-" & "yyyy")
    frmtData = Format(DateAdd("m", 1, fixDate2), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("mmm", -12, frmtData), "mmm" & "-" & "yyyy")
    endDate2 = Format(DateAdd("m", -24, frmtData), "mmm" & "-" & "yyyy")

    fnlEndDate = Format(DateAdd("m", 1, endDate1), "mmm" & "-" & "yyyy")
    fnlEndDate1 = Format(endDate2, "mmm" & "-" & "yyyy")
    frmEndDate = Format(fnlEndDate, "mmm" & "-" & "yyyy")
   
    Range("AE9").Select
    Do Until frmEndDate = frmtData
    ActiveCell.value = frmEndDate
    ActiveCell.Offset(0, 1).Select
    frmEndDate = Format(DateAdd("m", 1, frmEndDate), "mmm" & "-" & "yyyy")
    Loop
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlToLeft).Select
    fstAdd = ActiveCell.Offset(1, 0).Address
    Sheets("ETTR").UsedRange.Find(what:="Buildingblocks Aggregated", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlToRight)).Copy
    Range(fstAdd).PasteSpecial xlPasteAll
    pvtTbl.TableRange2.Select
    pvtTbl.TableRange2.Clear
    ActiveCell.End(xlDown).Select
    fstAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Offset(1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Copy
    Range("AE2").PasteSpecial xlPasteAll
    Range(fstAdd, lstAdd).ClearContents
    
'Enter ITM values for current year month
    fixedDate = Sheet1.combYear.value
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ITMDatacellAdd = ActiveCell.Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "mmm" & "-" & "yyyy")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    endDate = Format(DateAdd("yyyy", -2, fixedDate), "mmm" & "-" & "yyyy")

    endDate1 = Format(DateAdd("m", 1, endDate), "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    ITMMonth = ActiveCell.Address
    Range(ITMDatacellAdd).value = Range(ITMMonth).value

'Enter IMQ values for current year
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(startDate, 6, 2)

    endDate = Format(DateAdd("m", -3, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
        IMQDatacellAdd = Range(ITMDatacellAdd).Offset(0, 1).Address
    Range("A2").Select
    
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    IMQStrtMonth = ActiveCell.Address(False, False)
    Sheets("ETTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    IMQEndMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Range(IMQDatacellAdd).value = "=AVERAGE(" & IMQStrtMonth & ":" & IMQEndMonth & ")"
 
 'Enter YTD value in 4th Row
    
    fixedDate = Sheet1.combYear.value
    startDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    startDate1 = Format(startDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(startDate, 6, 2)

    endDate = Format(DateAdd("m", -EndDateMonth, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    Sheets("ETTR").UsedRange.Find(what:="Last Year", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    YTDDatacellAdd = Range(ITMDatacellAdd).Offset(0, 2).Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("ETTR").UsedRange.Find(what:=startDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTStrtMonth = ActiveCell.Address(False, False)
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    YDTEndMonth = ActiveCell.Address(False, False)
    Range(YTDDatacellAdd).value = "=AVERAGE(" & YDTStrtMonth & ":" & YDTEndMonth & ")"

'Enter MAT value in 4th Row
    fixedDate = Sheet1.combYear.value
    MATDate = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")

    startDate = Format(MATDate, "m" & "/" & "d" & "/" & "yyyy")
    EndDateMonth = Mid(MATDate, 6, 2)

    endDate = Format(DateAdd("m", -12, startDate), "yyyy" & "-" & "mm")
    CendDate = Format(DateAdd("m", 1, endDate), "yyyy" & "-" & "mm")
    endDate1 = Format(CendDate, "m" & "/" & "d" & "/" & "yyyy")
    MATDatacellAdd = Range(ITMDatacellAdd).Offset(0, 3).Address
    Range("A2").Select
    
    Selection.EntireRow.Select
    Sheets("ETTR").UsedRange.Find(what:=startDate, lookat:=xlWhole).Select
    
    ActiveCell.Offset(1, 0).Select
    MATStrtMonth = ActiveCell.Address(False, False)
    
    Sheets("ETTR").UsedRange.Find(what:=endDate1, lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    MATEndMonth = ActiveCell.Address(False, False)
    Range(MATDatacellAdd).value = "=AVERAGE(" & MATStrtMonth & ":" & MATEndMonth & ")"
    Sheets("ETTR").UsedRange.Find(what:="Avg. ETTR/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    fstAdd = ActiveCell.Address
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    Range(MATEndMonth).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    Range("A1").Select
    'ActiveWindow.Zoom = 85
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.value = "True" Then
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
    End If
    Cells(1, 1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    With Selection.Font
        .name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.EntireColumn.Select
    
    Sheets("ETTR").UsedRange.Find(what:="Trigger", lookat:=xlWhole).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 10.57
'Add Heading to DashBoard
    Sheets("ETTR").Select
    rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ETTR Dash Board for " & KPISheetName
    Range("A1:AB1").Select
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
        .Font.Bold = True
        .Font.Italic = True
        .Font.name = "Calibri"
        .Font.Size = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.399975585192419
        .Interior.PatternTintAndShade = 0
    End With
    Selection.Merge
    Selection.Font.Bold = True
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
     rows("1:1").Select
    Selection.RowHeight = 25
    Range("A2").Select
    Application.Workbooks(myPvtWorkBook).Close
    Sheets("KPI-All").Select
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
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > fixedDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
'Create a new sheet for the summary of the All subsystems and Building blocks
    Call allSSsummarySheet
   ' Call RRRDatainAllSSBB
'Save an output file with the productgroup name
    Application.Workbooks(myworkbook).Activate
    Sheets("ETTR").Select
    ActiveWindow.Zoom = 100
    Sheets("KPI-Master").Select
    ActiveWorkbook.SaveAs fileName:= _
        ThisWorkbook.Path & "\CTS_KPI_Summary_" & prdFileName & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

End Sub
Public Function ETTRPivotTableNew()
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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
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
    'prdNameFile = KPISheetName & "_" & dateValue
    prdFileName = KPISheetName

'Open Aggregated Data File
    'Open input file-Aggregated Data File

    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        myPvtWorkBook = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
    myPvtWorkBook = ActiveWorkbook.name
    
    Workbooks(myPvtWorkBook).Activate
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.EntireRow.Select
    Selection.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Range("A2").Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstAdd = ActiveCell.Address
    Range("A2").Select
    Selection.Copy
    Range(fstAdd, lstAdd).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(fstAdd, lstAdd).Select
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
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Sheets(AggrDataShtName).Activate
    Cells(1, 1).Select
    fstCellAdd = ActiveCell.Address
    ActiveCell.End(xlToRight).Select
    lastCellAdd = ActiveCell.Address
    ActiveSheet.Range(fstCellAdd, lastCellAdd).Select
    If ActiveSheet.AutoFilterMode = True Then
         ActiveSheet.Range(fstCellAdd, lastCellAdd).AutoFilter
         End If
     
'Add a new sheet to create a Pivot Table
    Sheets.Add after:=Worksheets(Worksheets.Count)
    Set wsPtTable = Worksheets(Sheets.Count)
    'Set wsPtTable = Worksheets(3)
    wsptName = wsPtTable.name
    Sheets(wsptName).Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets(AggrDataShtName).Activate
    Set wsData = Worksheets(AggrDataShtName)
    Worksheets(AggrDataShtName).Activate
    sourceSheet = ActiveSheet.name
    ActiveSheet.Cells(1, 1).Select
    fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    Sheets(wsptName).Activate
    rngData = fstAdd & ":" & lstAdd
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
    TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion15
             
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtETTR"
    wsPtTable.Activate
    Set pt = wsPtTable.PivotTables("pvtETTR")
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
        
    With ActiveSheet.PivotTables("pvtETTR").PivotFields("Period")
        .Orientation = xlPageField
        .Position = 1
    End With
     ActiveSheet.PivotTables("pvtETTR").AddDataField ActiveSheet.PivotTables( _
    "pvtETTR").PivotFields("Avg. ETTR (days)"), "#ETTR (days)", xlAverage
        

    ActiveSheet.PivotTables("pvtETTR").PivotFields("Part12NC").PivotItems( _
    "Non-Parts Aggregated").Caption = "Non-Parts"

    ActiveSheet.PivotTables("pvtETTR").PivotFields("Part12NC").PivotItems( _
    "Parts Aggregated").Caption = "Parts"
       
    With ActiveSheet.PivotTables("pvtETTR")
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
    End With
    
    ActiveSheet.PivotTables("pvtETTR").PivotFields("SubSystem").Subtotals = _
    Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
    ActiveSheet.PivotTables("pvtETTR").PivotFields("BuildingBlock").Subtotals _
    = Array(False, False, False, False, False, False, False, False, False, False, False, False _
    )
    With pt.PivotFields("Part12NC")
            pf.Orientation = xlColumnField
            pf.Position = 1
    End With
    
    Set pvtTbl = Worksheets(wsptName).PivotTables("pvtETTR")
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtETTR")
        .ColumnGrand = True
        .RowGrand = True
    End With
    ActiveSheet.PivotTables("pvtETTR").PivotFields("SubSystem").RepeatLabels = _
    True
    ActiveSheet.PivotTables("pvtETTR").PivotFields("BuildingBlock").RepeatLabels _
    = True
    pvtTbl.RefreshTable
    
    Columns("A:E").EntireColumn.AutoFit
    Windows("CTS_KPI_Summary.xlsx").Activate
    Workbooks(myPvtWorkBook).Activate
    Range("A1").Select
    ActiveSheet.PivotTables("pvtETTR").Location = _
        "'[CTS_KPI_Summary.xlsx]ETTR'!$AK$3"
    Windows("CTS_KPI_Summary.xlsx").Activate
    Sheets("ETTR").Activate
    Range("AF3").Select
    ActiveCell.FormulaR1C1 = "=R[1]C[5]"
    Range("AF3").Select
    Selection.Copy
    Range("AF3,AF91").Select
    Range("AF3,AF3:AJ91").Select
    ActiveSheet.Paste
    Range("AE2").Select
    ActiveCell.value = "SS&BB"
    Range("AE3").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],RC[2])"
    Range("AE3").Select
    Selection.AutoFill Destination:=Range("AE3:AE91")
    
End Function

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'""""RRR Data-Merge All CCC Files and Create a RRR Report using FDV Raw Data

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Public Sub ListSubfoldersFile()

    Dim StrFile, mypath As String
    Dim objFSO, destRow As Long
    Dim mainFolder, mySubFolder
    Dim wbName As String
    Dim MyFiles(), DirArr() As String
    Dim FNum As Long
    Dim myBook As Workbook
    Dim wb As Workbook
    Dim BaseWks As Worksheet
    Dim ws As Worksheet
    Dim CalcMode, rowCount, baseItemCount As Long
    Dim cnt As Integer
    Dim lastRowCnt1 As Range
    Dim strDate As String
    Dim strDate1 As String
    Dim fstAdd As String
    Dim lstadd1 As String
    Dim lstadd2 As String
    Dim firstAdd, lastAdd As String
    Dim outPutFileFrstAdd As String
    Dim outPutFileLstAdd As String
 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    mypath = ThisWorkbook.Path & "\RRR_Input_Files\"
    Set mainFolder = objFSO.GetFolder(mypath)
    StrFile = Dir(mypath & "*.xlsx*")
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
            wb.Sheets(1).UsedRange.Find(what:="CaseID", lookat:=xlWhole).Select

            'ActiveSheet.Cells(3, 1).Select
    
            If cnt <= 1 Then
                fstAdd = ActiveCell.Address
            Else
                fstAdd = ActiveCell.Offset(1, 0).Address
            End If
                strDate = Cells(1, 1).value
                strDate1 = Mid(strDate, 43, 8)
                ActiveSheet.Range(fstAdd, lstadd1).Select
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

    'Workbooks(wbName).Sheets(1).Name = "RR"
    Workbooks(wbName).SaveAs _
    fileName:=ThisWorkbook.Path & "\" & "CTS-Cost to Serve_RRR_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    srcfile = ThisWorkbook.Path & "\" & "CTS-Cost to Serve_RRR_" & _
    Format(Now(), "yyyy-mm-dd") & ".xlsx"
    Application.Workbooks.Open(srcfile).Activate
    Worksheets("MergedCCC1").Activate
    wbName1 = ActiveWorkbook.name
    Sheets("MergedCCC1").UsedRange.Find(what:="RR", lookat:=xlWhole).Select
    Selection.EntireColumn.Select
    Selection.Replace what:="100", Replacement:="1", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Sheets("MergedCCC1").UsedRange.Find(what:="PCQL", lookat:=xlWhole).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.EntireColumn.Delete
        Sheets("MergedCCC1").UsedRange.Find(what:="Period", lookat:=xlWhole).Select

       ' ActiveSheet.Cells(1, 1).Select
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
    Application.ScreenUpdating = False
    
    Application.DisplayAlerts = False


    Dim myFile As String

    Dim inputItem As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    mypath = ThisWorkbook.Path & "\RRR_Input_Files\"
    myFile = "BCTool_SWO_RawData_SingleVersionOfTheTruth"
    inputItem = mypath & "\" & Dir(mypath & "\" & "BCTool_SWO_RawData_SingleVersionOfTheTruth" & "*.xls*") 'input file path
    Application.Workbooks.Open (inputItem), False
    myworkbook = ActiveWorkbook.name
    
    Workbooks(myworkbook).Activate
    With ThisWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "CombinedFDV"
    End With

    ActiveWorkbook.Sheets("SWO CV").Activate
    ActiveSheet.UsedRange.Find(what:="SWO", lookat:=xlWhole).Select

    'ActiveSheet.Cells(1, 2).Select
    firstAdd = ActiveCell.Address
    ActiveCell.SpecialCells(xlCellTypeLastCell).Select
    lastAdd = ActiveCell.Address
    Range(firstAdd, lastAdd).Select
    Selection.Copy
    Sheets("CombinedFDV").Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveCell.PasteSpecial xlPasteAll
        
    ActiveWorkbook.Sheets("SWO GXR").Activate
    ActiveSheet.UsedRange.Find(what:="SWO", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
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
        If Not Workbooks(myworkbook).Sheets("CombinedFDV").UsedRange.Find(what:=findVal, lookat:=xlWhole) Is Nothing Then
            Workbooks(myworkbook).Activate
            Sheets("CombinedFDV").Activate
            Workbooks(myworkbook).Sheets("CombinedFDV").UsedRange.Find(what:=findVal, lookat:=xlWhole).Select
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
Workbooks(myworkbook).Activate
Sheets("CombinedFDV").Activate
Application.CutCopyMode = False
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
ActiveCell.EntireColumn.Delete
ActiveCell.Offset(0, -1).Select
Selection.EntireColumn.Select
Selection.Replace what:="100", Replacement:="1", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


ActiveWorkbook.Save

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


Dim lstAdd As String

    Workbooks(wbName1).Activate

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
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "RRRPvtTablePvtChart"

    wsptName = ActiveSheet.name
    Sheets("RRRPvtTablePvtChart").Activate
    ActiveSheet.Cells(1, 1).Select
    fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.Sheets("RRR_Report").Activate

    Set wsData = Worksheets("RRR_Report")
    Worksheets("RRR_Report").Activate
    sourceSheet = ActiveSheet.name

    Sheets("RRRPvtTablePvtChart").Activate
    rngData = fstCellAdd & ":" & addofLstClmn
          
    Workbooks(myPvtWorkBook).Connections.Add2 _
        "WorksheetConnection_" & "RRR_Report" & "!" & rngData, "", _
        "WORKSHEET;" & mypath & "\[myPvtWorkBook]RRR_Report" _
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
        
     ' Sheets(wsptName).name = "RRRPvtTablePvtChart"
    ActiveWorkbook.Save
    Call PrtChartCalculation
End Sub
Public Sub PrtChartCalculation()
Dim StrFile, mypath, myFile As String
Dim objFSO, destRow As Long
Dim wb As Workbook
Dim inputItem As String
Dim rng As Range
Dim chtObject As ChartObject
Dim rownum As Integer
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    mypath = ThisWorkbook.Path
    inputItem = mypath & "\" & Dir(mypath & "\" & "CTS-Cost to Serve_RRR_*" & "*.xls*") 'input file path
    Application.Workbooks.Open (inputItem), False
    myPvtWorkBook = ActiveWorkbook.name
    
    Application.Workbooks(myPvtWorkBook).Activate

    Sheets("RRRPvtTablePvtChart").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("")
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("")
    ActiveSheet.UsedRange.Find(what:="Row Labels", lookat:=xlWhole).Select
    fstAdd = ActiveCell.Offset(1, 0).Address
    lstAdd = ActiveCell.Offset(1, 1).Address
        'Row Labels
    Range(fstAdd, lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    With ThisWorkbook
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "ParetoChart"
    partShtName = ActiveSheet.name
    End With
    Sheets("ParetoChart").Select
    Range("A2").Select
    Selection.PasteSpecial xlPasteValues
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    Selection.EntireRow.Delete
    Cells(1, 1).value = "SS"
    Cells(1, 2).value = "Total"
    Sheets("RRRPvtTablePvtChart").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[0]")
    Range(fstAdd, lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
   ' Range("A6:B28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ParetoChart").Select
    Range("C2").Select
    Selection.PasteSpecial xlPasteValues
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    Range(Selection, Selection.Offset(0, 1)).Delete
    Sheets("RRRPvtTablePvtChart").Select
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[1]")
    Range(fstAdd, lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    'Range("A6:B22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ParetoChart").Activate
    Range("E2").Select
    ActiveSheet.Paste
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    Range(Selection, Selection.Offset(0, 1)).Delete
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
    Range("A2").Select
    fstAdd = ActiveCell.Address(False, False)
    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    lstadd1 = Range(lstAdd).End(xlDown).Address(False, fasle)
    Range(lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range(lstAdd) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range(fstAdd, lstadd1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
    fstAdd = ActiveCell.Address(False, False)
    lstAdd = ActiveCell.Offset(0, 1).Address(False, False)
    lstadd1 = Range(lstAdd).End(xlDown).Address(False, fasle)
    Range(lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range(lstAdd) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range(fstAdd, lstadd1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
    vlkupAFstAdd = ActiveCell.Address(True, True)
    vlkupALstAdd = ActiveCell.End(xlDown).Address(True, True)
    Range("E2").Select
    vlkupEFstAdd = ActiveCell.Address(True, True)
    vlkupFLstAdd = ActiveCell.End(xlDown).Offset(0, 1).Address(True, True)
    Range("C2").Select
    fstAdd = ActiveCell.Address
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C1:R24C1,R2C5:R24C6,2,0)"
    ActiveCell.Formula = "=VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & vlkupEFstAdd & ":" & vlkupFLstAdd & ",2,0)"
    '=VLOOKUP($A$2:$A$24,$E$2:$F$24,2,0)
    lstAdd = Range(fstAdd).Offset(0, -1).End(xlDown).Offset(0, 1).Address(False, False)
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
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
    Selection.AutoFill Destination:=Range(Range(fstAdd).Offset(0, 1).Address, Range(lstAdd).Offset(0, 1).Address)
    Calculate
    Range(Range(fstAdd).Offset(0, 1).Address, Range(lstAdd).Offset(0, 1).Address).Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:="#N/A", Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Range(Range(fstAdd).Offset(0, 2).Address, Range(fstAdd).End(xlDown).Offset(0, 3).Address).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range("E2:F24").Select
    'ActiveWindow.ScrollRow = 1
    'Range("C2:D2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
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
    'ActiveWorkbook.Names.Add name:="NRRTtl", RefersToR1C1:= _
        "=SUM(ParetoChart!R2C3:R24C3)"
    ActiveWorkbook.Names.Add name:="NRRTtl", RefersTo:= _
        "=SUM(" & partShtName & "!" & Range(fstCellAdd).Address & ":" & Range(fstCellAdd).End(xlDown).Address & ")"
    ActiveWorkbook.Names("NRRTtl").Comment = ""
 
    Cells(1, 8).Select
    Do Until ActiveCell.Address = percentlstCellAdd
        ActiveCell.Offset(1, 0).FormulaR1C1 = "=RC[-1]/NRRTtl*100"
        ActiveCell.Offset(1, 0).Select
    Loop

    Application.ScreenUpdating = False
    Range("A2").Select
    fstAdd = ActiveCell.Address
    Selection.End(xlDown).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Set myrange = Range(fstAdd, lstAdd)
'    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart(250, xlColumnStacked).Select
    Dim prtchrt As String
prtchrt = partShtName & "!" & fstAdd & "," & lstAdd
prtchrtVal = partShtName & "!" & Range(fstAdd).Offset(0, 1).Address & ":" & Range(lstAdd).Offset(0, 1).Address
prtchrtXVal = partShtName & "!" & fstAdd & ":" & lstAdd
prtchrtValRR = partShtName & "!" & Range(fstAdd).Offset(0, 3).Address & ":" & Range(lstAdd).Offset(0, 3).Address
prtchrtValNNRR = partShtName & "!" & Range(fstAdd).Offset(0, 7).Address & ":" & Range(lstAdd).Offset(0, 7).Address

    ActiveChart.SetSourceData Source:=Range(prtchrt)
'    ActiveChart.ChartTitle.Select
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).name = "=""Total"""
    ActiveChart.FullSeriesCollection(1).Values = prtchrtVal
    ActiveChart.FullSeriesCollection(1).XValues = prtchrtXVal
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).name = "=""RR"""
    ActiveChart.FullSeriesCollection(2).Values = prtchrtValRR
    ActiveChart.FullSeriesCollection(2).XValues = prtchrtXVal
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).name = "=""% NON RR"""
    ActiveChart.FullSeriesCollection(3).Values = prtchrtValNNRR
    ActiveChart.FullSeriesCollection(3).XValues = prtchrtXVal
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
    Range("A2").Select
    Selection.End(xlDown).Select
    chrtSizeAdd = ActiveCell.Address
    Dim Ch As ChartObject
    Set Ch = Worksheets("ParetoChart").ChartObjects(1)
    With Ch
        .Top = Range(chrtSizeAdd).Offset(3, 0).Top
        .Width = Range(Range(chrtSizeAdd).Offset(3, 0).Address, Range(chrtSizeAdd).Offset(3, 12).Address).Width
        .Height = Range(Range(chrtSizeAdd).Offset(3, 0).Address, Range(chrtSizeAdd).Offset(21, 12).Address).Height
    End With
       ActiveSheet.ChartObjects(1).Left = ActiveSheet.Columns(1).Left
   ActiveSheet.ChartObjects(1).Top = ActiveSheet.Range(Range(chrtSizeAdd).Offset(3, 0).Address).Top

    Sheets("RRRPvtTablePvtChart").Select
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("")
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "[Range].[MaterialDescription].[MaterialDescription]").VisibleItemsList = Array _
        ("")
    ActiveSheet.PivotTables("PivotTable1").PivotFields("[Range].[RR].[RR]"). _
        VisibleItemsList = Array("[Range].[RR].&[0]")
    
    
    ActiveSheet.UsedRange.Find(what:="Row Labels", lookat:=xlWhole).Select
    fstAdd = ActiveCell.Address
    lstAdd = ActiveCell.Offset(0, 1).Address
        'Row Labels
    Range(fstAdd, lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Sheets("ParetoChart").Select
    Range("J1").Select
    Selection.PasteSpecial xlPasteValues
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    Range(Selection, Selection.Offset(0, 1)).Delete
    Range("J2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Range("J2").Select
    barChartSrtFstAdd = ActiveCell.Address
    Selection.End(xlDown).Select
    barChartSrtLstAdd = ActiveCell.Address
    barChartSrtValFstAdd = Range(barChartSrtFstAdd).Offset(0, 1).Address
    barChartSrtValLstAdd = Range(barChartSrtLstAdd).Offset(0, 1).Address

    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range( _
        barChartSrtValFstAdd, barChartSrtValLstAdd), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range(barChartSrtFstAdd, barChartSrtValLstAdd)
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
    Range(fstAdd, lstAdd).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("ParetoChart").Select
    Range("L1").Select
    Selection.PasteSpecial xlPasteValues
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    Range(Selection, Selection.Offset(0, 1)).Delete
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ParetoChart").Sort.SortFields.Add Key:=Range( _
        Range(barChartSrtValFstAdd).Offset(0, 2).Address, Range(barChartSrtValLstAdd).Offset(0, 2).Address), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ParetoChart").Sort
        .SetRange Range(Range(barChartSrtValFstAdd).Offset(0, 1).Address, Range(barChartSrtValLstAdd).Offset(0, 2).Address)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(C[-2],C[-4],1,0),RC[-4])"
    Range("N2").Select
    Selection.AutoFill Destination:=Range(Range(barChartSrtValFstAdd).Offset(0, 3).Address, Range(barChartSrtValLstAdd).Offset(0, 3).Address)
    'Range("N2:N28").Select
    Calculate
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-1],C[-5]:C[-4],2,0)"
    Range("O2").Select
    Selection.AutoFill Destination:=Range(Range(barChartSrtValFstAdd).Offset(0, 4).Address, Range(barChartSrtValLstAdd).Offset(0, 4).Address)
    Range(Range(barChartSrtValFstAdd).Offset(0, 4).Address, Range(barChartSrtValLstAdd).Offset(0, 4).Address).Select
    Calculate
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-3]"
    Range("P2").Select
    Selection.AutoFill Destination:=Range(Range(barChartSrtValFstAdd).Offset(0, 5).Address, Range(barChartSrtValLstAdd).Offset(0, 5).Address)
    Range(Range(barChartSrtValFstAdd).Offset(0, 5).Address, Range(barChartSrtValLstAdd).Offset(0, 5).Address).Select
    'Range("P2:P28").Select
   
    Range("Q2:R2").Select
    Selection.ClearContents
    Columns("O:O").Select
'    Sheets("Sheet5").Select
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "SS"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Count of SWO NON RR"
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
    ActiveWorkbook.Names.Add name:="NRRTtl", RefersTo:= _
        "=SUM(" & partShtName & "!" & Range(percentFstCellAdd).Address & ":" & Range(percentFstCellAdd).End(xlDown).Address & ")"
    ActiveWorkbook.Names("NRRTtl").Comment = ""
 
'    ActiveWorkbook.Names.Add name:="NRRTtl", RefersToR1C1:= _
        "=SUM(ParetoChart!R2C2:R24C2)"
 '   ActiveWorkbook.Names("NRRTtl").Comment = ""

    Cells(1, 9).Select
        Do Until ActiveCell.Address = percentlstCellAdd
        ActiveCell.Offset(1, 0).FormulaR1C1 = "=RC[-5]/NRRTtl*100"
        ActiveCell.Offset(1, 0).Select
    Loop
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "% RR"
 
    Sheets("ParetoChart").Select
    
    Range("J2").Select
    fstAdd = ActiveCell.Address
    Selection.End(xlDown).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Set myrange = Range(fstAdd, Range(lstAdd).Offset(0, 1).Address)
    Range(Selection, Selection.End(xlDown)).Select
    'ActiveSheet.Shapes.AddChart(250, xlColumnStacked).Select
prtchrt = partShtName & "!" & fstAdd & ":" & Range(lstAdd).Offset(0, 1).Address
prtchrtVal = partShtName & "!" & Range(fstAdd).Offset(0, 1).Address & ":" & Range(lstAdd).Offset(0, 1).Address
prtchrtXVal = partShtName & "!" & fstAdd & ":" & lstAdd
prtchrtValRR = partShtName & "!" & Range(fstAdd).Offset(0, 3).Address & ":" & Range(lstAdd).Offset(0, 3).Address
prtchrtValPRNNRR = partShtName & "!" & Range(fstAdd).Offset(0, 6).Address & ":" & Range(lstAdd).Offset(0, 6).Address
prtchrtValNNRR = partShtName & "!" & Range(fstAdd).Offset(0, 7).Address & ":" & Range(lstAdd).Offset(0, 7).Address

    Range("J2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(201, xl3DColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(prtchrt)
    'ActiveChart.Axes(xlValue).MajorGridlines.Select
    'ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).name = "=""NON-RR Total"""
    ActiveChart.FullSeriesCollection(1).Values = prtchrtVal
    ActiveChart.FullSeriesCollection(1).XValues = prtchrtXVal
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).name = "=""NON Parts-NON RR"""
    ActiveChart.FullSeriesCollection(2).Values = prtchrtValRR
    ActiveChart.FullSeriesCollection(2).XValues = prtchrtXVal
    
    ActiveChart.seriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).name = "=""Part Replaced NON RR"""
    ActiveChart.FullSeriesCollection(3).Values = prtchrtValPRNNRR
    ActiveChart.FullSeriesCollection(3).XValues = prtchrtXVal
         
    ActiveChart.ApplyLayout (4)
    ActiveChart.ChartArea.Select
    With Selection
    .Width = 500
    .Height = 300
    
    End With
    'ActiveChart.ChartType = xl3DColumnClustered
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
        .ForeColor.RGB = RGB(255, o, o)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.ChartArea.Select
    Application.CommandBars("Format Object").Visible = False
'End With
   
    Range("N2").Select
    Selection.End(xlDown).Select
    chrtSizeAdd = ActiveCell.Address
    Set Ch = Worksheets("ParetoChart").ChartObjects(2)
    With Ch
        .Top = Range(chrtSizeAdd).Offset(3, 0).Top
        .Width = Range(Range(chrtSizeAdd).Offset(3, 0).Address, Range(chrtSizeAdd).Offset(3, 10).Address).Width
        .Height = Range(Range(chrtSizeAdd).Offset(3, 0).Address, Range(chrtSizeAdd).Offset(21, 10).Address).Height
    End With
   ActiveSheet.ChartObjects(2).Left = ActiveSheet.Columns(14).Left
   ActiveSheet.ChartObjects(2).Top = ActiveSheet.Range(Range(chrtSizeAdd).Offset(3, 0).Address).Top
   ActiveSheet.ChartObjects(2).Select
   ActiveChart.SetElement (msoElementPrimaryValueGridLinesMajor)
   ActiveChart.Axes(xlValue).Select
   ActiveChart.Axes(xlValue).MajorUnit = 50
   ActiveChart.ChartArea.Select
   ActiveWorkbook.Save
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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
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
    prdFileName = KPISheetName

'Open Aggregated Data File
    inputFileGlobal = prdFileName & ".xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & inputFileGlobal
        inputFlName = inputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path inputFileGlobal
        inputPath = sharedDrivePath
        inputFlName = inputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(inputFileGlobal).Windows(1).Visible = True
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
    'On Error Resume Next
     
'Delete Blank sheets from aggregated data file if any
    For Each sh In Sheets
        If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        
    Next sh
     
    Call DataBrekUpFrPivot

'Filter the Buildingblocks Aggregated data and delete the Buildingblocks Aggregated data
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
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
      
'Remove the values which are less then 10% of the top value in the Total Calls(#) column
    ActiveSheet.UsedRange.Find(what:="Total Calls (#)", lookat:=xlWhole).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    fstFiltCellAdd = ActiveCell.Address
    Range(fstAdd).Offset(1, 0).End(xlDown).Select
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
     
       Dim Part12NCClmn As Range
       Set Part12NCClmn = Sheets(AggrDataShtName).rows(2).Find("Part12NC", , , xlWhole, , , , False)
    
      If Not Part12NCClmn Is Nothing Then
        Application.ScreenUpdating = False
        Part12NCClmn.Offset(1, 0).Select
        Part12NcClmnAdd = ActiveCell.Address(False, False)
      End If
        
       Dim ttlCalls As Range
       Set ttlCalls = Sheets(AggrDataShtName).rows(2).Find("Total Calls (#)", , , xlWhole, , , , False)
    
      If Not ttlCalls Is Nothing Then
        Application.ScreenUpdating = False
        ttlCalls.Offset(1, 0).Select
        ttlCallsAdd = ActiveCell.Address(False, False)
      End If
        
       Dim AvgMTTRprCallHrs As Range
       Set AvgMTTRprCallHrs = Sheets(AggrDataShtName).rows(2).Find("Avg. MTTR/Call (hrs)", , , xlWhole, , , , False)
    
      If Not AvgMTTRprCallHrs Is Nothing Then
        Application.ScreenUpdating = False
        AvgMTTRprCallHrs.Offset(1, 0).Select
        AvgMTTRprCallHrsAdd = ActiveCell.Address(False, False)
      End If
            
       Dim visitsprCallNP As Range
       Set visitsprCallNP = Sheets(AggrDataShtName).rows(2).Find("# of calls with 1 visit", , , xlWhole, , , , False)
    
      If Not visitsprCallNP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallNP.Offset(1, 0).Select
        visitsprCallNPAdd = ActiveCell.Address(False, False)
      End If
      
       Dim visitsprCallP As Range
       Set visitsprCallP = Sheets(AggrDataShtName).rows(2).Find("Calls = 0 Visit", , , xlWhole, , , , False)
    
      If Not visitsprCallP Is Nothing Then
        Application.ScreenUpdating = False
        visitsprCallP.Offset(1, 0).Select
        visitsprCallPAdd = ActiveCell.Address(False, False)
      End If
      
      Dim avgVisitsprCallNP As Range
       Set avgVisitsprCallNP = Sheets(AggrDataShtName).rows(2).Find("Avg. Visits/call (#)", , , xlWhole, , , , False)
    
      If Not avgVisitsprCallNP Is Nothing Then
        Application.ScreenUpdating = False
        avgVisitsprCallNP.Offset(1, 0).Select
        avgVisitsprCallAdd = ActiveCell.Address(False, False)
      End If
      
       Dim ttlCostprPartNP As Range
       Set ttlCostprPartNP = Sheets(AggrDataShtName).rows(2).Find("Total Costs/part (EUR)", , , xlWhole, , , , False)
    
      If Not ttlCostprPartNP Is Nothing Then
        Application.ScreenUpdating = False
        ttlCostprPartNP.Offset(1, 0).Select
        ttlCostprPartAdd = ActiveCell.Address(False, False)
      End If
      
'Add one column for "Total Cost of Parts & Non-Parts"

  Dim found As Range
  Set found = Sheets(AggrDataShtName).rows(2).Find("Total Costs/part (EUR)", , , xlWhole, , , , False)
    
    If Not found Is Nothing Then
        Application.ScreenUpdating = False
        found.Offset(, 1).Resize(, 1).EntireColumn.Insert
  
  End If
  
        Workbooks(myPvtWorkBook).Sheets(AggrDataShtName).Activate

        found.End(xlDown).Select
        ActiveCell.Offset(0, 1).Select
        ttlCstLstAdd = ActiveCell.Address
        found.Offset(, 1).value = "Total Cost of Parts & Non-Parts"
        found.Offset(1, 1).Select
        ttlCstAdd = ActiveCell.Address
   
        'ActiveCell.Offset(, 0).Formula = "=IFERROR(IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallPAdd & "*" & 200 & ")," & "IF(" & Part12NcClmnAdd & Chr(61) & Chr(34) & "Non-Parts Aggregated" & Chr(34) & ",(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 100 & ")+(" & visitsprCallNPAdd & "*" & 200 & "))),0)"
        ActiveCell.Offset(, 0).Formula = "=(" & ttlCallsAdd & "*" & AvgMTTRprCallHrsAdd & "*" & 105 & ")+(" & avgVisitsprCallAdd & "*" & 200 & ")+(" & ttlCostprPartAdd & "*" & ttlCallsAdd & ")"

        Range(ttlCstAdd).Select
        Selection.Copy
        Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteFormulas
        Range(ttlCstAdd, ttlCstLstAdd).Select
        Selection.Copy
        Range(ttlCstAdd, ttlCstLstAdd).PasteSpecial xlPasteValues

'Add a new sheet to create a Pivot Table
        Sheets.Add after:=Worksheets(Worksheets.Count)

        Set wsPtTable = Worksheets(Sheets.Count)

        wsptName = wsPtTable.name
        Sheets(wsptName).Activate
        ActiveSheet.Cells(1, 1).Select
        fstadd1 = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveWorkbook.Sheets(AggrDataShtName).Activate

        Set wsData = Worksheets(AggrDataShtName)
        Worksheets(AggrDataShtName).Activate
        sourceSheet = ActiveSheet.name
        ActiveSheet.Cells(2, 1).Select
        Selection.EntireColumn.Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ActiveCell.Offset(1, 0).Select
        ActiveCell.value = "Period"
        ActiveCell.Offset(0, 1).value = "Period1"
        fstAdd = ActiveCell.Offset(1, 0).Address
        ActiveCell.Offset(0, 1).Select
        ActiveCell.End(xlDown).Select
        ActiveCell.Offset(0, -1).Select
        lstAdd = ActiveCell.Address
        Cells(3, 1).Select

    ActiveCell.FormulaR1C1 = "=MID(RC[1],1,4)&""-""&MID(RC[1],5,2)"
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(2, 1).Select
    
        fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        ActiveCell.End(xlDown).Select
        ActiveCell.End(xlToRight).Select

        lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
        
        Sheets(wsptName).Activate
        rngData = fstAdd & ":" & lstAdd
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=wsptName & "!" & fstadd1, TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
        Cells(2, 1).Select
        ActiveCell.PivotTable.name = "pvtKPIMASTER"
        wsPtTable.Activate
               
        Set pt = wsPtTable.PivotTables("pvtKPIMASTER")
        
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
        ActiveSheet.PivotTables("pvtKPIMASTER").CalculatedFields.Add _
        "Avg. MTTR/Call (hrs)/12", "='Avg. MTTR/Call (hrs)' /12", True
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Avg. MTTR/Call (hrs)/12"). _
        Orientation = xlDataField
        ActiveSheet.PivotTables("pvtKPIMASTER").DataPivotField.PivotItems( _
        "Sum of Avg. MTTR/Call (hrs)/12").Caption = "#Avg. MTTR/Call (hrs)/12"
        ActiveSheet.PivotTables("pvtKPIMASTER").CalculatedFields.Add "Avg. ETTR (days)/12" _
        , "='Avg. ETTR (days)' /12", True
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Avg. ETTR (days)/12"). _
        Orientation = xlDataField
        ActiveSheet.PivotTables("pvtKPIMASTER").DataPivotField.PivotItems( _
        "Sum of Avg. ETTR (days)/12").Caption = "#Avg. ETTR (days)/12"
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "'#Avg. MTTR/Call (hrs)/12'", _
        xlDataAndLabel, True
        With ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields( _
            "#Avg. MTTR/Call (hrs)/12")
            .NumberFormat = "0.00"
        End With
    
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "'#Avg. ETTR (days)/12'", _
        xlDataAndLabel, True
        With ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("#Avg. ETTR (days)/12")
            .NumberFormat = "0.00"
        End With
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Avg. Visits/call (#)"), "Visits/call (#)", xlAverage
        With ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Visits/call (#)")
            .NumberFormat = "0.00"
        End With
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC").PivotItems( _
        "Non-Parts Aggregated").Caption = "Non-Parts"

        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC").PivotItems( _
        "Parts Aggregated").Caption = "Parts"
        ActiveSheet.PivotTables("pvtKPIMASTER").AddDataField ActiveSheet.PivotTables( _
        "pvtKPIMASTER").PivotFields("Total Cost of Parts & Non-Parts"), _
        "#Total Costs", xlSum
    
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
     ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC").ShowAllItems = _
        True
    pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts"
    With ActiveSheet.PivotTables("pvtKPIMASTER")
        .ColumnGrand = True
        .RowGrand = False
    End With
    
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIMASTER")
    Set pf = pvtTbl.PivotFields("Part12NC")

        pf.ClearAllFilters
        pf.EnableMultiplePageItems = True
    
'    pf.PivotItems("Parts/Non-Parts Breakups").Visible = False
    ActiveSheet.PivotTables("pvtKPIMASTER").HasAutoFormat = False
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotSelect "", xlDataAndLabel, True
    Selection.ColumnWidth = 8
   
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts"). _
        ShowDetail = False
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail _
        = False
    
    With ActiveSheet.PivotTables("pvtKPIMASTER")
        .ColumnGrand = True
        .RowGrand = False
    End With
     ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("SubSystem").RepeatLabels = _
        True
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock"). _
        RepeatLabels = True
        ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts"). _
        RepeatLabels = True
        
    pvtTbl.RefreshTable

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIMASTER")
    
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("SubSystem").ShowDetail = True
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail = _
        True

' Add ConditionalFormatting of Data Bars on total calls of Parts and Non parts
     ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").ShowDetail = _
        False
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    'Range("E4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 4).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    
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
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 1).Select
    
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 5).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    
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

    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail = _
        False
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
'
    'Range("E4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 4).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
    
    ActiveSheet.UsedRange.Find(what:="# of Calls", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    'Range("F4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 5).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
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
  
   ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail = _
        True
   
'Add conditional formatting on MTTR and ETTR Calls
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

     'Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 6).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    lstadd1 = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. MTTR/Call (hrs)/12"",R3C1,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    
    'Range("H4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 7).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. ETTR (days)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    'Range("I4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 8).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstJAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",R3C1,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
     Range(fstAdd, lstJAdd).Select
     Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstJAdd & "," & 20 & ")"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",$A$3,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""#Avg. ETTR (days)/12"",$A$3,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="Visits/call (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    'Range("K4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 10).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstLAdd = ActiveCell.Address
    
    Range(fstAdd, lstLAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""Visits/call (#)"",$A$3,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""Visits/call (#)"",$A$3,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Total Costs", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    PfstAdd = Range(fstAdd).Offset(0, 1).Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 12).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    PlstAdd = Range(lstAdd).Offset(0, 1).Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$M$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range(PfstAdd, PlstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$N$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail = _
        False
   
     ActiveSheet.UsedRange.Find(what:="#Total Costs", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    PfstAdd = Range(fstAdd).Offset(0, 1).Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 12).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    PlstAdd = Range(lstAdd).Offset(0, 1).Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$M$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range(PfstAdd, PlstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Total Costs"",$N$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

'    Range("G4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 6).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. MTTR/Call (hrs)/12"",$G$4,""Part12NC"",""Non-Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. MTTR/Call (hrs)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Offset(0, 1).Select
    'Range("H4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 7).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd, lstAdd).Select

    Selection.FormatConditions.AddAboveAverage
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="#Avg. ETTR (days)/12", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select
   ' Range("I4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 8).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstJAdd = ActiveCell.Address
    
    Range(fstAdd, lstAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",$I$4,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
     Range(fstAdd, lstJAdd).Select
     Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=LARGE(" & fstAdd & ":" & lstJAdd & "," & 10 & ")"
    'Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""#Avg. ETTR (days)/12"",R3C1,""Part12NC"",""Non-Parts"")/100*10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveSheet.UsedRange.Find(what:="Visits/call (#)", lookat:=xlWhole).Select
    ActiveCell.Offset(2, 0).Select

    'Range("K4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
    ActiveCell.Offset(0, 10).Select
    ActiveCell.Offset(-1, 0).Select
    lstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    lstLAdd = ActiveCell.Address
    
    Range(fstAdd, lstLAdd).Select
    Range(fstAdd, lstLAdd).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=GETPIVOTDATA(""Visits/call (#)"",$K$4,""Part12NC"",""Non-Parts"")+GETPIVOTDATA(""Visits/call (#)"",$L$4,""Part12NC"",""Parts"")/100*20"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range(fstAdd, lstAdd).Select
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
    
    Columns("A:D").Select
    With Selection
        .ColumnWidth = 8
    End With
    Cells(1, 1).Select
    
    Worksheets(wsptName).PivotTables("pvtKPIMASTER").PreserveFormatting = False
    Sheets(wsptName).name = "PivotTableAggData"
    'ActiveWindow.Zoom = 85
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("BuildingBlock").ShowDetail = _
        True
    
    ActiveSheet.PivotTables("pvtKPIMASTER").PreserveFormatting = True
    ActiveSheet.name = "PivotTableAggData"
     'ActiveWindow.Zoom = 85

    Workbooks(myPvtWorkBook).Activate
    Set pvtTbl = ActiveSheet.PivotTables("pvtKPIMASTER")
     
'Open Output file CTS_KPI_Summary.xlsx
    outputFileGlobal = "CTS_KPI_Summary.xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        outputPath = ThisWorkbook.Path & "\" & outputFileGlobal
        outputFlName = outputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path outputFileGlobal
        outputPath = sharedDrivePath
        outputFlName = outputFileGlobal
    End If

    Application.Workbooks.Open (outputPath), False
    Application.Workbooks(outputFileGlobal).Windows(1).Visible = True
    
    Workbooks(outputFlName).Activate
    Sheets("KPI-Master").Select
    Range("A2").Select
    'pvtmstName = ActiveCell.PivotTable.name
    'Set pvtMstTbl = ActiveSheet.PivotTables(pvtmstName)
    'ActiveSheet.PivotTables(pvtmstName).PivotSelect "", xlDataAndLabel, True
    'ActiveSheet.PivotTables(pvtmstName).CalculatedFields("RRR").Delete
    'ActiveSheet.Shapes.Range(Array("Period")).Select
    'Selection.Delete

'Delete Slicers
    Dim slCaches As SlicerCaches
    Dim slCache As SlicerCache

    Set slCaches = ActiveWorkbook.SlicerCaches

    For Each slCache In slCaches

        If slCache.PivotTables.Item(1).name = "pvtKPIMASTER" Then
           slCache.Delete
        End If

    Next slCache

    Cells.Select
    Selection.Delete
    Workbooks(myPvtWorkBook).Activate
    pvtTbl.TableRange2.Copy
    
    Workbooks(outputFlName).Activate

    Sheets("KPI-Master").Select
    Range("a1").PasteSpecial
    Workbooks(myPvtWorkBook).Activate
    Workbooks(outputFlName).Activate
    Sheets("KPI-Master").Activate
        
'IB value for 12 months from IB sheet
    Sheets("IB").Select
    Range("N1").Select
    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")

    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").ClearAllFilters
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").EnableMultiplePageItems _
        = False
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").CurrentPage = "(All)"
    ActiveSheet.PivotTables("pvtIB").PivotFields("ProductGroup").ClearAllFilters
    pvtTbl.PivotFields("ProductGroup").PivotFilters.Add Type:=xlCaptionContains, Value1:=KPISheetName

    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
   
    Set pvtTbl = ActiveSheet.PivotTables("pvtIB")
    Set pf = pvtTbl.PivotFields("Period")

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select
        abc = ActiveCell.Offset(0, 1).value
        IBVal = abc / 12
        
 'Add RRR% and CallRate Columns
    'Workbooks(outputFlName).Activate
    Sheets("KPI-Master").Select
    Range("A1").Select
    ActiveCell.PivotTable.name = "pvtKPIMASTER"
    ActiveSheet.UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "RRR%"
    ActiveCell.Offset(0, 1).value = "CallRate/Sys/Yr"
    Range(ActiveCell.Address, ActiveCell.Offset(0, 1).Address).Select
    Selection.EntireColumn.Select
    Selection.ColumnWidth = 10.14
    
    Application.ScreenUpdating = False
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("pvtKPIMASTER"), _
       "Period").Slicers.Add ActiveSheet, , "Period", "Period", 69.75, 810, 144, 700
    
    ActiveSheet.Shapes.Range(Array("Period")).Select
   ' ActiveSheet.Shapes("Period").IncrementLeft 530
    ActiveSheet.Shapes("Period").IncrementTop -75
    ActiveSheet.Shapes.Range(Array("Period")).Select

'        'With ActiveWorkbook.SlicerCaches("Slicer_Period")
'            For i = 1 To .SlicerItems.Count
'                .SlicerItems(i).selected = True
'            Next i
'
'        End With
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").ClearAllFilters
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
           
    ActiveSheet.UsedRange.Find(what:="CallRate/Sys/Yr", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
        ActiveCell.FormulaR1C1 = "=(R[]C[-11]+R[]C[-10])/" & IBVal
        fstAdd = ActiveCell.Address
        Selection.Copy
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select

        ActiveCell.Offset(0, 15).Select
        lstAdd = ActiveCell.Address
        Range(fstAdd, lstAdd).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range(fstAdd, lstAdd).Select
    
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
    
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 20
        .Percent = False
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
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
   Application.CutCopyMode = False
   

   Range("N3").Select
   Selection.Copy
   ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
   Range(ActiveCell.Address, ActiveCell.Offset(0, 1).Offset(-2, 0).Address).Select
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

'Create RRR% column and calculate RRR%
     Range("A2").Select
    pvtmstName = ActiveCell.PivotTable.name
    Set pvtMstTbl = ActiveSheet.PivotTables(pvtmstName)
    ActiveSheet.PivotTables(pvtmstName).PivotSelect "", xlDataAndLabel, True
    Selection.Copy

    Range("AD1").Select
    ActiveSheet.Paste
    pvtName = ActiveCell.PivotTable.name
    
    Set pvtTbl = ActiveSheet.PivotTables(pvtName)

    pvtTbl.DataPivotField.Orientation = xlHidden
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC").Orientation = _
        xlHidden

    ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Total Calls (#)"), "# of Calls", xlSum
    ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Calls = 0 Visit"), "Sum of Calls = 0 Visit", xlSum
    ActiveSheet.PivotTables(pvtName).PivotFields("Sum of Calls = 0 Visit"). _
        Caption = "#Calls = 0 Visit"
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").ShowAllItems = _
        True
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").ClearAllFilters

        pvtTbl.CalculatedFields.Add "RRR", _
        "='Calls = 0 Visit' /'Total Calls (#)' *100", True
        pvtTbl.PivotFields("RRR").Orientation = _
        xlDataField
        
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").ClearAllFilters
    ActiveSheet.PivotTables(pvtName).PivotSelect "", xlDataAndLabel, True
    Selection.Cut
    Range("AD2").Select
    ActiveSheet.Paste
    ActiveSheet.PivotTables(pvtName).PivotSelect "Subsystem['-']", xlDataAndLabel, True
    ActiveCell.Offset(0, -1).Select
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(-1, -1).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],RC[2],RC[3])"
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd).Select
    ActiveCell.Offset(0, -1).Select
    fstpvtAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    lstPvtAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Sum of RRR", lookat:=xlWhole).Select
    ActiveCell.End(xlDown).Select
    RRRLstRAdd = ActiveCell.Offset(-1, 0).Address
    Range(fstpvtAdd).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-27],RC[-26],RC[-25])"
    Selection.AutoFill Destination:=Range(fstpvtAdd, lstPvtAdd)
    ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    vlkpRRRFstAdd = ActiveCell.Address
    ActiveCell.Offset(0, 1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    vlkpRRRLstAdd = ActiveCell.Address
    
    fstpvtAdd = Range(fstpvtAdd).Address(False, False)
    Range(vlkpRRRFstAdd).Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[13]:R54C28,R5C29:R54C36,8,0),0)"
    Range(vlkpRRRFstAdd).Formula = "=IFERROR(VLOOKUP(" & fstpvtAdd & ":" & lstPvtAdd & "," & fstAdd & ":" & RRRLstRAdd & ",8,0)" & "," & 0 & ")"
    
    Range(vlkpRRRFstAdd).Select
    Selection.AutoFill Destination:=Range(vlkpRRRFstAdd, vlkpRRRLstAdd)

   ' Range(DVVlkpResultAdd).Formula = "=IFERROR(VLOOKUP(" & vlkpRRRFstAdd & "," & "'Designed Data'!" & fstAdd & ":" & CRValLstAdd & ",4,0)" & "," & 0 & ")"

    Application.CutCopyMode = False
    
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
    
    
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").PivotFilters.Add Type:=xlCaptionEquals, Value1:="-"
    ActiveSheet.PivotTables("pvtKPIMASTER").PivotFields("Part12NC-Sub Parts").EnableMultiplePageItems _
        = True
        
'Apply Icon Set Conditional formatting on RRR Column Values
    ActiveSheet.UsedRange.Find(what:="RRR%", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select

'    Range("O4").Select
    fstAdd = ActiveCell.Address
    ActiveCell.End(xlDown).Select
    lstAdd = ActiveCell.Address
    Range(fstAdd).Select
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
    Range(fstAdd, lstAdd).Select
    Selection.NumberFormat = "0"

'Add Headings to DashBoard
    
    Range("A1").Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ActiveCell.FormulaR1C1 = "KPI-Master Dash Board for " & KPISheetName & " - Excluding Bottom 10%"
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="SubSystem", lookat:=xlWhole).Select
    ActiveCell.End(xlToRight).Select
    ActiveCell.Offset(-3, 0).Select
    lstAdd = ActiveCell.Address
    
    Range(fstAdd, lstAdd).Select
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
        .Font.Bold = True
        .Font.Italic = True
        .Font.name = "Calibri"
        .Font.Size = 15
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.399975585192419
        .Interior.PatternTintAndShade = 0
    End With
'    Selection.Merge
    Selection.Font.Bold = True
    Range("A1").Select
    Selection.EntireRow.Select
    Selection.RowHeight = 25
    Range("A2").Select
    pvtmstName = ActiveCell.PivotTable.name
    Set pvtMstTbl = ActiveSheet.PivotTables(pvtmstName)
    On Error Resume Next
    'Filter the KPI-All data for previous 12 months (month selected by the user will be considered as start of the month)
    ActiveSheet.Shapes.Range(Array("Period")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Period").PivotTables.AddPivotTable (ActiveSheet.PivotTables(pvtmstName))
    Dim slSlicer As Slicer
    'Dim slCache As SlicerCache
    Dim slItem As slicerItem

    fixedDate = Sheet1.combYear.value
    endDate1 = Format(DateAdd("yyyy", -1, fixedDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Sheets("KPI-Master").Activate
    ActiveSheet.Shapes.Range(Array("Period")).Select

    'Application.ScreenUpdating = False
    With ActiveWorkbook.SlicerCaches("Slicer_Period")
        '--deselect all items except the first
        For Each slItem In .VisibleSlicerItems
            If slItem.name < endDate Or slItem.name > fixedDate Then
                slItem.selected = False
                Else
                slItem.selected = True
                End If
        Next slItem

    End With
    Cells.Select
    Selection.ColumnWidth = 8.5
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("AA1:AP1").Select
    Selection.EntireColumn.Select
    Selection.EntireColumn.Hidden = True
    Application.Workbooks(myPvtWorkBook).Close
    Application.Workbooks(outputFlName).Activate
    ActiveWindow.Zoom = 100
    ActiveWorkbook.Save
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
End Sub

Sub DataBrekUpFrPivot()
    ActiveWorkbook.Sheets(2).Activate
    AggrDataShtName = ActiveSheet.name
    Cells(1, 1).Select
    ActiveCell.UnMerge
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.AutoFilter
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
    fstAdd = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    lstAdd = ActiveCell.Address
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    Selection.Offset(1, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""All Aggregated"",RC[-1]=""Parts Aggregated"",RC[-1]=""Non-Parts Aggregated""),""-"",RC[-1])"
    Selection.AutoFill Destination:=Range(fstAdd, lstAdd)
    Range(fstAdd, lstAdd).Select
    Calculate
    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC", lookat:=xlWhole).Select
    Selection.Offset(1, 0).Select
    Cells(2, 1).Select
    Selection.EntireRow.Select
    Selection.Find(what:="Part12NC-Sub Parts", lookat:=xlWhole).Select
    Selection.Offset(1, 0).Select
    Do Until ActiveCell.value = ""
        If ActiveCell.value = "-" Then
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.Offset(0, -1).value = "Parts Aggregated"
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
    
End Sub

Public Sub IBPivotTable()
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
Dim mypath As String
Dim fstAdd As String
Dim lstAdd As String
Dim CTSProductName, dateValue, prdNameFile, filePresent As String
Dim fstFiltCellAdd, lastFiltCellAdd, fstFiltCellAdd1 As String

Dim xWs As Worksheet
Dim xpvt As PivotTable
Dim sh As Variant
Dim Max, tenPercentofMax, cellVal
Dim rows As Range, cell As Range, value As Long
Dim lastRow As Integer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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
    prdNameFile = KPISheetName

'Open Output file CTS_KPI_Summary.xlsx
    outputFileGlobal = "CTS_KPI_Summary.xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        outputPath = ThisWorkbook.Path & "\" & outputFileGlobal
        outputFlName = outputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path outputFileGlobal
        outputPath = sharedDrivePath
        outputFlName = outputFileGlobal
    End If

    Application.Workbooks.Open (outputPath), False
    Application.Workbooks(outputFileGlobal).Windows(1).Visible = True
    
    myCTSWorkBook = ActiveWorkbook.name

    ActiveWorkbook.Sheets("IB").Activate
    Cells(1, 1).Select
    Selection.EntireColumn.Select
    Selection.EntireRow.Select
    Selection.EntireRow.Delete
    sourceSheet = ActiveSheet.name
    
'Open IB Data File
    outputFileGlobal = "IB_IXR.xlsx"
    If Sheet1.rdbLocalDrive.value = True Then
        inputPath = ThisWorkbook.Path & "\" & outputFileGlobal
        inputFlName = outputFileGlobal
    End If

    If Sheet1.rdbSharedDrive.value = True Then
        SharedDrive_Path outputFileGlobal
        inputPath = sharedDrivePath
        inputFlName = outputFileGlobal
    End If

    Application.Workbooks.Open (inputPath), False
    Application.Workbooks(outputFileGlobal).Windows(1).Visible = True
    
    IBWorkBook = ActiveWorkbook.name
    Cells(1, 1).Select
    Selection.EntireColumn.Select
    Selection.EntireRow.Select
    ActiveSheet.UsedRange.Copy
    Workbooks(myCTSWorkBook).Activate
    Sheets("IB").Select
    Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
    Selection.RowHeight = 15
    Workbooks(IBWorkBook).Close
    Workbooks(outputFlName).Activate
    Sheets(sourceSheet).Select
    ActiveSheet.UsedRange.Find(what:="IBContract", lookat:=xlWhole).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Delete
    
    SearchForWords = Array("*Old*")
   
       With ActiveSheet
            Range("A2").AutoFilter Field:=5, Criteria1:=SearchForWords
            .Cells.SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
        End With

     SearchForWords1 = Array("*TBD*")
     Range("A1").Select
     Selection.EntireRow.Select
        With Selection
            .AutoFilter Field:=5, Criteria1:=SearchForWords1
           Range("A1").Cells.SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
        End With

'Create a Pivot Table
    ActiveSheet.UsedRange.Find(what:="IBTotal", lookat:=xlWhole).Select
    pvtAdd = ActiveCell.Offset(0, 2).Address(ReferenceStyle:=xlR1C1)
    ActiveSheet.UsedRange.Find(what:="Period", lookat:=xlWhole).Select
   
    fstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    ActiveCell.End(xlDown).Select
    ActiveCell.End(xlToRight).Select
    lstAdd = ActiveCell.Address(ReferenceStyle:=xlR1C1)
    rngData = fstAdd & ":" & lstAdd
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    sourceSheet & "!" & rngData, Version:=xlPivotTableVersion15).CreatePivotTable _
    TableDestination:=sourceSheet & "!" & pvtAdd, TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion15
             
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    ActiveCell.PivotTable.name = "pvtIB"
    Set pt = ActiveSheet.PivotTables("pvtIB")
    Set pf = pt.PivotFields("ProductGroup")
    pf.Orientation = xlRowField
    pf.Position = 1
    Set pf = pt.PivotFields("Period")
    pf.Orientation = xlPageField
    pf.Position = 1
    ActiveSheet.PivotTables("pvtIB").AddDataField ActiveSheet.PivotTables( _
    "pvtIB").PivotFields("IBTotal"), "#IBTotal", xlSum
    
    Set pvtTbl = Worksheets("IB").PivotTables("pvtIB")

    pvtTbl.PivotFields("ProductGroup").PivotFilters.Add Type:=xlCaptionContains, Value1:=KPISheetName
   
   With ActiveSheet.PivotTables("pvtIB")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRows
    End With
    ActiveSheet.PivotTables("pvtIB").PivotFields("Period").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("pvtIB").PivotFields("ProductGroup").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("pvtIB").PivotFields("ProductGroup").RepeatLabels = _
        True
        Workbooks(outputFlName).Save
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
Sub allSSsummarySheet()
'Delete Blank sheets from aggregated data file if any
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Workbooks(myworkbook).Activate
    'For Each s In ActiveWorkbook.Sheets
     '   If Left(s.name, 16) = "All SS-BB" Then
      '      s.Delete
       ' End If
    'Next s
   

    Sheets("CR").Select
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A2:AB3").Select
    Range("A3").Activate
    Selection.Copy
    Sheets("All SS-BB").Select
    Range("B2").Select
    ActiveSheet.Paste
    'ActiveWindow.Zoom = 85
    Columns("B:G").Select
    Selection.ColumnWidth = 12
    Columns("H:AB").Select
    Selection.ColumnWidth = 7
    Columns("AC:AC").Select
    Selection.ColumnWidth = 8
    Range("B4").Select
    Sheets("CR").Select
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets("All SS-BB").Select
    Range("E2:G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "MAT  profiles"
    Range("H2:K2").Select
    ActiveCell.FormulaR1C1 = "Current Year"
    Range("Q2:AB2").Select
    ActiveCell.FormulaR1C1 = "Monthly Data"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "( / Sys / Yr )"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "( / Sys / Yr )"
    Sheets("CR").Select
    Range("A4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    lstAdd = ActiveCell.Offset(1, -1).Address

    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Sheets("All SS-BB").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("MTTR").Select
    Range("A4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    lstAdd = ActiveCell.Offset(1, -1).Address

    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Sheets("All SS-BB").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("ETTR").Select
    Range("A4").Select
    fstAdd = ActiveCell.Address
    ActiveSheet.UsedRange.Find(what:="Trend", lookat:=xlWhole).Select
    lstAdd = ActiveCell.Offset(1, -1).Address

    Range(fstAdd, lstAdd).Select
    Selection.Copy
    Sheets("All SS-BB").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("AC4").Select
    Range("$AC$4").SparklineGroups.Add Type:=xlSparkLine, SourceData:="Q4:AB4"
    Range("AC5").Select
    Range("$AC$5").SparklineGroups.Add Type:=xlSparkLine, SourceData:="Q5:AB5"
    Range("AC6").Select
    Range("$AC$6").SparklineGroups.Add Type:=xlSparkLine, SourceData:="Q6:AB6"
    Range("$AC$4:$AC$6").Select
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
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "KPIs"
    Range("B3").Select
    Selection.Copy
    Range("A2:A3").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CR"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "MTTR"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "ETTR"
    Range("A4:A6").Select
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
    Selection.Font.Bold = True
    Range("D4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    
    Columns("A:AD").Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWindow.DisplayGridlines = False
    Range("P4:P6").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TRUE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("E2").Select
    
    Selection.UnMerge
    Range(ActiveCell.Address, ActiveCell.Offset(0, 1).Address).Select
    Selection.EntireColumn.Select
    Selection.Delete
    Range("D2").Select
    ActiveCell.Offset(0, 1).value = "MAT  profile"
Range("A1").Select
    ActiveCell.FormulaR1C1 = "CTS KPI Dashboard Summary for All SubSystems and Building Blocks"
    Range("A1:AA1").Select
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
    With Selection.Font
        .name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10066176
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("A1:AA6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.UsedRange.Find(what:="ETTR", lookat:=xlWhole).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.value = "RRR%"
    ActiveCell.Offset(0, 1).value = "All Systems"
    ActiveCell.Offset(0, 2).value = "All Blocks"

    ActiveWindow.Zoom = 85
    Application.ActiveWorkbook.Save

End Sub

Sub RRRDatainAllSSBB()
    
    Sheets("All SS-BB").Select
    Range("AD3").Select
    Set pvtTbl = ActiveCell.PivotTable
    pvtName = ActiveCell.PivotTable.name
   
    Set pvtTbl = ActiveSheet.PivotTables(pvtName)
    Set pf = pvtTbl.PivotFields("Part12NC")
    pf.Orientation = xlPageField
    pf.Position = 2
       
    ActiveSheet.PivotTables(pvtName).PivotFields("Period").Orientation = _
        xlHidden
    
    Set pf = pvtTbl.PivotFields("Period")
    pf.Orientation = xlPageField
    pf.Position = 3
    
    'ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC").Orientation = _
        xlHidden
    'ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Total Calls (#)"), "# of Calls", xlSum
    ActiveSheet.PivotTables(pvtName).AddDataField ActiveSheet.PivotTables( _
        pvtName).PivotFields("Calls = 0 Visit"), "Sum of Calls = 0 Visit", xlSum
    ActiveSheet.PivotTables(pvtName).PivotFields("Sum of Calls = 0 Visit"). _
        Caption = "#Calls = 0 Visit"

'Add RRR% calculated field in Pivot table

        pvtTbl.CalculatedFields.Add "RRR", _
        "='Calls = 0 Visit' /'Total Calls (#)' *100", True
        pvtTbl.PivotFields("RRR").Orientation = _
        xlDataField

   
    With ActiveSheet.PivotTables(pvtName)
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
    End With

    Set pvtTbl = Worksheets("All SS-BB").PivotTables(pvtName)
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC"). _
        EnableMultiplePageItems = True
    'pvtTbl.PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionEndsWith, Value1:="Parts", Value2:="All"

    'ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC").PivotFilters.Add Type:=xlCaptionContains, Value1:="Parts", Value2:="Aggregatged"
    
    Dim pvtItm As PivotItem
    Set pvtTbl = ActiveSheet.PivotTables(pvtName)
    Set pf = pvtTbl.PivotFields("Part12NC")
   ' Set pvtItm = pf.PivotItem
    ActiveSheet.PivotTables(pvtName).PivotFields("Part12NC").ClearAllFilters

    For Each pvtItm In pvtTbl.PivotFields("Part12NC").PivotItems
        If pvtItm.name = "All Aggregated" Or pvtItm.name = "Non-Parts" Or pvtItm.name = "Parts" Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
 
    Next
    With ActiveSheet.PivotTables(pvtName)
        .ColumnGrand = True
        .RowGrand = True
    End With
    
    Set pf = pvtTbl.PivotFields("BuildingBlock")
    'pf.ClearAllFilters
    ActiveSheet.PivotTables(pvtName).PivotFields("BuildingBlock").ClearAllFilters
    
    For Each pvtItm In pvtTbl.PivotFields("BuildingBlock").PivotItems
        If pvtItm = "Buildingblocks Aggregated" Then
            pvtItm.Visible = True
        Else
            pvtItm.Visible = False
        End If
    Next pvtItm
    
    ActiveSheet.PivotTables(pvtName).PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables(pvtName).PivotFields("System"). _
        EnableMultiplePageItems = False
    ActiveSheet.PivotTables(pvtName).PivotFields("System").CurrentPage = "(All)"
    ActiveSheet.PivotTables(pvtName).PivotFields("System").ClearAllFilters
    ActiveSheet.PivotTables(pvtName).PivotFields("System").CurrentPage = _
        "System level"
   ' pvtTbl.RefreshTable

    fixedDate = Sheet1.combYear.value
    startDate = Mid(fixedDate, 1, 4) & " - " & Mid(fixedDate, 5, 2) & "-" & "01"
    startDate = Format(fixedDate, "yyyy" & "-" & "mm")
    endDate1 = Format(DateAdd("yyyy", -1, startDate), "yyyy" & "-" & "mm")
    endDate = Format(DateAdd("m", 1, endDate1), "yyyy" & "-" & "mm")
    Sheets("CR").Select
    Set pvtTbl = ActiveSheet.PivotTables(pvtName)
    Set pf = pvtTbl.PivotFields("Period")
    pf.ClearAllFilters
    pf.CurrentPage = "(All)"
    Cells(3, 18).Select
    i = 18
       
    Sheets("All SS-BB").Select
    
    
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
        Else
            pvtMonth = Format(pvtItm, "m" & "/" & "d" & "/" & "yyyy")
            Sheets("CR").UsedRange.Find(what:=pvtMonth, lookat:=xlWhole).Select
            ActiveCell.Offset(2, 0).Select
            myRow = ActiveCell.Row
            myCol = ActiveCell.Column
            pf.CurrentPage = pvtItm.Caption
            abc = pf.CurrentPage
            lr = Worksheets("CR").Cells(rows.Count, "D").End(xlUp).Row
            Range("AE3").Select
            fstAdd = ActiveCell.Address(True, True)
            Sheets("CR").UsedRange.Find(what:="#Calls (#)", lookat:=xlWhole).Select
            ActiveCell.Offset(1, -1).Select
            ActiveCell.End(xlDown).Select
'            ActiveCell.Offset(0, 6).Select
            lstAdd = ActiveCell.Address(True, True)
            rng = Range(fstAdd, lstAdd)
            Range("D3").Select
            vlkupAFstAdd = ActiveCell.Address(False, False)
            ActiveCell.End(xlDown).Select
            vlkupALstAdd = ActiveCell.Address(True, True)
            '#Calls (#)
    
            If i <= 29 Then
                For X = myRow To lr
                    On Error Resume Next
                    Cells(X, myCol).Formula = "=IFERROR(VLOOKUP(" & vlkupAFstAdd & ":" & vlkupALstAdd & "," & fstAdd & ":" & lstAdd & ",6,0)" & " / " & "(" & IBVal & "/" & 12 & ")" & "," & 0 & ")"
                    'Cells(X, MyCol).value = (Application.WorksheetFunction.VLookup(Cells(X, 4).value, rng, 6, False) / (IBVal / 12))
                    Cells(X, myCol).Copy
                    Cells(X, myCol).PasteSpecial xlValues
                    Cells(X, myCol).NumberFormat = "0.0000"
                Next X
             
            End If
                i = i + 1
        
        End If
      
    Next pvtItm

    pf.ClearAllFilters
    For Each pvtItm In pvtTbl.PivotFields("Period").PivotItems
        If pvtItm < endDate Or pvtItm > startDate Then
            pvtItm.Visible = False
        Else
            pvtItm.Visible = True
        End If
    Next pvtItm
            
        ActiveSheet.UsedRange.Find(what:="Grand Total", lookat:=xlWhole).Select

End Sub



