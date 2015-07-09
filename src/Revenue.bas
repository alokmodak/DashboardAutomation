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
Dim j As Integer
Dim zcswVal As Boolean
Dim countFstAddress As String
Dim countLstAddress As String

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

ActiveWorkbook.Sheets("Endura").Activate
ActiveSheet.Cells(2, 27).Select
Dim fstTableAdd As String
fstTableAdd = ActiveCell.Address
ActiveCell.End(xlToRight).Select

monthsForTable = DateAdd("m", -24, Date)

ActiveCell.Offset(0, 1).Select
For monthCellForTable = 2 To 37
    ActiveCell.value = monthsForTable
    ActiveCell.NumberFormat = "[$-409]mmm-yy;@"
        If monthCellForTable > 2 Then
            ActiveCell.Offset(0, 3).Select
            ActiveCell.Offset(0, 1).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Joined"
            ActiveCell.Offset(0, 2).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Dropped"
        Else
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Offset(0, 1).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Joined"
            ActiveCell.Offset(0, 2).value = Format(DateAdd("m", 1, monthsForTable), "mmmyy") & "-" & "Dropped"
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
If ActiveCell.value <> "" Then
            'leave row values blank if start or end date is not available
            If ActiveCell.Offset(0, 1).value = "" Then
                ActiveCell.Offset(0, 1).value = ActiveCell.Offset(-1, 1).value
            End If
            If ActiveCell.Offset(0, 2).value = "" Then
                ActiveCell.Offset(0, 2).value = ActiveCell.Offset(-2, 2).value
            End If
            duration = DateDiff("m", Replace(ActiveCell.Offset(0, 1).value, ".", "/"), Replace(ActiveCell.Offset(0, 2).value, ".", "/"))
            i = 1
            Do Until ActiveCell.Offset(i, 0).value <> ""
            'exit loop for last cell
                If ActiveCell.Offset(i, 3).value = "" Then
                Exit Do
                End If
            If ActiveCell.Offset(i, 1).value = "" Then
                ActiveCell.Offset(i, 1).value = ActiveCell.Offset(-1, 1).value
            End If
            If ActiveCell.Offset(i, 2).value = "" Then
                ActiveCell.Offset(i, 2).value = ActiveCell.Offset(-2, 2).value
            End If
            duration = duration + DateDiff("m", Replace(ActiveCell.Offset(i, 1).value, ".", "/"), Replace(ActiveCell.Offset(i, 2).value, ".", "/"))
            i = i + 1
            Loop
        
            monthCellForTable = 4
            For i = 1 To 36
            
        Dim k As Integer
        k = 0
        Do
        'exit for last cell
        If ActiveCell.Offset(k, 3).value = "" Then
            Exit Do
        End If
                fstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 1).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 1).value, ".", "/", 4)), 1)
                lstVal = DateSerial(Year(Replace(ActiveCell.Offset(k, 2).value, ".", "/", 4)), Month(Replace(ActiveCell.Offset(k, 2).value, ".", "/", 4)) + 1, 0)
                
                If fstVal <= CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) And CDate(ActiveCell.Offset(-topCelVal, monthCellForTable).value) <= lstVal Then
                    ActiveCell.Offset(0, monthCellForTable).value = "Yes"
                Else
                    ActiveCell.Offset(0, monthCellForTable).value = "No"
                End If
        k = k + 1
        Loop Until ActiveCell.Offset(k, 0).value <> ""

If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then
    If ActiveCell.Offset(0, monthCellForTable - 1).value = "Yes" Then
        If duration <= 12 Then
            ActiveCell.Offset(0, monthCellForTable + 2).value = "0To1Year"
        ElseIf 13 >= duration <= 36 Then
            ActiveCell.Offset(0, monthCellForTable + 2).value = "2To3Years"
        ElseIf 37 >= duration <= 60 Then
            ActiveCell.Offset(0, monthCellForTable + 2).value = "3To5Years"
        ElseIf duration >= 61 Then
            ActiveCell.Offset(0, monthCellForTable + 2).value = "MoreThan5Years"
        End If

    'condition for After warranty
    If ActiveCell.Offset(0, 3).value = "ZCSW" Then
    j = 1
    zcswVal = True
    Do Until ActiveCell.Offset(j, 0) <> ""
    'condition for last row exit loop
        If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
            If ActiveCell.Offset(1, 3).value = "" Then
                Exit Do
        End If
        zcswVal = False
    End If
    j = j + 1
    Loop
    If zcswVal = True Then
        ActiveCell.Offset(0, monthCellForTable + 2).value = "AfterWarranty"
    End If
End If

End If
End If
If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "No" Then

If ActiveCell.Offset(0, monthCellForTable - 3).value = "Yes" Then
If duration <= 12 Then
ActiveCell.Offset(0, monthCellForTable + 2).value = "0To1Year"
ElseIf 13 >= duration <= 36 Then
ActiveCell.Offset(0, monthCellForTable + 2).value = "2To3Years"
ElseIf 37 >= duration <= 60 Then
ActiveCell.Offset(0, monthCellForTable + 2).value = "3To5Years"
ElseIf duration >= 61 Then
ActiveCell.Offset(0, monthCellForTable + 2).value = "MoreThan5Years"
End If

If ActiveCell.Offset(0, 3).value = "ZCSW" Then
j = 1
zcswVal = True
Do Until ActiveCell.Offset(j, 0) <> ""
'condition for last row exit loop
If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
If ActiveCell.Offset(1, 3).value = "" Then
Exit Do
End If
zcswVal = False
End If
j = j + 1
Loop
If zcswVal = True Then
ActiveCell.Offset(0, monthCellForTable + 2).value = "AfterWarranty"
End If
End If

End If
End If


If i = 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
  If ActiveCell.Offset(0, monthCellForTable - 1).value = "No" Then
   If duration <= 12 Then
     ActiveCell.Offset(0, monthCellForTable + 1).value = "0To1Year"
   ElseIf 13 >= duration <= 36 Then
     ActiveCell.Offset(0, monthCellForTable + 1).value = "2To3Years"
   ElseIf 37 >= duration <= 60 Then
     ActiveCell.Offset(0, monthCellForTable + 1).value = "3To5Years"
   ElseIf duration >= 61 Then
     ActiveCell.Offset(0, monthCellForTable + 1).value = "MoreThan5Years"
   End If
'condition for After warranty
If ActiveCell.Offset(0, 3).value = "ZCSW" Then
   j = 1
zcswVal = True
Do Until ActiveCell.Offset(j, 0) <> ""
                    'condition for last row exit loop
                    If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                        If ActiveCell.Offset(1, 3).value = "" Then
                            Exit Do
                        End If
                        zcswVal = False
                    End If
                j = j + 1
                Loop
                If zcswVal = True Then
                    ActiveCell.Offset(0, monthCellForTable + 1).value = "AfterWarranty"
                End If
        End If

    End If
End If
                    If i > 2 And ActiveCell.Offset(0, monthCellForTable).value = "Yes" Then
                        
                        If ActiveCell.Offset(0, monthCellForTable - 3).value = "No" Then
                            If duration <= 12 Then
                                ActiveCell.Offset(0, monthCellForTable + 1).value = "0To1Year"
                            ElseIf 13 >= duration <= 36 Then
                                ActiveCell.Offset(0, monthCellForTable + 1).value = "2To3Years"
                            ElseIf 37 >= duration <= 60 Then
                                ActiveCell.Offset(0, monthCellForTable + 1).value = "3To5Years"
                            ElseIf duration >= 61 Then
                                ActiveCell.Offset(0, monthCellForTable + 1).value = "MoreThan5Years"
                            End If
                            
                            'condition for After warranty
                            If ActiveCell.Offset(0, 3).value = "ZCSW" Then
                                j = 1
                                zcswVal = True
                                    Do Until ActiveCell.Offset(j, 0) <> ""
                                        'condition for last row exit loop
                                        If ActiveCell.Offset(j, 3).value <> "ZCSW" Then
                                            If ActiveCell.Offset(1, 3).value = "" Then
                                                Exit Do
                                            End If
                                            zcswVal = False
                                        End If
                                    j = j + 1
                                    Loop
                                    If zcswVal = True Then
                                        ActiveCell.Offset(0, monthCellForTable + 1).value = "AfterWarranty"
                                    End If
                            End If
                        End If
                    End If

                If i > 1 Then
                    monthCellForTable = monthCellForTable + 3
                Else
                    monthCellForTable = monthCellForTable + 1
                End If
            Next
End If
topCelVal = topCelVal + 1
ActiveCell.Offset(1, 0).Select
Next

'Calculating total numbers for dropped and up

ActiveCell.Offset(3, 3).Select
ActiveCell.value = "Total"
ActiveCell.Offset(1, 0).value = "0To1Year"
ActiveCell.Offset(2, 0).value = "2To3Years"
ActiveCell.Offset(3, 0).value = "3To5Years"
ActiveCell.Offset(4, 0).value = "MoreThan5Years"
ActiveCell.Offset(5, 0).value = "AfterWarranty"
ActiveCell.Offset(6, 0).value = "EOL"

Dim totalVal As Integer
totalVal = 0
'loop for counting totals
For i = 1 To 36
    If i >= 2 Then
        countLstAddress = ActiveCell.Offset(-3, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 2), 1).Address
        For Each cell In Range(countFstAddress, countLstAddress)
            If cell.value = "Yes" Then
            
            End If
        Next
        
        ActiveCell.value = Application.WorksheetFunction.CountA(Range(countFstAddress, countLstAddress))
        countLstAddress = ActiveCell.Offset(0, 1).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 2), 1).Address
        ActiveCell.Offset(0, 1).value = Application.WorksheetFunction.CountA(Range(countFstAddress, countLstAddress))
        countLstAddress = ActiveCell.Offset(0, 2).Address
        countFstAddress = ActiveCell.Offset(-(topCelVal + 2), 2).Address
        ActiveCell.Offset(0, 2).value = Application.WorksheetFunction.CountA(Range(countFstAddress, countLstAddress))
        ActiveCell.Offset(0, 3).Select
    Else
        ActiveCell.Offset(0, 1).Select
    End If
Next

ActiveWorkbook.Sheets("Pivot").delete
ActiveWorkbook.Sheets("Data").delete
End Sub
