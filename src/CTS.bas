Attribute VB_Name = "CTS"
Option Explicit

Private Sub calculateRRR()
Dim inputItem As String
Dim fstMonthChk As String
Dim myWorkBook As String
Dim outputFile As String
Dim lookFor As Range
Dim srchRange As Range
Dim book1 As Workbook
Dim book2 As Workbook
Dim lastClmnUsed As Integer
Dim j, i, row1, row2, col1 As Integer

On Error Resume Next
fstMonthChk = Format(Sheet1.combYear.value, "yyyy")

inputItem = ThisWorkbook.Path & "\" & Dir(ThisWorkbook.Path & "\" & "iXR_RR_" & fstMonthChk & "*.xls*") 'input file path
Application.Workbooks.Open (inputItem), False
 myWorkBook = ActiveWorkbook.name
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets(fstMonthChk).Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find("SWO").Select
    Selection.EntireColumn.Copy
    Workbooks.Add
    'save out put file
    outputFile = ActiveWorkbook.name
    Sheets("Sheet1").Activate
    ActiveSheet.Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets(fstMonthChk).Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find("CaseCount").Select
    Selection.EntireColumn.Copy
    Workbooks(outputFile).Activate
    Sheets("Sheet1").Activate
    ActiveSheet.Range("B1").Select
    ActiveSheet.Paste
    
    Workbooks(myWorkBook).Activate
    ActiveWorkbook.Sheets(fstMonthChk).Activate
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.UsedRange.Find("RemotelyResolved").Select
    Selection.EntireColumn.Copy
    Workbooks(outputFile).Activate
    Sheets("Sheet1").Activate
    ActiveSheet.Range("C1").Select
    ActiveSheet.Paste
    
    i = 4
    j = 2
    
    lastClmnUsed = Workbooks(myWorkBook).Sheets(2).Cells(1, Columns.Count).End(xlToLeft).Column
    
    Set srchRange = Workbooks(myWorkBook).Sheets(2).Range("B1:Q" & Range("B" & Rows.Count).End(xlUp).Row)
    Set lookFor = Workbooks(outputFile).Sheets(1).Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)

    Workbooks(outputFile).Activate

    For j = 2 To lastClmnUsed

    lookFor.Offset(0, i).value = Application.VLookup(lookFor, srchRange, j, False)
    i = i + 1
    'j = j + 1
    Next j
    If i > lastClmnUsed Then
    Exit For
    End If
End Sub
