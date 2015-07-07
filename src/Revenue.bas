Attribute VB_Name = "Revenue"
Option Explicit

Public Sub Revenue_Graph_Creation()

Application.AlertBeforeOverwriting = False
Application.DisplayAlerts = False
Application.Workbooks.Add
ActiveWorkbook.SaveAs fileName:="D:\Philips\Assignments\Revenue\ContractDynamics_Waterfall_Jul15.xlsx", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
Dim outputFile As String
outputFile = ActiveWorkbook.name

Dim inputRevenue As String
inputRevenue = "D:\Philips\Assignments\Revenue\ContractDynamics_Waterfall.xlsx"
Dim inputFileNameContracts As String
Application.Workbooks.Open (inputRevenue)
inputFileNameContracts = Split(inputRevenue, "\")(UBound(Split(inputRevenue, "\")))
ActiveWorkbook.Sheets("SAPBW_DOWNLOAD").Activate
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole).Select
ActiveSheet.UsedRange.Find(what:="[C,S] System Code Material (Material no of  R Eq)", LookAt:=xlWhole, after:=ActiveCell).Select
Dim fstAddForPivot As String
Dim lstAddForPivot As String
fstAddForPivot = ActiveCell.Address
ActiveCell.End(xlDown).Select
ActiveCell.End(xlToRight).Select
lstAddForPivot = ActiveCell.Address
ActiveSheet.Range(fstAddForPivot, lstAddForPivot).Select
Selection.Copy

Workbooks(outputFile).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Paste
ActiveSheet.name = "Data"

Application.Workbooks(inputFileNameContracts).Close False







End Sub
