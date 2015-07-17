Attribute VB_Name = "UI"
'Date           Who     What

Option Explicit


Public Sub Automation_Tool_File_Open()
On Error Resume Next

Sheet1.rdbLocalDrive.value = 0
Sheet1.rdbSharedDrive.value = 0


Sheet1.combProductGroup.Clear
With Sheet1.combProductGroup
.AddItem ("IXR-MOS Endura-Y")
.AddItem ("IXR-MOS Pulsera-Y")
.AddItem ("IXR-MOS BV Vectra-N")
.AddItem ("IXR-MOS Veradius-Y")
.AddItem ("IXR-CV Allura FC-Y")
.AddItem ("IXR-MOS Libra-N")
.AddItem ("DXR-PrimaryDiagnost Digital-N")
.AddItem ("DXR-MicroDose Mammography-Y")
.AddItem ("DXR-MobileDiagnost Opta-N")
End With

Sheet1.comb6NC1.Clear
With Sheet1.comb6NC1
.AddItem ("IXR-MOS Endura-Y")
.AddItem ("IXR-MOS Pulsera-Y")
.AddItem ("IXR-MOS BV Vectra-N")
.AddItem ("IXR-MOS Veradius-Y")
.AddItem ("IXR-CV Allura FC-Y")
.AddItem ("IXR-MOS Libra-N")
.AddItem ("DXR-PrimaryDiagnost Digital-N")
.AddItem ("DXR-MicroDose Mammography-Y")
.AddItem ("DXR-MobileDiagnost Opta-N")
End With

Sheet1.rdbLocalDrive.value = True
Sheet1.processTime.value = 0
Sheet1.processTime.Enabled = False
Sheet1.minProcessTime.Enabled = False
Sheet1.combProductGroup.value = "Select Product Group"
Sheet1.comb6NC1.value = "Select Product Group"

Sheet1.chkRevenue.Enabled = False
Sheet1.chkAllGroups.value = False
Sheet1.minProcessTime.value = 0
Sheet1.processTime.value = 0

Dim yearValue As String
Dim monthVal As String
monthVal = Format(Now(), "mm")

yearValue = Year(Now()) & "-" & Format$(Now() - 31, "mm")
Sheet1.combYear.value = yearValue

End Sub

Public Sub Increase_Year_And_Month_Value()

Dim valYear As String, yearVal As String, monthVal As String
valYear = Sheet1.combYear.value
yearVal = Split(valYear, "-")(LBound(Split(valYear, "-")))
monthVal = Split(valYear, "-")(UBound(Split(valYear, "-")))

Select Case monthVal

Case "01"
valYear = yearVal & "-" & "02"
Case "02"
valYear = yearVal & "-" & "03"
Case "03"
valYear = yearVal & "-" & "04"
Case "04"
valYear = yearVal & "-" & "05"
Case "05"
valYear = yearVal & "-" & "06"
Case "06"
valYear = yearVal & "-" & "07"
Case "07"
valYear = yearVal & "-" & "08"
Case "08"
valYear = yearVal & "-" & "09"
Case "09"
valYear = yearVal & "-" & "10"
Case "10"
valYear = yearVal & "-" & "11"
Case "11"
valYear = yearVal & "-" & "12"
Case "12"
valYear = yearVal + 1 & "-" & "01"

End Select

Sheet1.combYear.value = valYear
End Sub

Public Sub Decrease_Year_And_Month_Value()

Dim valYear As String, yearVal As String, monthVal As String
valYear = Sheet1.combYear.value
yearVal = Split(valYear, "-")(LBound(Split(valYear, "-")))
monthVal = Split(valYear, "-")(UBound(Split(valYear, "-")))

Select Case monthVal

Case "01"
valYear = yearVal - 1 & "-" & "12"
Case "02"
valYear = yearVal & "-" & "01"
Case "03"
valYear = yearVal & "-" & "02"
Case "04"
valYear = yearVal & "-" & "03"
Case "05"
valYear = yearVal & "-" & "04"
Case "06"
valYear = yearVal & "-" & "05"
Case "07"
valYear = yearVal & "-" & "06"
Case "08"
valYear = yearVal & "-" & "07"
Case "09"
valYear = yearVal & "-" & "08"
Case "10"
valYear = yearVal & "-" & "09"
Case "11"
valYear = yearVal & "-" & "10"
Case "12"
valYear = yearVal & "-" & "11"

End Select

Sheet1.combYear.value = valYear

End Sub

