Attribute VB_Name = "CreatePresentation"
Option Explicit
Dim myworkbook As String

Sub Create_Presentation()
Dim ws As Worksheet
Dim wb As FileDialog
Dim FSO, oFolder, osubfolder, ofile, queue As Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
Dim slicerItem As slicerItem

If Application.FileDialog(msoFileDialogFilePicker).Show <> -1 Then
    MsgBox "User Canceled!", vbExclamation, "Creating Presentation"
    End
End If

ofile = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)

Application.Workbooks.Open (ofile)
myworkbook = ActiveWorkbook.name

ExportChartsToPPT myworkbook

'Workbooks(myworkbook).Close False

End Sub


Function getPPPres() As Object
Dim PPApp As Object

'Reference instance of PowerPoint
On Error Resume Next
'Check whether PowerPoint is running
Set PPApp = GetObject(, "PowerPoint.Application")
If PPApp Is Nothing Then
'PowerPoint is not running, create new instance
Set PPApp = CreateObject("PowerPoint.Application")
'For automation to work, PowerPoint must be visible
PPApp.Visible = True
End If
On Error GoTo 0

'Reference presentation and slide
On Error Resume Next
If PPApp.Windows.Count > 0 Then
'There is at least one presentation
'Use existing presentation
Set getPPPres = PPApp.ActivePresentation
Else
'There are no presentations
'Create New Presentation
Set getPPPres = PPApp.Presentations.Add
End If
Set PPApp = Nothing
End Function

Function getNewSlide(PPPres As Object) As Object
Set getNewSlide = PPPres.Slides.Add(PPPres.Slides.Count + 1, 12)
End Function

Sub ExportChartsToPPT(myBook As String)
Dim PPPres As Object
Dim PPSlide As Object
Dim cht As ChartObject
Dim wksChartsFromSheet As Worksheet

Application.Workbooks(myBook).Activate

For Each wksChartsFromSheet In ActiveWorkbook.Sheets
If wksChartsFromSheet.ChartObjects.Count = 0 Then
'MsgBox "No Chart to Export to Powerpoint", vbInformation, ""
Exit Sub
End If

Set PPPres = getPPPres

' If PPPres.Slides.Count = 0 Then
' Set PPSlide = getNewSlide(PPPres)
' End If

For Each cht In wksChartsFromSheet.ChartObjects

Set PPSlide = getNewSlide(PPPres)
cht.CopyPicture
PPSlide.Select
PPSlide.Shapes.Paste
PPSlide.Application.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
PPSlide.Application.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
PPSlide.Select
Next cht

Next wksChartsFromSheet

Set cht = Nothing
Set PPSlide = Nothing
Set PPPres = Nothing

End Sub

