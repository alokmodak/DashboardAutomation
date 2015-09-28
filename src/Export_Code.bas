Attribute VB_Name = "Export_Code"
'* Export Code incorporated by Jitendra Deshmukh Dated 6/30/2015

'Date           Who     What



Public Sub exportVbaCode()

Dim objUserEnvVars As Object
Dim strVar As String
Set objUserEnvVars = CreateObject("WScript.Shell").Environment("User")
strVar = objUserEnvVars.Item("Dashboard_Automation")
If Not InStr(strVar, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Sub
    End If
    
Dim vbaProject As VBProject

Set vbaProject = ThisWorkbook.VBProject

    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=True)

    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) Then
            'Debug.Print "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, ".bas"
                Case vbext_ct_MSForm
                    exportComponent export_path, component, ".frm"
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        End If
    Next component
End Sub


Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.CodeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.CodeModule.Lines(1, 1))
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function


'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    Debug.Print "exporting " & component.name & extension
    component.Export exportPath & "\" & component.name & extension
End Sub


'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    Debug.Print "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim FSO As New scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    outStream.Write (component.CodeModule.Lines(1, component.CodeModule.CountOfLines))
    outStream.Close
End Sub

Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function


' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function


' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function

Public Function getSourceDir(fullWorkbookPath As String, createIfNotExists As Boolean) As String
    Dim objUserEnvVars As Object
Dim strVar As String
Set objUserEnvVars = CreateObject("WScript.Shell").Environment("User")
strVar = objUserEnvVars.Item("Dashboard_Automation")

    ' First check if the fullWorkbookPath contains a \.
    If Not InStr(strVar, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim FSO As New scripting.FileSystemObject
    Dim projDir As String
    projDir = strVar
    
    Dim srcDir As String
    srcDir = projDir & "\src\"
    Dim exportDir As String
    exportDir = srcDir

    If createIfNotExists Then
        If Not FSO.FolderExists(srcDir) Then
            FSO.CreateFolder srcDir
            Debug.Print "Created Folder " & srcDir
        End If
        If Not FSO.FolderExists(exportDir) Then
            FSO.CreateFolder exportDir
            Debug.Print "Created Folder " & exportDir
        End If
    Else
        
    End If
    getSourceDir = srcDir
End Function



