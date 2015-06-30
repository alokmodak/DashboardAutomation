Attribute VB_Name = "Import_Code"

Option Explicit


Private Const IMPORT_DELAY As String = "00:00:02"

'We need to make these variables public such that they can be given as arguments to application.ontime()
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

Public Sub testImport()
    Dim proj_name As String
    proj_name = "Dashboard_Automation"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    MsgBox proj_name
    Import_Code.importVbaCode vbaProject
End Sub

Public Function getSourceDir(fullWorkbookPath As String, createIfNotExists As Boolean) As String
    ' First check if the fullWorkbookPath contains a \.
    If Not InStr(fullWorkbookPath, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim FSO As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = FSO.GetParentFolderName(fullWorkbookPath) & "\"
    Dim srcDir As String
    srcDir = projDir & "src\"
    Dim exportDir As String
    exportDir = srcDir & FSO.GetFileName(fullWorkbookPath) & "\"

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
        If Not FSO.FolderExists(exportDir) Then
            Debug.Print "Folder does not exist: " & exportDir
            exportDir = ""
        End If
    End If
    getSourceDir = exportDir
End Function

' Usually called after the given workbook is opened. The option includeClassFiles is False by default because
' they don't import correctly from VBA. They'll have to be imported manually instead.
Public Sub importVbaCode(vbaProject As VBProject, Optional includeClassFiles As Boolean = False)
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
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=False)
    If export_path = "" Then
        'The source directory does not exist, code has never been exported for this vbaProject.
        Debug.Print "No import directory for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    'initialize globals for Application.OnTime
    Set componentsToImport = New Dictionary
    Set sheetsToImport = New Dictionary
    Set vbaProjectToImport = vbaProject

    Dim FSO As New Scripting.FileSystemObject
    Dim projContents As Folder
    Set projContents = FSO.GetFolder(export_path)
    Dim file As Object
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file, includeClassFiles
    Next

    Dim componentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        removeComponent vbaProject, componentName
    Next
    'Then import them
    Debug.Print "Invoking 'Import_Code.importComponents'with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'Import_Code.importComponents'"
    Debug.Print "almost finished importing code for " & vbaProject.name
End Sub


Private Sub checkHowToImport(file As Object, includeClassFiles As Boolean)
    Dim fileName As String
    fileName = file.name
    Dim componentName As String
    componentName = Left(fileName, InStr(fileName, ".") - 1)
    If componentName = "Import_Code" Then
        '"don't remove or import ourself
        Exit Sub
    End If

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                    'import lines into sheet: importLines vbaProjectToImport, file
                    sheetsToImport.Add componentName, file
                Else
                    ' .cls files don't import correctly because of a bug in excel, therefore we can exclude them.
                    ' In that case they'll have to be imported manually.
                    If includeClassFiles Then
                        'importComponent vbaProject, file
                        componentsToImport.Add componentName, file.Path
                    End If
                End If
            Case ".bas", ".frm"
                'importComponent vbaProject, file
                componentsToImport.Add componentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub


' Only removes the vba component if it exists
Private Sub removeComponent(vbaProject As VBProject, componentName As String)
    If componentExists(vbaProject, componentName) Then
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(componentName)
        Debug.Print "removing " & c.name
        vbaProject.VBComponents.Remove c
    End If
End Sub


Public Sub importComponents()
    If componentsToImport Is Nothing Then
        Debug.Print "Failed to import! Dictionary 'componentsToImport' was not initialized."
        Exit Sub
    End If
    Dim componentName As String
    Dim vComponentName As Variant
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        importComponent vbaProjectToImport, componentsToImport(componentName)
    Next

    'Import the sheets
    For Each vComponentName In sheetsToImport.Keys
        componentName = vComponentName
        importLines vbaProjectToImport, sheetsToImport(componentName)
    Next

    Debug.Print "Finished importing code for " & vbaProjectToImport.name
    'We're done, clear globals explicitly to free memory.
    Set componentsToImport = Nothing
    Set vbaProjectToImport = Nothing
End Sub


' Assumes any component with same name has already been removed.
Private Sub importComponent(vbaProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    'This next line is a bug! It imports all classes as modules!
    vbaProject.VBComponents.Import filePath
End Sub


Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim componentName As String
    componentName = Left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(vbaProject, componentName) Then
        ' Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(componentName, vbaProject.fileName)
        Set c = vbaProject.VBComponents(addedSheetCodeName)
        c.name = componentName
    End If
    Set c = vbaProject.VBComponents(componentName)
    Debug.Print "Importing lines from " & componentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromFile file.Path
    On Error GoTo 0
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


