Attribute VB_Name = "WriteCode"
Option Explicit

' this VBA project requires
' 1 - references to Microsoft Visual Basic For Applications Extensibility 5.3
' add it via Tools > References
'
' 2 - trust access to VBA project object model
' In spreadsheet view go to Excel(application options) >> Trust Centre >> Macro Settings
' tick the Trust Access to VBA project object model

Sub Main()
    Dim arrayNames As Variant, v As Variant

    ' create an array which will hold the module names
    arrayNames = Array("array_one", "array_two", "array_three")

    ' add a new module to the VBA Project and assign
    ' the module's name to the moduleName variable
    Dim moduleName As String
    moduleName = addModule

    ' iterate over the array with names
    For Each v In arrayNames
    ' write a new line in the module for each name in the array
        WriteToModule moduleName, CStr(v)
    Next

    ' adds an End Sub to the module
    FinishWritingModule
End Sub

Private Function addModule() As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule

    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    Set CodeMod = VBComp.CodeModule

    With CodeMod
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, "Public Sub DynamicallyCreatedArrays()"
        .InsertLines 2, " ' code for the sub"
    End With

    addModule = VBComp.name
End Function

Private Sub WriteToModule(moduleName As String, arrayName As String)
    With ActiveWorkbook.VBProject.VBComponents(moduleName).CodeModule
        .InsertLines .CountOfLines + 2, " Dim " & arrayName & " as Variant"
    End With
End Sub

Private Sub FinishWritingModule()
    With ActiveWorkbook.VBProject.VBComponents("Module2").CodeModule
        .InsertLines .CountOfLines + 2, "End Sub"
    End With
End Sub
