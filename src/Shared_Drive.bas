Attribute VB_Name = "Shared_Drive"
Option Explicit
Public sharedDrivePath As String
Public inputFileName As String
Public fileExists As Boolean
Public fd As FileDialog

Public Function SharedDrive_Path(inputFileName As String)
    On Error Resume Next
    
    
    Dim FSO, ofolder, osubfolder, ofile, queue As Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    
If fileExists = False Then 'Open filedialog if file not present
'GoTo fileNotPresent
End If

Set fd = Application.FileDialog(msoFileDialogFolderPicker)
fileExists = False
    If Application.FileDialog(msoFileDialogFolderPicker).SelectedItems.Count = 0 Then
fileNotPresent:
        If Application.FileDialog(msoFileDialogFolderPicker).Show <> -1 Then
        MsgBox "No Folder Selected"
        End
        End If
    End If
    
    queue.Add FSO.GetFolder(Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1))

    Do While queue.Count > 0
        Set ofolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        For Each osubfolder In ofolder.SubFolders
            queue.Add osubfolder 'enqueue
        Next osubfolder
        
        For Each ofile In ofolder.Files
            If inputFileName = ofile.name Then
                sharedDrivePath = ofile.Path
                fileExists = True
            End If
        Next ofile
    Loop
Set fd = Nothing
End Function
