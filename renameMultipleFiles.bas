Attribute VB_Name = "renameMultipleFiles"
Sub listFileNameFromFolder()
 
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
 
Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sFolder)
 
For Each oFile In oFolder.Files
    Cells(i + 1, 1) = oFile.Name 'list all the file name at column a
    i = i + 1
Next oFile
 
End Sub

Sub renameMultipleFiles()
    With Application.FileDialog(msoFileDialogFolderPicker)     ' Open the select folder prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            selectDirectory = .SelectedItems(1)
            dFileList = Dir(selectDirectory & Application.PathSeparator & "*")
        
            Do Until dFileList = "" 'Replace orinigal file from column B with new file name in column D
                curRow = 0
                On Error Resume Next
                curRow = Application.Match(dFileList, Range("B:B"), 0)
                If curRow > 0 Then
                    Name selectDirectory & Application.PathSeparator & dFileList As _
                    selectDirectory & Application.PathSeparator & Cells(curRow, "D").Value
                End If
        
                dFileList = Dir
            Loop
        End If
    End With
End Sub

