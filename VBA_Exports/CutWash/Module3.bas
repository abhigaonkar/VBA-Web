Attribute VB_Name = "Module3"
Sub ExportAllModules()
    Dim vbComponent As Object
    Dim exportPath As String
    Dim fileName As String
    Dim fileExt As String

    ' Set export folder path (change this to your desired folder)
    exportPath = "C:\Users\abhishek.gaonkar\VBA_Exports\CutWash\"
    
    ' Ensure path ends with backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Create folder if it doesn't exist
    On Error Resume Next
    MkDir exportPath
    On Error GoTo 0

    ' Export each component
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        Select Case vbComponent.Type
            Case 1: fileExt = ".bas" ' Standard module
            Case 2: fileExt = ".cls" ' Class module
            Case 3: fileExt = ".frm" ' UserForm
            Case Else: fileExt = ""  ' Sheet/Workbook, skip them
        End Select
        If fileExt <> "" Then
            fileName = exportPath & vbComponent.Name & fileExt
            vbComponent.Export fileName
        End If
    Next vbComponent

    MsgBox "All modules exported to: " & exportPath
End Sub
