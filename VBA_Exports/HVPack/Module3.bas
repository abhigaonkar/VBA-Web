Attribute VB_Name = "Module3"
Sub ExportAllModules()
    Dim vbComponent As Object
    Dim exportPath As String
    
    ' Set export folder path (change this to your desired folder)
    exportPath = "C:\Users\abhishek.gaonkar\VBA_Exports\HVPack\" ' Create this folder first
    
    ' Create folder if it doesn't exist
    On Error Resume Next
    MkDir exportPath
    On Error GoTo 0
    
    ' Export each component
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        If vbComponent.Type <> 100 Then ' Skip document modules
            vbComponent.Export exportPath & vbComponent.Name & ".bas"
        End If
    Next vbComponent
    
    MsgBox "All modules exported to: " & exportPath
End Sub

