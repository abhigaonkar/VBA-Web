Attribute VB_Name = "Module1"
Option Explicit

Private objSession As clsSAPSession
Private objStatus As clsStatus
Public Const printerName As String = "HEBPRNP060"
Public Const remFile As String = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\FSCutWashLog.xlsx"
'Public Const remFile As String = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\FSCutWashLog.xlsx"
Public Const printMe As Boolean = True
Public Const productionMode As Boolean = True
Public onlineMode As Boolean
Public thisWB As Workbook
Public dataWS As Worksheet
Public settingsWS As Worksheet

Function SAPSession() As clsSAPSession
    On Error GoTo ErrHandler
    If objSession Is Nothing Then
        Set objSession = New clsSAPSession
    End If
    Set SAPSession = objSession
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in SAPSession of modSingleton"
End Function

Function addZeros(sourceNum As Integer) As String
'pads a an integer with 0
Dim thisLen As Long
thisLen = Len(CStr(sourceNum))
Select Case thisLen
    Case 1
        addZeros = "000" & CStr(sourceNum)
    Case 2
        addZeros = "00" & CStr(sourceNum)
    Case 3
        addZeros = "0" & CStr(sourceNum)
    Case 4
        addZeros = CStr(sourceNum)
End Select
End Function

Function SAPSessionIsLoaded() As Boolean
On Error GoTo ErrHandler
    If SAPSession.CurrentSession Is Nothing Then
        SAPSessionIsLoaded = False
    Else
        SAPSessionIsLoaded = True
    End If
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in SAPSessionIsLoaded of modSingleton"
End Function
Function SAPStatus() As clsStatus
On Error GoTo ErrHandler
    If objStatus Is Nothing Then
        Set objStatus = New clsStatus
    End If
    Set SAPStatus = objStatus
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in SAPStatus of modSingleton"
End Function

Function killSessionAndStatus() As Boolean
    If Not objStatus Is Nothing Then
        Set objStatus = Nothing
    End If
    If Not objSession Is Nothing Then
        Set objSession = Nothing
    End If
End Function

Public Sub PauseMe(secs As Variant, displayForm As Boolean)
    Dim xTim
    Dim p As Long
On Error GoTo ErrHandler
    xTim = Timer
    p = 1
    Do While xTim + secs > Timer
        If p = 14000 Then p = 0
        p = p + 1
    Loop
    Exit Sub
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in PauseMe of Module1"
    Resume
End Sub
Public Sub StdErrorHandler(ErrTxt As String)
On Error GoTo ErrHandler
    MsgBox ErrTxt
    If Err.Number = -21047023179# Then
        killSessionAndStatus
        Resume Next
        Exit Sub
    End If
    If MsgBox("Would you like to stop code?", vbYesNo) = vbNo Then
        Resume Next
    Else
        On Error GoTo 0
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err & ": " & Error(Err) & " in StdErrorHandler of Module1"
    Resume Next
End Sub
Function setStartScreen() As Boolean
    Dim x As Integer
    Dim xSession As GuiSession
    
On Error GoTo ErrHandler
    setStartScreen = False
    killSessionAndStatus
    If SAPSession.CurrentSession Is Nothing Then SAPSession.ConnectToExistingSession
    If SAPSessionIsLoaded Then
        Set xSession = SAPSession.CurrentSession
        x = 1
        setStartScreen = False
        Do While x < 5
            If Not Left(xSession.ActiveWindow.Text, 15) = "SAP Easy Access" Then
                xSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
            End If
            x = x + 1
        Loop
        If Left(xSession.ActiveWindow.Text, 15) = "SAP Easy Access" Then setStartScreen = True
        Set xSession = Nothing
    Else
        objSession.ConnectToExistingSession
        setStartScreen = False
    End If
    Exit Function
    
ErrHandler:
    If Err.Number = -2147417848 Then
        Err.Clear
        If x < 5 Then Resume Next
    ElseIf Err.Number = -2147023179 Then
        MsgBox "Please log in to SAP before starting this program.", vbOKOnly
        application.Quit
        Exit Function
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in setStartScreen of Module1"
    
End Function

