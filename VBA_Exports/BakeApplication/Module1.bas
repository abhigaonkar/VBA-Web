Attribute VB_Name = "Module1"
Option Explicit

Private objSession As clsSAPSession
Private objStatus As clsStatus
Public printerName As String
Public Const printYes As Boolean = False
Public Const productionMode As Boolean = False
'Public Const filePath As String = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\OvenLog.xlsx"
Public Const filePath As String = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\OvenLog.xlsx"
Public localSheet As Worksheet
Public onlineMode As Boolean

Function IsWorkBookOpen(filename As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next 'Ignore error if the file is not open
    Set wb = Workbooks(filename)
    On Error GoTo 0 'Resume normal error handling
    IsWorkBookOpen = Not wb Is Nothing 'Return True if the workbook is open
End Function
Sub ketchup()
    Dim startRow As Integer
    Dim s As Integer
    Dim nextPalletCard As String
    Dim thisPalletCard As String
    Dim keepGoing As Boolean
    keepGoing = True
    Set localSheet = ThisWorkbook.Sheets("Sheet1")
    startRow = InputBox("What row would you like to begin with?")
    s = startRow
    Do While keepGoing
        If localSheet.Cells(s, 10).Value = False Then
            thisPalletCard = localSheet.Cells(s, 3).Value
            saveToSAP s
            s = s + 1
            nextPalletCard = localSheet.Cells(s, 3).Value
            If nextPalletCard <> thisPalletCard Then
                doHumo startRow, s - 1
            End If
        Else
            keepGoing = False
        End If
    Loop
    keepGoing = True
End Sub
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

Sub LogEvent(message As String, Optional messageType As String = "Transaction")
    Dim csvFilePath As String
    ' Use the correct path separator
    If Not productionMode Then
        csvFilePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Furnace Deck\TransactionLog.csv"
    Else
        csvFilePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\TransactionLog.csv"
    End If
        
    ' Check if the file exists
    Dim fileNum As Integer
    fileNum = FreeFile()
    On Error GoTo ErrorHandler ' Add error handling
    Open csvFilePath For Append As #fileNum
    
    Print #fileNum, Format(Date, "mm/dd/yyyy") & "," & Time() & "," & Environ("computername") & ",Draw Application," & Environ$("UserName") & "," & messageType & "," & message
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.description
End Sub
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
            If Not left(xSession.ActiveWindow.Text, 15) = "SAP Easy Access" Then
                xSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
            End If
            x = x + 1
        Loop
        If left(xSession.ActiveWindow.Text, 15) = "SAP Easy Access" Then setStartScreen = True
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
        Application.Quit
        Exit Function
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in setStartScreen of Module1"
    
End Function

Sub saveRemoteData(ByVal startRowNum As Long, ByVal endRowNum As Long)
    Dim col As Integer
    Dim rw As Integer
    Dim thisWB As Workbook
    Dim remoteWB As Workbook
    
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
On Error GoTo ErrHandler
                ' --- Open the external workbook ---
        If IsWorkBookOpen(filePath) Then ' Check if the workbook is already open
            'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
        Else
            Workbooks.Open filename:=filePath, ReadOnly:=False, ignorereadonlyrecommended:=True  ' Open the workbook if it's not already open
        End If
    
        
        Set thisWB = ThisWorkbook
        Set thisSheet = thisWB.Sheets("Sheet1")
        
        Set remoteWB = Workbooks("OvenLog.xlsx")
        Set remoteSheet = remoteWB.Sheets("Sheet1")
        For rw = startRowNum To endRowNum
            For col = 1 To 38
                remoteSheet.Cells(rw, col).Value = thisSheet.Cells(rw, col)
            Next col
        Next rw
        Set thisWB = Nothing
        Set thisSheet = Nothing
        
        remoteWB.Close SaveChanges:=True
        Set remoteWB = Nothing
        Set remoteSheet = Nothing

    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteData of Module 1"
End Sub

