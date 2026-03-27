Attribute VB_Name = "Module1"
Option Explicit

Private objSession As clsSAPSession
Private objStatus As clsStatus
Public localSheet As Worksheet
Public onlineMode As Boolean
Public Const printerName As String = "HEBPRNP057"
Public Const printYes As Boolean = False
Public Const productionMode As Boolean = False

Public Const laborTime As Variant = 0.75

Function IsWorkBookOpen(ByVal filename As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next 'Ignore error if the file is not open
    Set wb = Workbooks(filename)
    On Error GoTo 0 'Resume normal error handling
    IsWorkBookOpen = Not wb Is Nothing 'Return True if the workbook is open
End Function
Sub ketchup()
    Dim startRow As Integer
    Dim s As Integer
    Dim keepGoing As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook ' Set the workbook object
    Set localSheet = wb.Sheets("Sheet1") ' Set the worksheet object

    keepGoing = True
    
    startRow = InputBox("What row would you like to begin with?")
    s = startRow
    Do While keepGoing
        If localSheet.Cells(s, 6).Value = False Then
            saveToSAP s
            s = s + 1
        Else
            keepGoing = False
        End If
    Loop
End Sub

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
' This function finds the next empty row in a specified column of a worksheet.
'
' Args:
'   shet: The worksheet to search in.
'   col: The column letter (e.g., "A", "B", "C") to search in.
'   rowStart: (Optional) The row number to start the search from.
'             Defaults to 6 if not provided.
'
' Returns:
'   The row number of the first empty cell in the specified column.
Function nextEmptyRow(shet As Worksheet, col As String, Optional rowStart As Long = 6) As Long

  On Error GoTo ErrHandler ' Error handling

  ' Initialize the starting row
  Dim x As Long
  x = rowStart

  ' Loop through the cells in the specified column
  Do While Not IsEmpty(shet.Cells(x, col).Value)
    x = x + 1
  Loop

  ' Return the row number of the first empty cell
  nextEmptyRow = x

  Exit Function ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextEmptyRow Function of Module1"
End Function
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
        Application.Quit
        Exit Function
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in setStartScreen of Module1"
    
End Function

Sub saveRemoteData(ByVal startRowNum As Long, ByVal endRowNum As Long)
 '   Dim filePath As String
    Dim col As Integer
    Dim rw As Integer
    Dim thisWB As Workbook
    Dim remoteWB As Workbook
    Dim remoteFilePath As String
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
On Error GoTo ErrHandler
        If productionMode Then
            remoteFilePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\HVPackLog.xlsx"
        Else
            remoteFilePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\HVPackLog.xlsx"
        End If
                ' --- Open the external workbook ---
        If IsWorkBookOpen(remoteFilePath) Then ' Check if the workbook is already open
            'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
        Else
            Workbooks.Open filename:=remoteFilePath   ' Open the workbook if it's not already open
        End If
    
        
        Set thisWB = ThisWorkbook
        Set thisSheet = thisWB.Sheets("Sheet1")
        
        Set remoteWB = Workbooks("HVPackLog.xlsx")
        Set remoteSheet = remoteWB.Sheets("Sheet1")
        For rw = startRowNum To endRowNum
            For col = 1 To 32
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

