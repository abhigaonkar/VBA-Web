Attribute VB_Name = "Module1"
Option Explicit

Private objSession As clsSAPSession
Private objStatus As clsStatus
Public onlineMode As Boolean
Public furnName As String
Public equipName As String
Public Const printMe As Boolean = True
Public Const productionMode As Boolean = True
Public printerName As String
Public computerName As String
Public sSandQty As Single
Public localWB As Workbook
Public localWS As Worksheet

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
    If objSession.CurrentSession Is Nothing Then
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
   ' LogEvent ErrTxt, "Error"
    MsgBox ErrTxt
    If Err.Number = -21047023179# Then
        killSessionAndStatus
        Resume Next
        Exit Sub
    ElseIf Err.Number = 0 Then
        On Error GoTo 0
    End If
    If MsgBox("Would you like to stop code?", vbYesNo) = vbNo Then
        Resume Next
    Else
        On Error GoTo 0
        Exit Sub
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
Function nextEmptyRow(shet As Worksheet, ByVal col As Long, Optional ByVal rowStart As Long) As Long

  On Error GoTo ErrHandler ' Error handling

  ' Initialize the starting row
  Dim x As Long
  x = rowStart
  ' Loop through the cells in the specified column
'  Do While Not IsEmpty(localWS.Cells(x, col).Value)
  Do While shet.Cells(x, col).Value <> ""
    x = x + 1
  Loop

  ' Return the row number of the first empty cell
  nextEmptyRow = x

  Exit Function ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextEmptyRow Function of Module1"
End Function
' This function finds the next row in a specified column of a worksheet with the specified string.
'
' Args:
'   shet: The worksheet to search in.
'   col: The column letter (e.g., "A", "B", "C") to search in.
'   rowStart: (Optional) The row number to start the search from.
'             Defaults to 6 if not provided.
'
' Returns:
'   The row number of the first empty cell in the specified column.
Function nextRowWith(shet As Worksheet, col As Long, whatStr As String, Optional rowStart As Long = 2) As Long

    On Error GoTo ErrHandler ' Error handling
    
    ' Initialize the starting row
    Dim x As Long
    x = rowStart
    ' Loop through the cells in the specified column
    '  Do While Not IsEmpty(localWS.Cells(x, col).Value)
    Do While shet.Cells(x, col).Value <> whatStr
        If shet.Cells(x, col).Value = "" Then
         '   Err.Raise 0, , "No sand transfers needing completed."
            x = 0
            Exit Function
        Else
            
        End If
        x = x + 1
    Loop
    
    ' Return the row number of the first empty cell
    nextRowWith = x
    
    Exit Function ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextRowWith Function of Module1"
End Function
' This function records sand transferred to furnaces.
' It logs the transfer locally, to a backup spreadsheet, and if SAP is online,
' it also performs the transfer in SAP.
'
' Args:
'   itemNum: The item number of the sand.
'   batch: The batch number of the sand.
'   qty: The quantity of sand transferred.
'   fromLoc: The location the sand is transferred from.
'   toLoc: The location the sand is transferred to.
'
' Returns:
'   True if the transfer was successful in SAP (if online), False otherwise.
Function transferMaterial(ByVal itemNum As Long, ByVal Batch As String, QTY As Variant, fromLoc As String, toLoc As String, toSS As Boolean) As Boolean

  ' Declare variables
  Dim x As Long
  Dim onl As Boolean
  Dim thisMove As GuiSession

  On Error GoTo ErrHandler ' Error handling

  ' Check if SAP is online
  ' (Assumes "Online" status is in cell AA1 of "Sheet2")
  If onlineMode Then
    onl = True
  Else
    onl = False
  End If
    If toSS Then
        ' Record the transfer in the local "SandTrans" sheet
        x = nextEmptyRow(ThisWorkbook.Sheets("SandTrans"), 1, 2) ' Get next empty row
        With ThisWorkbook.Sheets("SandTrans")
            .Cells(x, 1).Value = itemNum
            .Cells(x, 2).Value = UCase(Batch)     ' Store batch in uppercase
            .Cells(x, 3).Value = QTY
            .Cells(x, 4).Value = fromLoc
            .Cells(x, 5).Value = toLoc
            .Cells(x, 6).Value = onl             ' Record online/offline status
        End With
    End If
  ' Save the transfer data to the backup spreadsheet
    If productionMode Then
        saveRemoteSandTransfers x
    End If
  ' If SAP is online, perform the transfer in SAP
  If onlineMode Then
    setStartScreen ' Initialize SAP screen (assuming this is a custom function)
    Set thisMove = SAPSession.CurrentSession ' Get the current SAP session

    With thisMove
      ' Navigate to the relevant transaction in SAP (ZMMS2S)
      .FindById("wnd[0]/tbar[0]/okcd").Text = "ZMMS2S"
      .FindById("wnd[0]").SendVKey 0 ' Enter

      ' Populate the required fields in SAP
      .FindById("wnd[0]/usr/txtMCHB-LGORT").Text = fromLoc
      .FindById("wnd[0]/usr/txtV_BATCH").Text = Batch
      .FindById("wnd[0]/usr/ctxtMARA-MATNR").Text = itemNum
      .FindById("wnd[0]").SendVKey 5 ' Enter

      .FindById("wnd[0]/usr/txtLAGP-RKAPV").Text = CDec(QTY) ' Quantity as decimal
      .FindById("wnd[0]/usr/txtT001L-LGORT").Text = toLoc
      .FindById("wnd[0]").SendVKey 5 ' Enter

      ' Check for success message in SAP
      If .FindById("wnd[0]/usr/txtV_MESSAGE3").Text = "Created." Then
        transferMaterial = True
      Else
        transferMaterial = False
      End If

      ' Close the transaction in SAP
      .FindById("wnd[0]").SendVKey 8 ' F8 (possibly to save)
      .FindById("wnd[0]/tbar[0]/btn[3]").Press ' Back button
    End With

    Set thisMove = Nothing ' Release the SAP session object
  End If

  Exit Function ' Normal exit

ErrHandler:
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in transferMaterial Function of Module1"
End Function
Public Sub massSandTransfer()
    Dim keepGoing As Boolean
    Dim x As Long
    Dim thisWB As Workbook
    Dim thisSheet As Worksheet
    x = 2
    Set thisWB = ThisWorkbook
    Set thisSheet = thisWB.Sheets("SandTrans")

    keepGoing = True
    With thisSheet
    Do While keepGoing
        If .Cells(x, 6).Value = "" Then
            keepGoing = False
        ElseIf .Cells(x, 6).Value = False Then
            If transferMaterial(.Cells(x, "A"), .Cells(x, "B").Value, .Cells(x, "C").Value, .Cells(x, "D").Value, .Cells(x, "E").Value, False) Then
                .Cells(x, "F").Value = True
            End If
        End If
        x = x + 1
    Loop
    End With
    MsgBox "Transfers completed."
    Set thisSheet = Nothing
    Set thisWB = Nothing
End Sub

