VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSAPSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class to manage SAP GUI sessions
Private SAPGUI As Object
Private SAPApplication As Object
Private connection As Object
Private session As Object

' Constructor to initialize the SAP GUI scripting engine
Public Sub Initialize()
  On Error Resume Next ' Handle potential errors during initialization
  Set SAPGUI = GetObject("SAPGUI")
  If Not SAPGUI Is Nothing Then
    Set SAPApplication = SAPGUI.GetScriptingEngine
  Else
    ' Handle the case where SAP GUI is not found
    ' Consider logging an error or displaying a message to the user
    startNetweaver
  End If
End Sub

' Method to connect to an existing session
Public Function ConnectToExistingSession(Optional SessionIndex As Integer = 0) As Boolean
  On Error GoTo ErrorHandler
  If SAPApplication Is Nothing Then Me.Initialize

  If SAPApplication.Children.Count > SessionIndex Then
    Set connection = SAPApplication.Children(0) ' Use the provided SessionIndex
    Set session = connection.Children(0)
    ConnectToExistingSession = True
  Else
    ' No existing session found, start a new one
    startNetweaver
    ' After starting Netweaver, try to connect again
    If SAPApplication.Children.Count > SessionIndex Then
      Set connection = SAPApplication.Children(SessionIndex)
      Set session = connection.Children(0)
      ConnectToExistingSession = True
    Else
      ' Still unable to connect, handle the error
      ' Consider logging an error or displaying a message to the user
      ConnectToExistingSession = False
    End If
  End If

  Exit Function

ErrorHandler:
  ConnectToExistingSession = False
  ' Handle the error appropriately (e.g., logging, user message)
End Function

' Method to get the current session
Public Property Get CurrentSession() As Object
  If session Is Nothing Then Me.ConnectToExistingSession
  Set CurrentSession = session
End Property

Private Sub startNetweaver()
  ' Purpose: Starts SAP Logon and connects to the Netweaver system.

    Dim path As String ' Variable to store the path to saplogon.exe
    Dim yy As Long ' Variable to store the process ID (optional)
    On Error GoTo ErrHandler ' Enable error handling
    
    MsgBox "SAP Logon Starting. One moment please.", vbOKOnly ' Inform the user
    
    path = "C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe" ' Set the path
    
    yy = Shell(path, vbNormalFocus) ' Launch SAP Logon

  ' *** IMPORTANT: Consider improving this section ***
    PauseMe 4, True ' Pause for 4 seconds (consider making this dynamic)
    Application.SendKeys "{UP}"
    Application.SendKeys "{UP}"
    Application.SendKeys "{UP}"
    Application.SendKeys ("s") ' Send "s" to select the Netweaver entry (assumes it's the second entry in list)
    Application.SendKeys ("~") ' Send Enter to log in
    Application.SendKeys "{UP}"
    ' ***  This is unreliable. Explore alternatives. ***

  Exit Sub ' Normal exit

ErrHandler:
  StdErrorHandler "Error in startNetweaver of clsSAPSession - " & Err & ": " & Error(Err) ' Custom error handling
End Sub    ' Method to start a transaction
