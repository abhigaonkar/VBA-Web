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
    Private application As Object
    Private connection As Object
    Private session As Object

    ' Constructor to initialize the SAP GUI scripting engine
    Public Sub Initialize()
        Set SAPGUI = GetObject("SAPGUI")
        If Not SAPGUI Is Nothing Then
            Set application = SAPGUI.GetScriptingEngine
        Else
            Err.Raise vbObjectError + 1, "SAPSessionManager", "SAP GUI not found."
        End If
    End Sub

  ' Method to connect to an existing session
    Public Function ConnectToExistingSession(Optional SessionIndex As Integer = 0) As Boolean
        On Error GoTo ErrorHandler
        If application Is Nothing Then Me.Initialize
        If application.Children.Count > SessionIndex Then
            Set connection = application.Children(0)
            Set session = connection.Children(0)
            ConnectToExistingSession = True
        Else
            Err.Raise vbObjectError + 2, "SAPSessionManager", "No SAP GUI session found. Please log in to SAP and come back and click OK."
            ConnectToExistingSession = False
        End If
        Exit Function

ErrorHandler:
        ConnectToExistingSession = False
    '    If MsgBox("Error connecting to SAP GUI session: " & Err.Description, vbOKCancel) = vbCancel Then
    '        If MsgBox("Do you want to remain offline?", vbYesNo) = vbNo Then ThisWorkbook.Close
    '    End If
    End Function
    ' Method to get the current session
    Public Property Get CurrentSession() As Object
        If session Is Nothing Then Me.ConnectToExistingSession
        Set CurrentSession = session
    End Property
