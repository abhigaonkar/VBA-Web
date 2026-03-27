VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cExpectedStatus As String, cCurrentStatus As String
Dim cSAPWindowName As String
Dim cPassed As Boolean, cAllowNotice As Boolean
Dim cMySession As GuiSession
Private Sub Class_Initialize()
On Error GoTo ErrHandler
    cAllowNotice = False
    killSessionAndStatus
    Set cMySession = SAPSession.CurrentSession
    
    Exit Sub
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Initialize of Class module clsStatus"
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandler

    Set cMySession = Nothing
    cExpectedStatus = vbNullString
    cCurrentStatus = vbNullString
    Exit Sub
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Terminate of Class module clsStatus"
End Sub

Public Property Get CurrentStatus() As String
On Error GoTo ErrHandler
    If cMySession Is Nothing Then Set cMySession = SAPSession.CurrentSession
    If cMySession.FindById("wnd[0]/sbar") Is Nothing Then
        cCurrentStatus = vbNullString
        Exit Property
    Else
        cCurrentStatus = cMySession.FindById("wnd[0]/sbar").Text
    End If
    CurrentStatus = cCurrentStatus
    Exit Property
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get CurrentStatus of Class module clsStatus"
End Property
Public Property Get SAPWindowName() As String
On Error GoTo ErrHandler
    cSAPWindowName = cMySession.ActiveWindow.Text
    SAPWindowName = cSAPWindowName
    Exit Property
ErrHandler:
    If Err.Number = 91 Then
        If ThisWorkbook.Sheets("Sheet2").Range("AA1").Value = "Online" Then
            If MsgBox("You are not logged in to SAP, do you want to stay offline?", vbYesNo) = vbYes Then
                UserForm1.ToggleButton2 = True
                Resume Next
            Else
                MsgBox "Please log in to SAP and then re-open this spreadsheet.", vbOKOnly
                Application.Quit
            End If
        Else
            UserForm1.ToggleButton2 = True
            UserForm1.ToggleButton2.Enabled = False
        End If
    Else
        StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get SAPWindowName of Class module clsStatus"
    End If
End Property

Public Property Get passed() As Boolean
    passed = cPassed
End Property
Public Property Let ExpectedStatus(vNewValue As String)
    cExpectedStatus = vNewValue
End Property
Function IsThereAModalWindow() As Boolean

  ' This function checks if there is a modal window open in the current SAP GUI session.

  Dim x As Integer  ' Loop counter
  
  On Error GoTo ErrHandler ' Enable error handling

  ' If the session object is not initialized, get the current session
  If cMySession Is Nothing Then
      Set cMySession = SAPSession.CurrentSession
  End If

  ' Loop through all the children of the current session
  For x = 0 To cMySession.Children.Count - 1
      ' Check if the child is a modal window
      If cMySession.Children(x).Type = "GuiModalWindow" Then
          IsThereAModalWindow = True ' Modal window found
          Exit Function ' Exit the function immediately after finding a modal window
      End If
  Next x

  ' If no modal window is found after checking all children, set the return value to False
  IsThereAModalWindow = False
  Exit Function

ErrHandler:
  ' Handle specific errors
  If Err.Number = 619 Then  ' Error 619: "The object does not support this property or method"
      Err.Clear ' Clear the error
      Resume Next ' Continue to the next iteration of the loop
  ElseIf Err.Number = 91 Then ' Error 91: "Object variable or With block variable not set"
      Err.Clear ' Clear the error
      Exit Function ' Exit the function
  End If

  ' Handle other errors by calling a custom error handler
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in IsThereAModalWindow of Class module clsStatus"
End Function

Function CheckModalWindow() As Boolean
    'returns true if there was a modal window and it was "cleared" or there is no modal window
    'returns false if there is a modal window that isn't resolved
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
On Error GoTo ErrHandler
    For x = 0 To cMySession.Children.Count - 1
        If cMySession.Children(x).Type = "GuiModalWindow" Then
            CheckModalWindow = False
            Dim myWindow As GuiModalWindow
            Set myWindow = cMySession.Children(x)
            If myWindow.PopupDialogText = "A confirmation has been entered Do you want to save the confirmation?" Then
                'MsgBox myWindow.PopupDialogText, , myWindow.Text
                
                For y = 0 To myWindow.Children.Count - 1
                    If myWindow.Children(y).Type = "GuiUserArea" Then
                        Dim myUserArea As GuiUserArea
                        Set myUserArea = myWindow.Children(y)
                            For z = 0 To myUserArea.Children.Count - 1
                                If myUserArea.Children(z).Type = "GuiButton" Then
                                    Dim myButton As GuiButton
                                    Set myButton = myUserArea.Children(z)
                                    If myButton.Text = "Yes" Then
                                        myButton.Press
                                        CheckModalWindow = True
                                        Exit Function
                                    Else
                                        CheckModalWindow = False
                                        Exit Function
                                    End If
                                    Set myButton = Nothing
                                End If
                            Next z
                        Set myUserArea = Nothing
                    End If
                Next y
            Else
                CheckModalWindow = True
            End If
            
            Set myWindow = Nothing
        End If
    Next x

    Set cMySession = Nothing
    
    Exit Function
    
ErrHandler:
    If Err.Number = 619 Then
        Err.Clear
        Resume Next
    ElseIf Err.Number = 91 Then
        Err.Clear
        Exit Function
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in CheckModalWindow of Class module clsStatus"

End Function

Function statusCheck(Optional ExpectedStatus As String, Optional allowNotice As Boolean) As Boolean
    Dim strStatus As String
On Error GoTo ErrHandler
    cExpectedStatus = ExpectedStatus
    cAllowNotice = allowNotice
    cPassed = True
    If Me.IsThereAModalWindow Then
        If cMySession.FindById("wnd[1]/usr/btnSPOP-OPTION2") Is Nothing Then 'check if there's a popup from SAP
            'no popup from SAP
        Else
            'popup exists from SAP
            If Not CheckModalWindow Then
                statusCheck = False
                Exit Function
            End If
        End If
    Else
        cCurrentStatus = cMySession.FindById("wnd[0]/sbar").Text
        If cExpectedStatus <> vbNullString Then
            If cCurrentStatus <> vbNullString And cCurrentStatus <> cExpectedStatus Then
             '   Application.WindowState = xlNormal
                If allowNotice Then MsgBox cCurrentStatus, vbOKOnly, "SAP Error Condition"
                cPassed = False
                Exit Function
            End If
        End If
        cCurrentStatus = cMySession.FindById("wnd[0]/usr/txtCMIMSG-HEADLINE").Text
        If cCurrentStatus = "Goods movements with errors" Then
      '      Application.WindowState = xlNormal
            If allowNotice Then MsgBox cCurrentStatus, vbOKOnly, "SAP Error Condition"
            cPassed = False
            Exit Function
        ElseIf cExpectedStatus <> vbNullString Then
            If cCurrentStatus <> vbNullString And cCurrentStatus <> cExpectedStatus Then
      '          Application.WindowState = xlNormal
                If allowNotice Then MsgBox cCurrentStatus, vbOKOnly, "SAP Error Condition"
                cMySession.FindById("wnd[0]/usr/btnI_DETAIL").Press
                cPassed = False
                Exit Function
            End If
        End If
    End If
    statusCheck = cPassed
    Exit Function
    
ErrHandler:
    If Err.Number = 619 Then
        Err.Clear
        Resume Next
    ElseIf Err.Number = 91 Then
        Err.Clear
        Exit Function
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in statusCheck of Class module clsStatus"
End Function




