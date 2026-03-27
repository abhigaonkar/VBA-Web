VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SandForm 
   Caption         =   "Sand Transfer"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "SandForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SandForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_AfterUpdate()
    SandTransfer.Enabled = True
End Sub

Private Sub ComboBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
SandTransfer.Enabled = True
End Sub

Private Sub ComboBox1_Change()
    statusUpdate ""
End Sub

Private Sub CommandButton1_Click()
SandForm.Hide
NotSerialized.Show
End Sub

Private Sub SandTransfer_Click()
  ' Declare variables
    Dim sandItemNum As Variant
    Dim lastSand As Long
    Dim sandBatchNum As Variant, sandStatement As String, xLoc As String, fromLoc As String, lastBatch As String
    Dim thisSession As GuiSession
    Dim b As Integer, x As Integer
    Dim transferAmt As Variant, defAmt As Variant, xAmt As Variant
    Dim thisSandLevelCheck As clsFIFO ' Assuming clsFIFO is a custom class
    Dim xs As Worksheet
    
  On Error GoTo ErrHandler ' Error handling

  ' Get the current SAP session
    Set xs = ThisWorkbook.Sheets("SandTrans")
    
  ' Get the last used sand item number and batch from the "SandTrans" sheet
    With xs
        lastSand = FindLastOccurrence(ComboBox1.Value, .Range("e:e")).Offset(, -4)
        lastBatch = FindLastOccurrence(ComboBox1.Value, .Range("e:e")).Offset(, -3)
'        lastBatch = .Cells(nextEmptyRow(xs, 1, 2) - 1, 2).Value
'        lastSand = .Cells(nextEmptyRow(xs, 1, 2) - 1, 1).Value
    End With

  ' Get the sand item number and batch number from the user
  sandItemNum = InputBox("What sand item number are you transferring?", "Sand Transfer Form", lastSand)
  If sandItemNum = "" Then Exit Sub ' Exit if no item number is entered

  sandBatchNum = InputBox("What sand batch are you transferring?", "Sand Transfer Form", lastBatch)
  If sandBatchNum = "" Then Exit Sub ' Exit if no batch number is entered

  ' Check if the system is online (ToggleButton2 indicates online status)
    If onlineMode Then
        statusUpdate "Looking up data."
        ' --- SAP Interaction (Online) ---
        Set thisSession = SAPSession.CurrentSession
        setStartScreen ' Initialize SAP screen
    
        With thisSession
            .FindById("wnd[0]/tbar[0]/okcd").Text = "mb52" ' Open transaction MB52
            .FindById("wnd[0]").SendVKey 0 ' Enter
            
            ' Enter the sand item number, batch number, and storage location in SAP
            .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = sandItemNum
            .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = sandBatchNum
            .FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "Q105"
            .FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
            .FindById("wnd[0]/usr/chkXMCHB").Selected = True
            .FindById("wnd[0]/usr/chkNOZERO").Selected = True
            .FindById("wnd[0]/usr/chkNOVALUES").Selected = False
            .FindById("wnd[0]/usr/radPA_FLT").Select
            .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "0015"
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "0100"
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").CaretPosition = 4
            .FindById("wnd[1]/tbar[0]/btn[8]").Press
            .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
             
            .FindById("wnd[0]").SendVKey 8
            If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
                MsgBox "There is no sand in inventory with this item number/batch combination.", vbOKOnly
                Exit Sub
            End If
            b = .FindById("wnd[0]/usr").Children.Count / 8 - 1
            ' Build a string with sand availability information
            sandStatement = "Sand Availability:"
            For x = 3 To 2 + b
                sandStatement = sandStatement & vbCrLf
                xLoc = .FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Location
                xAmt = .FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
                sandStatement = sandStatement & "Location: " & xLoc
                sandStatement = sandStatement & vbTab & "Amount: " & xAmt
            Next x
            
            ' Get the transfer amount and source location from the user
            If b = 1 Then
                fromLoc = xLoc
                If xAmt < 250 Then
                  defAmt = CDec(xAmt)
                Else
                  defAmt = ""
                End If
                sandStatement = sandStatement & vbCrLf & vbCrLf
                transferAmt = InputBox(sandStatement & "How much would you like to transfer to this furnace?", "Sand Availability", defAmt)
            Else
                fromLoc = InputBox(sandStatement & vbCrLf & vbCrLf & "Which location would you like to transfer from?", "Sand Location Query")
                transferAmt = InputBox(sandStatement & vbCrLf & vbCrLf & "How much would you like to transfer to this furnace?", "Sand Availability")
            End If
    
        End With ' thisSession
    Else
        ' --- Offline Mode ---
        fromLoc = "" ' No location needed in offline mode
        transferAmt = InputBox("How much are you transferring?")
    End If

    ' Perform the transfer (record in spreadsheets and update SAP if online)
    Dim msgBoxTxt As String
    msgBoxTxt = "You're about to transfer item: " & vbTab & sandItemNum & vbCrLf _
     & "Batch number: " & vbTab & vbTab & vbTab & sandBatchNum & vbCrLf _
     & "Quantity: " & vbTab & vbTab & vbTab & transferAmt & vbCrLf _
     & "to " & vbTab & vbTab & vbTab & vbTab & ComboBox1.Value & vbCrLf _
     & "Is this correct?"
    If MsgBox(msgBoxTxt, vbYesNo) = vbYes Then
        transferMaterial sandItemNum, sandBatchNum, transferAmt, fromLoc, ComboBox1.Value, True
    End If
      
    Set thisSession = Nothing ' Release the SAP session object
    
    If MsgBox("Do you have more to transfer?", vbYesNo, "Stay Open") = vbNo Then
        SaveAndCloseWorkbook
    Else
        UserForm_Activate
    End If
    Exit Sub ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in cmdTransferSand_Click of UserForm1"

End Sub
Sub statusUpdate(msg As String)
    Label1.Caption = msg
    Label1.TextAlign = fmTextAlignCenter
End Sub


Private Sub UserForm_Activate()
    statusUpdate "Pick a furnace"

End Sub

