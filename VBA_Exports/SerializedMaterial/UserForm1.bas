VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Furnace Deck Interface"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13350
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hrs As Single
Dim localWB As Workbook
Dim localWS As Worksheet
Dim appControl As Worksheet
Private ODMax As Variant
Private ODMin As Variant
Private wallMax As Variant
Private wallMin As Variant
Private thisItemNum As String


Private Sub bow_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.bow.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub

Private Sub ComboBox1_Change()
    Dim defect As String
    Dim hideIt As Boolean
    hideIt = False
    defect = ComboBox1.Value
    Select Case defect
        Case "BATW"
            hideIt = True
        Case "BBBL"
            hideIt = True
        Case "BLSX"
            hideIt = True
        Case "CLEG"
            hideIt = True
        Case "DBUB"
            hideIt = True
        Case "DCOL"
            hideIt = True
        Case "Gray"
            hideIt = True
        Case "METL"
            hideIt = True
        Case "METS"
            hideIt = True
        Case "MMP7"
            hideIt = True
        Case "MOLO"
            hideIt = True
        Case "MOLY"
            hideIt = True
        Case "MSCD"
            hideIt = True
        Case "SACH"
            hideIt = True
    End Select
    If hideIt Then
        Webb.Visible = False
        ScrapD.Visible = False
    Else
        Webb.Visible = True
        ScrapD.Visible = True
    End If
End Sub

Private Sub CommandButton1_Click()
'for 215
    If Not checkInputs(TextBox1, TextBox3) Then Exit Sub
    Dim p As Long
    p = offlineProcess(TextBox1, TextBox3, tpt215)
    localWS.Cells(p, 5).Value = ProcOrdNum215 'save the process order number instead of looking it up
    saveRemoteData p, False
    If onlineMode Then
        onlineProcess (p)
    Else
        Me.Hide
        tubeIDAlert.Label1.Caption = localWS.Cells(p, 8)
        tubeIDAlert.Show
    End If
    saveRemoteData p, True
    ThisWorkbook.Sheets("Sheet2").Cells(22, 13).Value = TextBox1.Value
    ThisWorkbook.Sheets("Sheet2").Cells(23, 13).Value = TextBox3.Value
    ThisWorkbook.Sheets("Sheet2").Cells(24, 13).Value = label72.Value
    ThisWorkbook.Sheets("Sheet2").Cells(25, 13).Value = tpt215.Value
    MsgBox "This transaction is complete.", vbOKOnly
    SaveAndCloseWorkbook
End Sub

Private Function offlineProcess(ByVal palletCard As Long, ByVal itemNum As Long, ByVal timePerTube As Variant) As Long
    'it returns the row number in the spreadsheet
    ' This subroutine handles the offline processing for creating a new tube entry.
    ' It gathers data from the user interface, updates the spreadsheet, but does NOT run the COR6 transaction in SAP.
    Dim p As Long
    p = nextTube(furnName)
    LogEvent "Saving pallet card " & palletCard & " item " & itemNum
    If ToggleButton1 Then 'if this item in non-conforming
        If ComboBox1.Value = "" Then
            MsgBox "Please select a defect reason.", vbOKOnly
            ComboBox1.SetFocus
            offlineProcess = 0
            Exit Function
        End If
        If Webb Then
            thisItemNum = WebbItemCombo.Value
        ElseIf ScrapD Then
            thisItemNum = "2100094"
        ElseIf plate Then
            thisItemNum = "2304043"
        ElseIf Trash Then
            thisItemNum = "trash"
        End If
 '       Me.NOTES = thisItemNum
        localWS.Cells(p, 23) = thisItemNum
        localWS.Cells(p, 24) = True
        localWS.Cells(p, 25) = ComboBox1.Column(1)
    Else
        localWS.Cells(p, 24) = False
    End If
    localWS.Cells(p, 22) = False
    updSpreadsheetBeforeSAP p, palletCard, itemNum, timePerTube
    Dim theseDims As New clsDimEntry ' Create an instance of the clsDimEntry class
    theseDims.itemNum = localWS.Cells(p, 6) ' Set the item number
    theseDims.batch = localWS.Range("H" & p).Value ' Set the batch number (same as tube serial number)
    ' Set the dimension values from controls on the form (OD1, OD2, WALL1, etc.)
    theseDims.OD1 = localWS.Cells(p, 9)
    theseDims.OD2 = localWS.Cells(p, 10)
    theseDims.OD3 = localWS.Cells(p, 11)
    theseDims.OD4 = localWS.Cells(p, 12)
    theseDims.WALL1 = localWS.Cells(p, 13)
    theseDims.WALL2 = localWS.Cells(p, 14)
    theseDims.WALL3 = localWS.Cells(p, 15)
    theseDims.WALL4 = localWS.Cells(p, 16)
    theseDims.WALL5 = localWS.Cells(p, 17)
    theseDims.WALL6 = localWS.Cells(p, 18)
    theseDims.WALL7 = localWS.Cells(p, 19)
    theseDims.WALL8 = localWS.Cells(p, 20)
    theseDims.actualLength = localWS.Cells(p, 36)
    theseDims.bow = localWS.Cells(p, 21).Value
    theseDims.updateSHwithDims p
    Set theseDims = Nothing
'    clearForm ' Call a function to clear the form
'    MsgBox "This tube's serial number is " & localWS.Range("H" & x).Value
    offlineProcess = p
    
End Function

Private Sub CommandButton10_Click()
If Label94.Visible = True Then
    Label94.Visible = False
    CommandButton10.Caption = "Show Specs"
Else
    Label94.Visible = True
    CommandButton10.Caption = "Hide Specs"
End If

End Sub

Private Sub CommandButton11_Click()
If Label95.Visible = True Then
    Label95.Visible = False
    CommandButton11.Caption = "Show Specs"
Else
    Label95.Visible = True
    CommandButton11.Caption = "Hide Specs"
End If

End Sub

Private Sub CommandButton4_Click()
On Error GoTo ErrHandler
    If Not checkInputs(TextBox22, TextBox23) Then Exit Sub
    
    Dim p As Long
    p = offlineProcess(TextBox22, TextBox23, tptMisc)

    If onlineMode Then
        onlineProcess p
    Else
        Me.Hide
        tubeIDAlert.Label1.Caption = localWS.Cells(p, 8)
        tubeIDAlert.Show
    End If
    saveRemoteData p, False
    ThisWorkbook.Sheets("Sheet2").Cells(22, 14).Value = TextBox22.Value
    ThisWorkbook.Sheets("Sheet2").Cells(23, 14).Value = TextBox23.Value
    ThisWorkbook.Sheets("Sheet2").Cells(24, 14).Value = label72.Value
    ThisWorkbook.Sheets("Sheet2").Cells(25, 14).Value = tptMisc.Value
    MsgBox "This transaction is complete."
    SaveAndCloseWorkbook
Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in CommandButton4_Click Sub of UserForm1"
End Sub

' This subroutine handles the transfer of sand to a furnace.
' It records the transfer in two spreadsheets and updates SAP inventory
' if the system is online.
Private Sub CommandButton6_Click()

  ' Declare variables
    Dim sandItemNum As Variant
    Dim lastSand As Long
    Dim sandBatchNum As Variant, sandStatement As String, xLoc As String, fromLoc As String, lastBatch As String
    Dim thisSession As GuiSession
    Dim b As Integer, x As Integer
    Dim transferAmt As Variant, defAmt As Variant, xAMT As Variant
    Dim thisSandLevelCheck As clsFIFO ' Assuming clsFIFO is a custom class
    Dim xs As Worksheet
    Dim rowNum As Long
    Dim onl As String
  On Error GoTo ErrHandler ' Error handling

  ' Get the current SAP session
    Set xs = ThisWorkbook.Sheets("SandTrans")
    
  ' Get the last used sand item number and batch from the "SandTrans" sheet
    With xs
        lastBatch = .Cells(nextEmptyRow(xs, "A", 2) - 1, 2).Value
        lastSand = .Cells(nextEmptyRow(xs, "A", 2) - 1, 1).Value
    End With

  ' Get the sand item number and batch number from the user
  sandItemNum = InputBox("What sand item number are you transferring?", "Sand Transfer Form", lastSand)
  If sandItemNum = "" Then Exit Sub ' Exit if no item number is entered

  sandBatchNum = InputBox("What sand batch are you transferring?", "Sand Transfer Form", lastBatch)
  If sandBatchNum = "" Then Exit Sub ' Exit if no batch number is entered

  ' Check if the system is online (ToggleButton2 indicates online status)
    If Not ToggleButton2 Then
        onl = "True"
        ' --- SAP Interaction (Online) ---
        Set thisSession = SAPSession.CurrentSession
    '    setStartScreen ' Initialize SAP screen
        
        With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52" ' Open transaction MB52
        .FindById("wnd[0]").SendVKey 0 ' Enter
        
        ' Enter the sand item number, batch number, and storage location in SAP
        .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = sandItemNum
        .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = sandBatchNum
        .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
        
        .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
        .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
        .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
        .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = furnName
        .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "0400"
        .FindById("wnd[1]/tbar[0]/btn[8]").Press
          
        .FindById("wnd[0]").SendVKey 8
        If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
            MsgBox "There is no sand in inventory with this item number/batch combination.", vbOKOnly
            Exit Sub
        End If
        .FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status")
        b = .FindById("wnd[1]/usr/lbl[17,8]").Text ' Extract the number of entries from the status window
        .FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window
        ' Build a string with sand availability information
        sandStatement = "Sand Availability:"
        For x = 3 To 2 + b
            sandStatement = sandStatement & vbCrLf
            xLoc = .FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Location
            xAMT = .FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
            sandStatement = sandStatement & "Location: " & xLoc
            sandStatement = sandStatement & vbTab & "Amount: " & xAMT
        Next x
        
        ' Get the transfer amount and source location from the user
        If b = 1 Then
            fromLoc = xLoc
            If xAMT < 250 Then
                defAmt = CDec(xAMT)
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
    Set thisSession = Nothing ' Release the SAP session object
    Else
        ' --- Offline Mode ---
        fromLoc = "" ' No location needed in offline mode
        transferAmt = InputBox("How much are you transferring?")
        onl = "False"
    End If
    rowNum = nextEmptyRow(xs, "A", 2)
    xs.Cells(rowNum, 1) = sandItemNum
    xs.Cells(rowNum, 2).Value = sandBatchNum
    xs.Cells(rowNum, 3).Value = transferAmt
    xs.Cells(rowNum, 4).Value = fromLoc
    xs.Cells(rowNum, 5).Value = furnName
    xs.Cells(rowNum, 6).Value = onl

    ' Perform the transfer (record in spreadsheets and update SAP if online)
    If onl = "True" Then
        transferMatFromSS rowNum
        Set thisSandLevelCheck = New clsFIFO ' Create an instance of clsFIFO
        thisSandLevelCheck.useMB52 furnName ' Update sand level using MB52 data (assuming this is a method of clsFIFO)
        Set thisSandLevelCheck = Nothing ' Release the object
    End If
    
    setSandMeter ' Update the sand level indicator on the form
      
    
    Exit Sub ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in CommandButton6_Click of UserForm1"

End Sub

Private Sub CommandButton7_Click()
Me.Hide
frmReports.Show
End Sub

Private Sub CommandButton8_Click()
If Label91.Visible = True Then
    Label91.Visible = False
    CommandButton8.Caption = "Show Specs"
Else
    Label91.Visible = True
    CommandButton8.Caption = "Hide Specs"
End If
End Sub


Private Sub CommandButton9_Click()
    If MsgBox("This clears all current measurement fields so you may start again. Is this what you want to do?", vbYesNo) = vbNo Then Exit Sub
    OD1.Text = "" ' Clear OD1 field
    OD1.BackColor = RGB(192, 224, 255)
    OD2.Text = "" ' Clear OD2 field
    OD2.BackColor = RGB(192, 224, 255)
    OD3.Text = "" ' Clear OD3 field
    OD3.BackColor = RGB(192, 224, 255)
    OD4.Text = "" ' Clear OD4 field
    OD4.BackColor = RGB(192, 224, 255)
    WALL1.Text = "" ' Clear WALL1 field
    WALL1.BackColor = RGB(255, 192, 204)
    WALL2.Text = "" ' Clear WALL2 field
    WALL2.BackColor = RGB(255, 192, 204)
    WALL3.Text = "" ' Clear WALL3 field
    WALL3.BackColor = RGB(255, 192, 204)
    WALL4.Text = "" ' Clear WALL4 field
    WALL4.BackColor = RGB(255, 192, 204)
    WALL5.Text = "" ' Clear WALL5 field
    WALL5.BackColor = RGB(255, 192, 204)
    WALL6.Text = "" ' Clear WALL6 field
    WALL6.BackColor = RGB(255, 192, 204)
    WALL7.Text = "" ' Clear WALL7 field
    WALL7.BackColor = RGB(255, 192, 204)
    WALL8.Text = "" ' Clear WALL8 field
    WALL8.BackColor = RGB(255, 192, 204)
    bow.Text = "" ' Clear BOW field
    TextBox26.Text = ""
    OD1.SetFocus
End Sub





Private Sub MultiPage1_Change()
    If MultiPage1.Value = 0 Then
        MultiPage1.BackColor = vbGreen
    ElseIf MultiPage1.Value = 1 Then
        MultiPage1.BackColor = vbRed
'        TextBox26.Value = dataWS.Cells(26, 12).Value
    ElseIf MultiPage1.Value = 2 Then
        MultiPage1.BackColor = vbBlue
'        TextBox26.Value = dataWS.Cells(26, 13).Value
    ElseIf MultiPage1.Value = 3 Then
        MultiPage1.BackColor = vbYellow
'        TextBox26.Value = dataWS.Cells(26, 14).Value
    End If

    TextBox26.Value = dataWS.Cells(26, MultiPage1.Value + 11).Value
    ODMax = dataWS.Cells(MultiPage1.Value * 8 + 1, 24).Value
    ODMin = dataWS.Cells(MultiPage1.Value * 8 + 2, 24).Value
    wallMax = dataWS.Cells(MultiPage1.Value * 8 + 3, 24).Value
    wallMin = dataWS.Cells(MultiPage1.Value * 8 + 4, 24).Value

End Sub

Private Sub NOTES_AfterUpdate()
    Dim strLen As Long
    strLen = Len(NOTES.Text)
    If strLen > 40 Then
        MsgBox "You've entered " & strLen & " characters for your note. " _
         & "This will be saved unchanged with the spreadsheet. " _
         & "However, it will be truncated to 40 characters for saving in SAP. " _
         & "It will appear as:" & vbCrLf & "'" & left(NOTES.Text, 40) & "'.", vbInformation, "You are long winded!"
    End If
End Sub



Private Sub OD1_AfterUpdate()
    If Not checkODInput(OD1.Value) Then
        OD1.SetFocus
        OD1.BackColor = vbRed
    Else
        OD1.BackColor = vbGreen
    End If
End Sub

Private Sub OD1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.OD1.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub

Private Sub OD2_AfterUpdate()
    If Not checkODInput(OD2.Value) Then
        OD2.SetFocus
        OD2.BackColor = vbRed
    Else
        OD2.BackColor = vbGreen
    End If
End Sub
Private Sub OD2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.OD2.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub

Private Sub OD3_AfterUpdate()
    If Not checkODInput(OD3.Value) Then
        OD3.SetFocus
        OD3.BackColor = vbRed
    Else
        OD3.BackColor = vbGreen
    End If
End Sub
Private Sub OD3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.OD3.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub

Private Sub OD4_AfterUpdate()
    If Not checkODInput(OD4.Value) Then
        OD4.SetFocus
        OD4.BackColor = vbRed
    Else
        OD4.BackColor = vbGreen
    End If
End Sub
Private Sub OD4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.OD4.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub


Function checkODInput(data As Variant) As Boolean
    checkODInput = True
    If CDec(data) > ODMax + 10 Then
        MsgBox "This OD is out of range. The maximum you may input is " & ODMax + 10 & ". Please correct.", vbCritical
        checkODInput = False
    ElseIf CDec(data) < ODMin - 10 Then
        MsgBox "This OD is out of range. The minimum you may input is " & ODMin - 10 & ". Please correct.", vbCritical
        checkODInput = False
    End If
End Function

Private Sub plate_Click()
    WebbItemCombo.Visible = False
    Label89.Visible = False
End Sub

Private Sub TextBox26_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.TextBox26.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub

Private Sub wall1_AfterUpdate()
    If Not checkWallInput(WALL1.Value) Then
        WALL1.SetFocus
        WALL1.BackColor = vbRed
    Else
        WALL1.BackColor = vbGreen
    End If
    WALL1.Value = CDec(WALL1.Value)
End Sub
Private Sub wall2_AfterUpdate()
    If Not checkWallInput(WALL2.Value) Then
        WALL2.SetFocus
        WALL2.BackColor = vbRed
    Else
        WALL2.BackColor = vbGreen
    End If
    WALL2.Value = CDec(WALL2.Value)
End Sub
Private Sub wall3_AfterUpdate()
    If Not checkWallInput(WALL3.Value) Then
        WALL3.SetFocus
        WALL3.BackColor = vbRed
    Else
        WALL3.BackColor = vbGreen
    End If
    WALL3.Value = CDec(WALL3.Value)
End Sub
Private Sub wall4_AfterUpdate()
    If Not checkWallInput(WALL4.Value) Then
        WALL4.SetFocus
        WALL4.BackColor = vbRed
    Else
        WALL4.BackColor = vbGreen
    End If
    WALL4.Value = CDec(WALL4.Value)
End Sub
Private Sub wall5_AfterUpdate()
    If Not checkWallInput(WALL5.Value) Then
        WALL5.SetFocus
        WALL5.BackColor = vbRed
    Else
        WALL5.BackColor = vbGreen
    End If
    WALL5.Value = CDec(WALL5.Value)
End Sub
Private Sub wall6_AfterUpdate()
    If Not checkWallInput(WALL6.Value) Then
        WALL6.SetFocus
        WALL6.BackColor = vbRed
    Else
        WALL6.BackColor = vbGreen
    End If
    WALL6.Value = CDec(WALL6.Value)
End Sub
Private Sub wall7_AfterUpdate()
    If Not checkWallInput(WALL7.Value) Then
        WALL7.SetFocus
        WALL7.BackColor = vbRed
    Else
        WALL7.BackColor = vbGreen
    End If
    WALL7.Value = CDec(WALL7.Value)
End Sub
Private Sub wall8_AfterUpdate()
    If Not checkWallInput(WALL8.Value) Then
        WALL8.SetFocus
        WALL8.BackColor = vbRed
    Else
        WALL8.BackColor = vbGreen
    End If
    WALL8.Value = CDec(WALL8.Value)
End Sub

Function checkWallInput(data As Variant) As Boolean
    checkWallInput = True
    If CDec(data) > wallMax + 3 Then
        MsgBox "This wall is out of range. The maximum you may input is " & wallMax + 3 & ". Please correct.", vbCritical
        checkWallInput = False
    ElseIf CDec(data) < wallMin - 3 Then
        MsgBox "This wall is out of range. The minimum you may input is " & wallMin - 3 & ". Please correct.", vbCritical
        checkWallInput = False
    End If
End Function
Private Sub wall1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL1.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL2.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL3.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL4.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL5.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL6.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL7.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub wall8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' Allow digits 0-9
            ' Do nothing, allow the key press
            
        Case 46 ' Allow the decimal point character "." (ASCII 46)
            ' Check if a decimal point already exists in the textbox
            If InStr(1, Me.WALL8.Text, ".") > 0 Then
                ' If it exists, cancel the key press
                KeyAscii = 0
                Beep ' Optional: Notify user
            End If
            ' If no decimal exists yet, the key press is allowed implicitly
            
        Case vbKeyBack ' Allow Backspace
            ' Do nothing, allow the key press
            
        Case Else ' Disallow all other characters
            KeyAscii = 0 ' Cancel the key press
            Beep ' Optional: Notify user
    End Select

End Sub
Private Sub ProcessIt_Click()
On Error GoTo ErrHandler
    If Not checkInputs(pCardNum, procOrdNum) Then Exit Sub
    Dim p As Long
    p = offlineProcess(pCardNum, procOrdNum, tpt)
    saveRemoteData p, False
    If onlineMode Then 'if we're in online mode
        If ComboBox1.Visible Then  'if this item in non-conforming in online mode
            If Webb Then
                If MsgBox("You are saving one piece of Webb material using " & tpt & " hours. Does this look correct?", vbYesNo) = vbNo Then Exit Sub
            ElseIf plate Then
                If MsgBox("You are saving one piece of plate material using " & tpt & " hours. Does this look correct?", vbYesNo) = vbNo Then Exit Sub
            ElseIf ScrapD Then
                If MsgBox("You are saving one piece of scrap D material using " & tpt & " hours. Does this look correct?", vbYesNo) = vbNo Then Exit Sub
            ElseIf Trash Then
                If MsgBox("You are scrapping one piece of " & label70.Value & " using " & tpt & " hours. Does this look correct?", vbYesNo) = vbNo Then Exit Sub
            Else
'                onlineProcess p
            End If
            onlineProcess p
        Else ' if this tube in conforming in online mode
            If MsgBox("You are processing one " & label70.Value & " in online mode, using " & tpt & " hours. Does this look correct?", vbYesNo) = vbNo Then
                Exit Sub
            Else
                onlineProcess p
            End If
        End If
    Else
        Me.Hide
        tubeIDAlert.Label1.Caption = localWS.Cells(p, 8)
        tubeIDAlert.Show
    End If
    
    saveRemoteData p, True
    ThisWorkbook.Sheets("Sheet2").Cells(22, 11).Value = pCardNum
    ThisWorkbook.Sheets("Sheet2").Cells(23, 11).Value = procOrdNum
    ThisWorkbook.Sheets("Sheet2").Cells(24, 11).Value = label70.Value
    ThisWorkbook.Sheets("Sheet2").Cells(25, 11).Value = tpt.Value
    MsgBox "This transaction is complete."
    SaveAndCloseWorkbook
    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in processIt_Click of UserForm1"
End Sub

Private Sub procOrdNum_AfterUpdate()
On Error GoTo ErrHandler
'when the item number changes
    label70.Value = ""
    If onlineMode And procOrdNum.Text <> "" Then
        statusUpdate "Looking up item info"
        setStartScreen
        Dim thisItem As New clsMaterial
        dataWS.Cells(24, 11).Value = thisItem.getDescription(procOrdNum.Text)
        label70.Value = dataWS.Cells(24, 11).Value
        tpt.Value = Round(thisItem.LENGTH / 1524, 2)
        dataWS.Cells(27, 11).Value = thisItem.getSpecsAsString
        Label91.Caption = dataWS.Cells(27, 11).Value
        'complete range setting for dimensional inputs
        thisItem.calculateMaxAndMins
        ODMax = thisItem.ODMax
        ODMin = thisItem.ODMin
        wallMax = thisItem.wallMax
        wallMin = thisItem.wallMin
        dataWS.Cells(26, 11).Value = Round((thisItem.LENGTH / 25.4) + 1.5, 2)
        TextBox26.Value = dataWS.Cells(26, 11).Value
        CommandButton8.Visible = True
        Set thisItem = Nothing
        statusUpdate ""
    End If
    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in procOrdNum_AfterUpdate of UserForm1"
End Sub

Private Sub procOrdNum84130_AfterUpdate()
'when the item number changes
On Error GoTo ErrHandler
    label71.Value = ""
    If onlineMode And procOrdNum84130.Text <> "" Then
        statusUpdate "Looking up item info"
        Dim thisItem As New clsMaterial
        dataWS.Cells(23, 12).Value = procOrdNum84130.Text
        dataWS.Cells(24, 12).Value = thisItem.getDescription(dataWS.Cells(23, 12).Value)
        label71.Value = dataWS.Cells(24, 12).Value
        dataWS.Cells(25, 12).Value = Round(thisItem.LENGTH / 1524, 2)
        tpt84130.Value = Round(dataWS.Cells(25, 12).Value, 2)
        dataWS.Cells(26, 12).Value = Round((thisItem.LENGTH / 25.4) + 1.5, 2)
        TextBox26.Value = dataWS.Cells(26, 12).Value
        dataWS.Cells(27, 12).Value = thisItem.getSpecsAsString
        Label94.Caption = dataWS.Cells(27, 12).Value
        thisItem.calculateMaxAndMins
        dataWS.Cells(9, 24).Value = thisItem.ODMax
        ODMax = dataWS.Cells(9, 24).Value
        dataWS.Cells(10, 24).Value = thisItem.ODMin
        ODMin = dataWS.Cells(10, 24).Value
        dataWS.Cells(11, 24).Value = thisItem.wallMax
        wallMax = dataWS.Cells(11, 24).Value
        dataWS.Cells(12, 24).Value = thisItem.wallMin
        wallMin = dataWS.Cells(12, 24).Value
        Set thisItem = Nothing
    End If
        statusUpdate ""
    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in procOrdNum84130_AfterUpdate of UserForm1"
End Sub

Private Sub Save84130_Click()
On Error GoTo ErrHandler

    If Not checkInputs(TextBox21, procOrdNum84130) Then Exit Sub
    Dim p As Long
    p = offlineProcess(TextBox21, procOrdNum84130, tpt84130.Value)
    saveRemoteData p, False
    If onlineMode Then
        onlineProcess p
    Else
        Me.Hide
        tubeIDAlert.Label1.Caption = localWS.Cells(p, 8)
        tubeIDAlert.Show
    End If
    saveRemoteData p, True
    localWB.Sheets("Sheet2").Cells(22, 12).Value = TextBox21.Value
    localWB.Sheets("Sheet2").Cells(23, 12).Value = procOrdNum84130.Value
    localWB.Sheets("Sheet2").Cells(24, 12).Value = label71.Value
    localWB.Sheets("Sheet2").Cells(25, 12).Value = tpt84130.Value
    SaveAndCloseWorkbook
Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Save84130_Click of UserForm1"
End Sub

Private Sub ScrapD_Click()
    WebbItemCombo.Visible = False
    Label89.Visible = False
End Sub

Private Sub TextBox23_AfterUpdate()
On Error GoTo ErrHandler

    If onlineMode And TextBox23.Text <> "" Then
        Dim thisItem As New clsMaterial
        label77.Value = thisItem.getDescription(TextBox23.Text)
        tptMisc = Round(thisItem.LENGTH / 1524, 2)
        thisItem.calculateMaxAndMins
        ODMax = thisItem.ODMax
        ODMin = thisItem.ODMin
        wallMax = thisItem.wallMax
        wallMin = thisItem.wallMin
        TextBox26.Value = Round(thisItem.LENGTH / 25.4, 2)
        dataWS.Cells(26, 14).Value = Round(thisItem.LENGTH / 25.4, 2)
        Set thisItem = Nothing
    End If

    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in TextBox23_AfterUpdate of UserForm1"
End Sub


Private Sub TextBox3_AfterUpdate()
    ' This event handler is triggered when the value in TextBox3 (presumably the item number) is updated.
    ' It retrieves and displays the item description from SAP if the "Offline" mode is not active.
On Error GoTo ErrHandler
    label72.Value = ""
    If onlineMode And TextBox3.Text <> "" Then ' Check if "Offline" mode (ToggleButton2) is NOT active and TextBox3 is not empty
        
        Dim thisItem As New clsMaterial
        dataWS.Cells(23, 13).Value = TextBox3.Text
        dataWS.Cells(24, 13).Value = thisItem.getDescription(dataWS.Cells(23, 13).Value)
        label72.Value = dataWS.Cells(24, 13).Value
        dataWS.Cells(25, 13).Value = Round(thisItem.LENGTH / 1524, 2)
        tpt215.Value = Round(dataWS.Cells(25, 13).Value, 2)
        dataWS.Cells(26, 13).Value = Round((thisItem.LENGTH / 25.4) + 1.5, 2)
        TextBox26.Value = dataWS.Cells(26, 13).Value
        dataWS.Cells(27, 13).Value = thisItem.getSpecsAsString
        Label95.Caption = dataWS.Cells(27, 13).Value
        thisItem.calculateMaxAndMins
        dataWS.Cells(17, 24).Value = thisItem.ODMax
        ODMax = dataWS.Cells(17, 24).Value
        dataWS.Cells(18, 24).Value = thisItem.ODMin
        ODMin = dataWS.Cells(18, 24).Value
        dataWS.Cells(19, 24).Value = thisItem.wallMax
        wallMax = dataWS.Cells(19, 24).Value
        dataWS.Cells(20, 24).Value = thisItem.wallMin
        wallMin = dataWS.Cells(20, 24).Value
        Set thisItem = Nothing
    End If
    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in TextBox3_AfterUpdate of UserForm1"
End Sub


Private Sub ToggleButton2_Click()
    ' This event handler is triggered when ToggleButton2 (presumably the "Online/Offline" toggle) is clicked.
    ' It switches the application between "Online" and "Offline" modes and updates the UI accordingly.

    If ToggleButton2 Then ' If the toggle button is currently ON (meaning "Offline" mode is active)
        If ThisWorkbook.Sheets("Sheet2").Range("AA1").Value = "Online" Then
            If MsgBox("Are you sure that SAP is offline?", vbYesNo) = vbNo Then ' Confirm with the user if they want to go online
                ToggleButton2 = False ' If the user clicks "No", keep the "Offline" mode active
                OD1.SetFocus ' Set focus to the OD1 field
                Exit Sub ' Exit the subroutine
            Else ' If the user clicks "Yes"
                MsgBox "If you think Jesse Parkinson is not aware of this outage, please call him now. 740-334-5289", vbOKOnly ' Display a message about the outage
                ThisWorkbook.Sheets("Sheet2").Range("AA1").Value = "Offline"
                ToggleButton2.Enabled = False
                onlineMode = False
            End If
        End If
        UserForm1.BackColor = vbRed ' Set the background color of the form to red to indicate "Offline" mode
        ToggleButton2.Caption = "Offline" ' Change the caption of the toggle button to "Offline"
        'ToggleButton2.ControlTipText = "Click here to go ONLINE" ' Update the tooltip to reflect the current state

    Else ' If the toggle button is currently OFF (meaning "Online" mode is active)
        UserForm1.BackColor = &H8000000F ' Set the background color of the form back to the default
        ToggleButton2.Caption = "Online" ' Change the caption of the toggle button to "Online"
        ToggleButton2.ControlTipText = "Click here to go OFFLINE" ' Update the tooltip to reflect the current state
    End If
End Sub


Private Sub ToggleButton1_Click()
    If ToggleButton1.Value = True Then
        ComboBox1.Visible = True
        Label17.Visible = True
        Webb.Visible = True
        ScrapD.Visible = True
        Trash.Visible = True
        plate.Visible = True
       ' ToggleButton1 = False
        Label79.Caption = "BAD TUBE"
    Else
        ComboBox1.Visible = False
        Label17.Visible = False
        Webb.Visible = False
        ScrapD.Visible = False
        Trash.Visible = False
        plate.Visible = False
        Label79.Caption = "GOOD TUBE"
        WebbItemCombo.Visible = False
'        ToggleButton1 = True
        Label89.Visible = False
    End If
End Sub
Private Sub changeMeasurements(toWhat As Boolean)
    OD1.Enabled = toWhat
    OD2.Enabled = toWhat
    OD3.Enabled = toWhat
    OD4.Enabled = toWhat
    bow.Enabled = toWhat
    WALL1.Enabled = toWhat
    WALL2.Enabled = toWhat
    WALL3.Enabled = toWhat
    WALL4.Enabled = toWhat
    WALL5.Enabled = toWhat
    WALL6.Enabled = toWhat
    WALL7.Enabled = toWhat
    WALL8.Enabled = toWhat
    If Not toWhat Then
        OD1.BackColor = 8421504
        OD2.BackColor = 8421504
        OD3.BackColor = 8421504
        OD4.BackColor = 8421504
        WALL1.BackColor = 8421504
        WALL2.BackColor = 8421504
        WALL3.BackColor = 8421504
        WALL4.BackColor = 8421504
        WALL5.BackColor = 8421504
        WALL6.BackColor = 8421504
        WALL7.BackColor = 8421504
        WALL8.BackColor = 8421504
        bow.BackColor = 8421504
    Else
        OD1.BackColor = -2147483643
        OD2.BackColor = -2147483643
        OD3.BackColor = -2147483643
        OD4.BackColor = -2147483643
        WALL1.BackColor = -2147483643
        WALL2.BackColor = -2147483643
        WALL3.BackColor = -2147483643
        WALL4.BackColor = -2147483643
        WALL5.BackColor = -2147483643
        WALL6.BackColor = -2147483643
        WALL7.BackColor = -2147483643
        WALL8.BackColor = -2147483643
        bow.BackColor = -2147483643
    End If
    
End Sub
Private Function checkInputs(palletCard As Control, itemNum As Control) As Boolean
    Dim result As Boolean
    If OD1.Text = "" Or Not IsNumeric(OD1.Text) Then
        MsgBox "Please complete all OD measurements.", vbOKOnly
        OD1.SetFocus
        result = False
    ElseIf OD2.Text = "" Or Not IsNumeric(OD2.Text) Then
        MsgBox "Please complete all OD measurements.", vbOKOnly
        OD2.SetFocus
        result = False
    ElseIf OD3.Text = "" Or Not IsNumeric(OD3.Text) Then
        MsgBox "Please complete all OD measurements.", vbOKOnly
        OD3.SetFocus
        result = False
    ElseIf OD4.Text = "" Or Not IsNumeric(OD4.Text) Then
        MsgBox "Please complete all OD measurements.", vbOKOnly
        OD4.SetFocus
        result = False
    ElseIf WALL1.Text = "" Or Not IsNumeric(WALL1.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL1.SetFocus
        result = False
    ElseIf WALL2.Text = "" Or Not IsNumeric(WALL2.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL2.SetFocus
        result = False
    ElseIf WALL3.Text = "" Or Not IsNumeric(WALL3.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL3.SetFocus
        result = False
    ElseIf WALL4.Text = "" Or Not IsNumeric(WALL4.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL4.SetFocus
        result = False
    ElseIf WALL5.Text = "" Or Not IsNumeric(WALL5.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL5.SetFocus
        result = False
    ElseIf WALL6.Text = "" Or Not IsNumeric(WALL6.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL6.SetFocus
        result = False
    ElseIf WALL7.Text = "" Or Not IsNumeric(WALL7.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL7.SetFocus
        result = False
    ElseIf WALL8.Text = "" Or Not IsNumeric(WALL8.Text) Then
        MsgBox "Please complete all wall measurements.", vbOKOnly
        WALL8.SetFocus
        result = False
    ElseIf bow.Text = "" Or Not IsNumeric(bow.Text) Then
        MsgBox "Please enter the bow measurement.", vbOKOnly
        bow.SetFocus
        result = False
    ElseIf palletCard.Text = "" Then
        MsgBox "Please enter the pallet card number.", vbOKOnly
        palletCard.SetFocus
        result = False
    ElseIf operator.Text = "" Then
        MsgBox "Please enter something to identify the operator.", vbOKOnly
        operator.Value = Environ$("UserName")
        operator.SetFocus
        result = False
    ElseIf itemNum.Text = "" Then
        MsgBox "Please enter the item number.", vbOKOnly
        itemNum.SetFocus
        result = False
    ElseIf TextBox26.Text = "" Then
        MsgBox "Please enter the exact length of this tube.", vbOKOnly
        TextBox26.SetFocus
        result = False
    ElseIf ComboBox1.Visible Then
        If IsNull(ComboBox1) Then
            MsgBox "Please select a scrap reason.", vbOKOnly
            ComboBox1.SetFocus
            result = False
        Else
            result = True
        End If
    Else
        result = True
    End If
    
    checkInputs = result
End Function
' This subroutine updates a spreadsheet with dimensional data, timestamps,
' operator information, and other relevant details.
'
' Args:
'   whatRow: The row number in the spreadsheet to update.
'   whatPCard: The process card identifier.
'   whatProcOrder: The process order number.
Private Sub updSpreadsheetBeforeSAP(whatRow As Long, ByVal palletCard As Long, ByVal itemNum As Long, elapsedTime As Variant)
', ByVal whatProcOrder As String
  On Error GoTo ErrHandler ' Error handling

  With localWS
    .Cells(whatRow, 1).Value = furnName      ' Furnace name
    .Cells(whatRow, 2).Value = Now()         ' Current date and time
    .Cells(whatRow, 3).Value = palletCard      ' Pallet card ID
    .Cells(whatRow, 4).Value = operator.Text  ' Operator name
'    .Cells(whatRow, 5).Value = whatProcOrder  ' Process order number (string)
    .Cells(whatRow, 6).Value = itemNum  ' item number

    ' --- Dimensional data (rounded to 2 decimal places) ---
    .Cells(whatRow, 9).Value = Round(CDec(OD1), 2)  ' Outer diameter 1
    .Cells(whatRow, 10).Value = Round(CDec(OD2), 2)  ' Outer diameter 2
    .Cells(whatRow, 11).Value = Round(CDec(OD3), 2)  ' Outer diameter 3
    .Cells(whatRow, 12).Value = Round(CDec(OD4), 2)  ' Outer diameter 4

    .Cells(whatRow, 13).Value = Round(CDec(WALL1), 2) ' Wall thickness 1
    .Cells(whatRow, 14).Value = Round(CDec(WALL2), 2)  ' Wall thickness 2
    .Cells(whatRow, 15).Value = Round(CDec(WALL3), 2)  ' Wall thickness 3
    .Cells(whatRow, 16).Value = Round(CDec(WALL4), 2)  ' Wall thickness 4
    .Cells(whatRow, 17).Value = Round(CDec(WALL5), 2)  ' Wall thickness 5
    .Cells(whatRow, 18).Value = Round(CDec(WALL6), 2)  ' Wall thickness 6
    .Cells(whatRow, 19).Value = Round(CDec(WALL7), 2)  ' Wall thickness 7
    .Cells(whatRow, 20).Value = Round(CDec(WALL8), 2)  ' Wall thickness 8

    .Cells(whatRow, 21).Value = bow.Value    ' Bow measurement
    .Cells(whatRow, 36).Value = TextBox26.Value 'length
    .Cells(whatRow, 26).Value = Round(CDec(elapsedTime), 2)   ' Hours
    'The following makes the pertinent spreadsheet info visible to the operator
    .Activate
    .Cells(whatRow + 50, 4).Select  ' Select a cell 50 rows down
    .Cells(whatRow, 4).Select     ' Select the original cell
  End With

  Exit Sub ' Normal exit

ErrHandler:
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in updSpreadsheet sub of UserForm1"
End Sub
Private Sub clearForm()
    ' This subroutine clears all the input fields on the form.

    OD1.Text = "" ' Clear OD1 field
    OD1.BackColor = RGB(192, 224, 255)
    OD2.Text = "" ' Clear OD2 field
    OD2.BackColor = RGB(192, 224, 255)
    OD3.Text = "" ' Clear OD3 field
    OD3.BackColor = RGB(192, 224, 255)
    OD4.Text = "" ' Clear OD4 field
    OD4.BackColor = RGB(192, 224, 255)
    WALL1.Text = "" ' Clear WALL1 field
    WALL1.BackColor = RGB(255, 192, 204)
    WALL2.Text = "" ' Clear WALL2 field
    WALL2.BackColor = RGB(255, 192, 204)
    WALL3.Text = "" ' Clear WALL3 field
    WALL3.BackColor = RGB(255, 192, 204)
    WALL4.Text = "" ' Clear WALL4 field
    WALL4.BackColor = RGB(255, 192, 204)
    WALL5.Text = "" ' Clear WALL5 field
    WALL5.BackColor = RGB(255, 192, 204)
    WALL6.Text = "" ' Clear WALL6 field
    WALL6.BackColor = RGB(255, 192, 204)
    WALL7.Text = "" ' Clear WALL7 field
    WALL7.BackColor = RGB(255, 192, 204)
    WALL8.Text = "" ' Clear WALL8 field
    WALL8.BackColor = RGB(255, 192, 204)
    bow.Value = "" ' Clear BOW field
    NOTES.Text = "" ' Clear NOTES field
    ToggleButton1.Value = False
    Webb.Visible = False
    ScrapD.Visible = False
    Trash.Visible = False
    ComboBox1.Visible = False
    WebbItemCombo.Visible = False
    Label17.Visible = False
End Sub

Private Sub Trash_Click()
    WebbItemCombo.Visible = False
    Label89.Visible = False
End Sub

Private Sub UserForm_Activate()
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")
    Set dataWS = localWB.Sheets("Sheet2")
    Set appControl = localWB.Sheets("AppControl")
    localWB.RefreshAll
    If ThisWorkbook.Sheets("Sheet2").Range("AA1").Value = "Offline" Then
        ToggleButton2.Value = True
        ToggleButton2.Enabled = False
        onlineMode = False
    Else
        onlineMode = True
    End If
    ToggleButton1 = False
    Label16.Caption = "Furnace " & right(furnName, 2)
    operator.Value = Environ$("UserName")
    localWS.Activate
    MultiPage1.Value = 0
    setSandMeter
    tpt84130.Value = Round(dataWS.Cells(25, 12).Value, 2)
    tpt215.Value = Round(dataWS.Cells(25, 13).Value, 2)
    Label91.Caption = dataWS.Cells(27, 11).Value
    Label94.Caption = dataWS.Cells(27, 12).Value
    Label95.Caption = dataWS.Cells(27, 13).Value
    ODMax = dataWS.Cells(1, 24)
    ODMin = dataWS.Cells(2, 24)
    wallMax = dataWS.Cells(3, 24)
    wallMin = dataWS.Cells(4, 24)
    If appControl.Cells(2, 2).Value = False Then Webb.Enabled = False
    If appControl.Cells(3, 2).Value = False Then ScrapD.Enabled = False
    If appControl.Cells(4, 2).Value = False Then plate.Enabled = False
    
    
End Sub

Private Sub UserForm_Deactivate()
    Set localWS = Nothing
    Set localWB = Nothing
End Sub

Private Sub Webb_Click()
If Webb Then
    WebbItemCombo.Visible = True
    Label89.Visible = True
End If
End Sub
