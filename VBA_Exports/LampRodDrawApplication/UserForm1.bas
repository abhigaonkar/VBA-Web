VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Non-Serialized Data Entry"
   ClientHeight    =   7290
   ClientLeft      =   80
   ClientTop       =   320
   ClientWidth     =   8730.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private frmItemInfo As String
Private frmOrderInfo As String
Private frmItemNum As Long
Private frmBatch As String
Private frmDesc As String
Private frmOrderQty As Variant
Private frmRemaining As Variant

Private Sub cmdMatMemoExists_Click()
    If MsgBox("This cannot be displayed here so I will take you to the memo in SAP. Is that OK?", vbYesNo) = vbNo Then Exit Sub
    Dim thisMaterial As New clsMaterial
    thisMaterial.showMemo frmItemNum
    Set thisMaterial = Nothing
End Sub

Private Sub cmdTransferSand_Click()

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
        lastSand = FindLastOccurrence(furnName, .Range("e:e")).Offset(, -4)
        lastBatch = FindLastOccurrence(furnName, .Range("e:e")).Offset(, -3)
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
            
'            .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
'            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
 '           .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
 '           .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = furnName
 '           .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "0400"
 '           .FindById("wnd[1]/tbar[0]/btn[8]").Press
             .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
             
            .FindById("wnd[0]").SendVKey 8
            If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
                MsgBox "There is no sand in inventory with this item number/batch combination.", vbOKOnly
                Exit Sub
            End If
            .FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status")
            If Not productionMode Then
                b = .FindById("wnd[1]/usr/lbl[17,8]").Text ' Extract the number of entries from the status window
            Else
                b = .FindById("wnd[1]/usr/lbl[17,10]").Text ' Extract the number of entries from the status window
            End If
            .FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window
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
     & "to " & vbTab & vbTab & vbTab & vbTab & furnName & vbCrLf _
     & "Is this correct?"
    If MsgBox(msgBoxTxt, vbYesNo) = vbYes Then
        LogEvent "Transferring " & transferAmt & " KG of " & sandItemNum & " batch: " & sandBatchNum & " to furnace " & furnName
        If transferMaterial(sandItemNum, sandBatchNum, transferAmt, fromLoc, furnName, True) Then
 '           Set thisSandLevelCheck = New clsFIFO ' Create an instance of clsFIFO
 '           thisSandLevelCheck.useMB52 furnName ' Update sand level using MB52 data (assuming this is a method of clsFIFO)
 '           Set thisSandLevelCheck = Nothing ' Release the object
        End If
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

Private Sub ComboBox1_AfterUpdate()
    furnName = ComboBox1.Value
    ComboBox1.Text = furnName
    equipName = "FR" & right(furnName, 2)
    CommandButton2_Click
    Dim furnNum As Integer
    furnNum = CInt(right(furnName, 2))
    If localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value <> "" Then
        frmItemInfo = localWB.Sheets("Sheet2").Cells(9, furnNum - 10).Value
        frmOrderInfo = localWB.Sheets("Sheet2").Cells(8, furnNum - 10).Value
'        procOrdNum.Text = localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value
    End If
    
End Sub

Private Sub ComboBox1_Change()
    Dim ctrl As MSForms.Control
    furnName = ComboBox1.Value
    ComboBox1.Text = furnName
    If getSandLevel > 0 Then
        ' Loop through each control in the UserForm's Controls collection
        For Each ctrl In Me.Controls
            ctrl.Enabled = True
        Next ctrl
        technician.Enabled = False
        palletCard.SetFocus
        furnName = ComboBox1.Value
        ComboBox1.Text = furnName
        equipName = "FR" & right(furnName, 2)
        CommandButton2_Click
        Dim furnNum As Integer
        furnNum = CInt(right(furnName, 2))
        If localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value <> "" Then
            frmItemInfo = localWB.Sheets("Sheet2").Cells(9, furnNum - 10).Value
            frmOrderInfo = localWB.Sheets("Sheet2").Cells(8, furnNum - 10).Value
        '      procOrdNum.Text = localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value
        End If
    Else
        For Each ctrl In Me.Controls
            ctrl.Enabled = False
        Next ctrl
        ComboBox1.Enabled = True
        cmdTransferSand.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim p As Long, y As Long
    Dim x As Integer, numScrapEntries As Integer, numTransactions As Integer
    numScrapEntries = 0
    
    If checkInputs Then
        LogEvent "Saving furnace number: " & furnName & " pallet card: " & palletCard.Text & " process order number: " & procOrdNum.Text
        'figure out how many scrap entries to determine the number of transactions
        printerName = ComboBox2.Value
        For x = 1 To 4
            If Controls("scrap" & x & "Reason").Text <> "" Then
                numScrapEntries = numScrapEntries + 1
            End If
        Next x
        If numScrapEntries = 0 Then
            numTransactions = 1
        Else
            numTransactions = numScrapEntries
        End If
        For x = 1 To numTransactions
            statusUpdate "Saving locally"
            Select Case x
                Case 1
                    p = nextEmptyRow(localWS, 2, 2)
                    y = p
                    localWS.Cells(p, 1).Value = furnName
                    localWS.Cells(p, 2).Value = Now()
                    localWS.Cells(p, 3).Value = palletCard.Text
                    localWS.Cells(p, 4).Value = technician.Text
                    localWS.Cells(p, 5).Value = procOrdNum.Text
                    localWS.Cells(p, 9).Value = QTY.Text
                    localWS.Cells(p, 10).Value = laborHours.Text
                    localWS.Cells(p, 11).Value = False
                    localWS.Cells(p, 12).Value = notes.Text
                    If Controls("scrap" & x & "Reason").Text <> "" Then
                        localWS.Cells(p, 13).Value = Controls("scrap" & x & "Reason").Column(1)
                        localWS.Cells(p, 14).Value = Controls("scrap" & x & "Amt").Text
                        If nonConform1 <> 0 Then
         '                   localWS.Cells(p, 12).Value = "Scrap D " & localWS.Cells(p, 12).Value
                            localWS.Cells(p, 18).Value = nonConform1.Value
                            localWS.Cells(p, 19).Value = localWS.Cells(p, 14).Value / (localWS.Cells(p, 14).Value + localWS.Cells(p, 9).Value)
                        End If
                    End If
                    statusUpdate "Saving a backup"
                    saveRemoteData p, False
                    statusUpdate ""
                Case 2, 3, 4
                    statusUpdate "Saving locally"
                    p = nextEmptyRow(localWS, 2, 2)
                    localWS.Cells(p, 1).Value = furnName
                    localWS.Cells(p, 2).Value = Now()
                    localWS.Cells(p, 3).Value = palletCard.Text
                    localWS.Cells(p, 4).Value = technician.Text
                    localWS.Cells(p, 5).Value = procOrdNum.Text
                    localWS.Cells(p, 9).Value = 0
                    localWS.Cells(p, 10).Value = 0
                    localWS.Cells(p, 11).Value = False
                    localWS.Cells(p, 12).Value = notes.Text
                    localWS.Cells(p, 13).Value = Controls("scrap" & x & "Reason").Column(1)
                    localWS.Cells(p, 14).Value = Controls("scrap" & x & "Amt").Text
                    If Controls("nonConform" & x) Then
              '          localWS.Cells(p, 12).Value = "Scrap D " & localWS.Cells(p, 12).Value
                        localWS.Cells(p, 18).Value = Controls("nonConform" & x).Value
                        localWS.Cells(p, 19).Value = 1
                    End If
                    statusUpdate "Saving a backup"
                    saveRemoteData p, False
                    statusUpdate ""
            End Select
        Next x
        statusUpdate "Saving entered data."
        ActiveWorkbook.Save
        If onlineMode Then
            statusUpdate "Entering SAP data."
            enterSavedData y, numTransactions
            Dim t As Long
            statusUpdate "Updating the backup"
            For t = y To p
                saveRemoteData t, True
            Next t
            statusUpdate ""
        End If
        
        statusUpdate "Closing"
        SaveAndCloseWorkbook
    End If
End Sub
Function checkInputs() As Boolean
    Dim result As Boolean
    Dim x As Integer
    statusUpdate "Checking Inputs"
    If ComboBox1.Text = "" Then
        MsgBox "Please select the furnace this is for.", vbCritical
        ComboBox1.SetFocus
        result = False
    ElseIf procOrdNum.Text = "" Then
        MsgBox "Please enter a valid process order number.", vbCritical
        procOrdNum.SetFocus
        result = False
    ElseIf laborHours.Text = "" Then
        MsgBox "Please enter the number of hours for this transaction. Hint: 0 is acceptable", vbCritical
        laborHours.SetFocus
        result = False
    ElseIf QTY.Text = "" Then
        MsgBox "Please enter the amount received for this transaction. Hint: 0 is acceptable", vbCritical
        QTY.SetFocus
        result = False
    ElseIf technician.Text = "" Then
        MsgBox "Please enter your name or number.", vbCritical
        technician.SetFocus
        result = False
    ElseIf palletCard.Text = "" Then
        MsgBox "Please enter the pallet card number.", vbCritical
        palletCard.SetFocus
        result = False
    Else
        result = True
    End If
    For x = 1 To 4
        If Controls("scrap" & x & "Reason").Text <> "" And Controls("scrap" & x & "Amt").Text = "" Then
            MsgBox "Please enter a scrap amount for this reason.", vbCritical
            Controls("scrap" & x & "Amt").SetFocus
            result = False
        End If
    Next x
    statusUpdate ""
    checkInputs = result
End Function
Private Sub CommandButton2_Click()
'clear the fields
    Dim x As Integer
    Label6.Caption = ""
    procOrdNum.Value = ""
    laborHours.Value = ""
    QTY.Text = ""
    For x = 1 To 4
        Controls("scrap" & x & "Amt").Text = ""
        Controls("scrap" & x & "Reason").Text = ""
    Next x
    palletCard.Value = ""
End Sub

Private Sub CommandButton3_Click()
    If procOrdNum = "" Then
        MsgBox "Please enter a valid process order number.", vbCritical
        procOrdNum.SetFocus
    Else
        If frmOrderInfo = "" And onlineMode Then
            statusUpdate "Looking up order data."
            Dim thisOrder As New clsCOR3
            thisOrder.loadOrderInfo procOrdNum
            Label6.Caption = thisOrder.detailsAsString(5)
            Label6.TextAlign = fmTextAlignLeft
            frmOrderInfo = Label6.Caption
            frmItemNum = thisOrder.detailsAsString(0)
            frmBatch = thisOrder.detailsAsString(1)
            frmDesc = thisOrder.detailsAsString(2)
            frmOrderQty = thisOrder.detailsAsString(3)
            frmRemaining = thisOrder.detailsAsString(4)
            Set thisOrder = Nothing
        Else
            Label6.Caption = frmOrderInfo
            Label6.TextAlign = fmTextAlignLeft
        End If
    End If
End Sub

Private Sub CommandButton4_Click()
    If procOrdNum = "" Then
        MsgBox "Please enter a valid process order number.", vbCritical
        procOrdNum.SetFocus
    Else
        If frmItemInfo = "" Then
            statusUpdate "Looking up data."
            Dim thisItem As New clsMaterial
            Dim thisOrder As New clsCOR3
            thisOrder.loadOrderInfo procOrdNum
            thisItem.getDescription thisOrder.itemNum
            Label6.Caption = thisItem.getSpecsAsString
            Label6.TextAlign = fmTextAlignLeft
            frmItemInfo = Label6.Caption
            Set thisOrder = Nothing
            Set thisItem = Nothing
        Else
            Label6.Caption = frmItemInfo
            Label6.TextAlign = fmTextAlignLeft
        End If
    End If
End Sub

Private Sub switchModes(online As Boolean)
    If online Then
        MsgBox "Please make sure all sand has been transferred to the furnace before doing draw transactions.", vbOKOnly
        CommandButton3.Visible = True
        CommandButton4.Visible = True
        Label10.Caption = "Online"
        Label10.ForeColor = vbGreen
        ThisWorkbook.Sheets("Sheet2").Cells(1, 5).Value = "Online"
        onlineMode = True
    Else
        CommandButton3.Visible = False
        CommandButton4.Visible = False
        Label6.Visible = False
        Label10.Caption = "Offline"
        Label10.ForeColor = vbRed
        ThisWorkbook.Sheets("Sheet2").Cells(1, 5).Value = "Offline"
        onlineMode = False
    End If
End Sub

Private Sub CommandButton5_Click()
    If MsgBox("This button is to save sand transfers to SAP that were done while SAP was offline. Is that what you want to do?", vbYesNo) = vbNo Then Exit Sub
    Dim startingRow As Long
    startingRow = nextRowWith(localWB.Sheets("SandTrans"), 6, "False")
    
    bulkSandTransfer startingRow
    ActiveWorkbook.Save
End Sub

Private Sub Label10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'switch modes
    If CommandButton4.Visible Then
        switchModes False
    Else
        switchModes True
    End If
End Sub

Private Sub procOrdNum_AfterUpdate()
'    frmOrderInfo = ""
    Label6.Caption = ""
    Dim tempArray As Variant
    ReDim tempArray(6)
    Dim furnNum As Integer
    furnNum = CInt(right(furnName, 2))

    If procOrdNum.Value <> "" Then
        If procOrdNum.Value = CStr(localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value) Then
            frmItemInfo = localWB.Sheets("Sheet2").Cells(9, furnNum - 10).Value
            frmOrderInfo = localWB.Sheets("Sheet2").Cells(8, furnNum - 10).Value
   '         procOrdNum.Text = localWB.Sheets("Sheet2").Cells(2, furnNum - 10).Value
        ElseIf onlineMode Then
            statusUpdate ("Looking up order data")
            Dim thisItem As New clsMaterial
            Dim thisOrder As New clsCOR3
            If thisOrder.loadOrderInfo(procOrdNum) Then
                thisItem.getDescription thisOrder.itemNum
                tempArray = thisOrder.detailsAsString
                tempArray(6) = thisItem.getSpecsAsString
                frmItemInfo = tempArray(6)
                Label6.Caption = tempArray(5)
                Label6.TextAlign = fmTextAlignLeft
                frmOrderInfo = tempArray(5)
                frmItemNum = tempArray(0)
                frmBatch = tempArray(1)
                frmDesc = tempArray(2)
                frmOrderQty = tempArray(3)
                frmRemaining = tempArray(4)
            End If
            If thisOrder.itemNum = 2107046 Then
                MsgBox "This is a startup order. Please enter all quantities in KGs.", vbOKOnly
            End If
            updateFurnaceData localWB.Sheets("Sheet2"), procOrdNum, tempArray, frmItemInfo, CInt(right(furnName, 2))
            
            If thisItem.materialMemoExists Then
                cmdMatMemoExists.Visible = True
            Else
                cmdMatMemoExists.Visible = False
            End If
            If thisOrder.POType <> "Unbaked" And thisOrder.POType <> "UNBAKED" And thisOrder.POType <> "UnBaked" Then
                MsgBox "This doesn't appear to be a draw process order. It appears to be a " & thisOrder.POType & " process order. Please check it.", vbCritical
                procOrdNum.SetFocus
            End If

            Set thisOrder = Nothing
            Set thisItem = Nothing
        End If
    End If
End Sub



Private Sub scrap1Amt_AfterUpdate()
    nonConform1.Visible = True
    Label12.Visible = True
End Sub


Private Sub scrap1Reason_AfterUpdate()
    Label12.Visible = True
End Sub


Private Sub scrap1Reason_Change()
    Label5.Visible = True
    scrap1Amt.Visible = True
    Label12.Visible = True
    nonConform1.Visible = True
    
End Sub

Private Sub scrap2Amt_AfterUpdate()
    nonConform2.Visible = True
End Sub
Private Sub scrap2Reason_AfterUpdate()
    nonConform2.Visible = True
    scrap2Amt.Visible = True
End Sub
Private Sub scrap3Amt_AfterUpdate()
    nonConform3.Visible = True
    scrap3Amt.Visible = True
End Sub

Private Sub scrap3Reason_AfterUpdate()
    nonConform3.Visible = True
    scrap3Amt.Visible = True
End Sub
Private Sub scrap4Amt_AfterUpdate()
    nonConform4.Visible = True
    scrap4Amt.Visible = True
End Sub


Private Sub scrap4Reason_AfterUpdate()
    nonConform4.Visible = True
    scrap4Amt.Visible = True
End Sub


Private Sub UserForm_Activate()
    Set localWB = ThisWorkbook
    Set localWS = ThisWorkbook.Sheets("Sheet1")
    If ThisWorkbook.Sheets("Sheet2").Cells(1, 5).Value = "Online" Then
        If onlineMode = False Then
            If setStartScreen Then switchModes True
        End If
    Else
        If onlineMode = False Then switchModes False
    End If
    If computerName = "MT-4223KD3" Then
        ComboBox2.Value = "HEBPRNP055"
    ElseIf computerName = "MT-4204KD3" Then
        ComboBox2.Value = "HEBPRNP056"
    Else
        ComboBox2.Value = "HEBPRNP054"
    End If
    Dim menuFile As String
    menuFile = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Furnace Deck\MainMenu.xlsm"
    If IsWorkBookOpen(menuFile) Then
        Workbooks(menuFile).Close SaveChanges:=False
    End If

    technician.Value = Environ("Username")
End Sub

