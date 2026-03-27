VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PackForm 
   Caption         =   "HV Pack Transactions"
   ClientHeight    =   8230.001
   ClientLeft      =   80
   ClientTop       =   320
   ClientWidth     =   7010
   OleObjectBlob   =   "PackForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim wb As Workbook, wbData As Workbook ' Workbook objects
    Dim thisSheet As Worksheet ' Worksheet object for "BakeTrial1.xlsm"

Private Sub Batch1_AfterUpdate()
    checkBatches
End Sub

Private Sub Batch1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim numTubes As Variant
    Dim rightDig, x As Integer
    If Batch1.Text = "" Then
        Batch1.Text = InputBox("What's the lowest batch number?")
    End If
    numTubes = InputBox("How many tubes are you packing?")
    If IsNull(numTubes) Then Exit Sub
    rightDig = CInt(Right(Batch1.Text, 4))
    For x = 2 To CInt(numTubes)
        Me.Controls("Batch" & x).Value = Left(Batch1.Text, 5) & addZeros(rightDig + (x - 1))
    Next x
End Sub

Private Sub makeHamperReady()
    Frame1.Visible = True
    Frame2.Visible = True
    ComboBox3.Visible = True
    Label25.Visible = True
    retQTY.Visible = True
    ComboBox2.Visible = True
    Label26.Visible = True
    TextBox6.Visible = True
    CommandButton1.Visible = True
    
End Sub

Private Sub Batch2_AfterUpdate()
checkBatches
End Sub

Private Sub Batch3_afterupdate()
checkBatches
End Sub

Private Sub Batch4_afterupdate()
checkBatches
End Sub


Private Sub Batch5_afterupdate()
checkBatches
End Sub

Private Sub Batch6_afterupdate()
checkBatches
End Sub

Private Sub Batch7_afterupdate()
checkBatches
End Sub

Private Sub Batch8_afterupdate()
checkBatches
End Sub

Private Sub SaveOfflineData(ByVal rowNum As Long, ByVal ctrlNum As Integer)
    With thisSheet
        .Cells(rowNum, 1).Value = Now()
        .Cells(rowNum, 2).Value = TextBox5.Value
        .Cells(rowNum, 4).Value = finItemNum.Value
        .Cells(rowNum, 6).Value = False
        .Cells(rowNum, 8).Value = UCase(Me.Controls("Batch" & ctrlNum).Value)
        .Cells(rowNum, 10).Value = 1
        .Cells(rowNum, 11).Value = 1
        .Cells(rowNum, 12).Value = comments
        .Cells(rowNum, 22).Value = Me.Controls("length" & ctrlNum).Value
        .Cells(rowNum, 23).Value = Me.Controls("bow" & ctrlNum & "a").Value
        .Cells(rowNum, 24).Value = Me.Controls("bow" & ctrlNum & "b").Value
    End With

End Sub


Private Sub bow1b_AfterUpdate()
    Complete1.Visible = True
    Complete1.SetFocus
End Sub

Private Sub bow2b_AfterUpdate()
    Complete2.Visible = True
    Complete1.Visible = False
    Complete2.SetFocus
End Sub
Private Sub bow3b_AfterUpdate()
    Complete3.Visible = True
    Complete2.Visible = False
    Complete3.SetFocus
End Sub
Private Sub bow4b_AfterUpdate()
    Complete4.Visible = True
    Complete3.Visible = False
    Complete4.SetFocus
End Sub
Private Sub bow5b_AfterUpdate()
    Complete5.Visible = True
    Complete4.Visible = False
    Complete5.SetFocus
End Sub
Private Sub bow6_AfterUpdate()
    Complete6.Visible = True
    Complete5.Visible = False
    Complete6.SetFocus
End Sub



Private Sub bow7b_AfterUpdate()
    Complete7.Visible = True
    Complete6.Visible = False
    Complete7.SetFocus
End Sub
Private Sub bow8b_AfterUpdate()
    Complete8.Visible = True
    Complete7.Visible = False
    Complete8.SetFocus
End Sub



Private Sub CommandButton1_Click()
    ' This subroutine processes data from a form, performs calculations, and saves the data to
    ' multiple locations, including an Excel spreadsheet and potentially SAP.
    ' It interacts with SAP using the clsCOR6 class and handles online and offline modes.

    ' --- Check input values ---
    If Not checkHamperInputs Then Exit Sub ' Exit the subroutine if the input validation fails (checkInputs function not defined here)
'    LogEvent "Saving finished item " & finItemNum.Value
    ' --- Declare variables ---
    Dim numGMs As Integer ' Loop counters
    Dim t As Long   ' Row counter for the spreadsheet
    Dim packaging(3, 4) As Variant
    Dim x As Integer, numBatches As Integer, numGoodsToIssue As Integer
    For x = 1 To 8
        If Me.Controls("Batch" & x).Value = "" Then
            numBatches = x
            x = 8
        End If
    Next x

    numGMs = 0
    ' --- Determine the number of batches ---
    ' Count the number of batches by checking the values in controls named "Batch1", "Batch2", etc.
    'find the last batch entry and add the hampers to it on the spreadsheet
    'then do COR6 to add packaging to that process order number
    t = nextEmptyRow(thisSheet, "A", 2) - 1
    If ComboBox2.Value <> "" Then
        numGMs = numGMs + 2
        thisSheet.Cells(t, 13).Value = ComboBox2.Value
        thisSheet.Cells(t, 14).Value = TextBox6.Value
        thisSheet.Cells(t, 15).Value = "0100"
        thisSheet.Cells(t, 16).Value = 261
'        packaging(0, 0) = thisSheet.Cells(t, 13).Value
'        packaging(0, 1) = thisSheet.Cells(t, 14).Value
'        packaging(0, 2) = thisSheet.Cells(t, 15).Value
'        packaging(0, 3) = thisSheet.Cells(t, 16).Value
        thisSheet.Cells(t, 17).Value = "147664"
        thisSheet.Cells(t, 18).Value = 24 * TextBox6.Value
        thisSheet.Cells(t, 19).Value = "0100"
        thisSheet.Cells(t, 20).Value = 261
'        packaging(1, 0) = thisSheet.Cells(t, 17).Value
'        packaging(1, 1) = thisSheet.Cells(t, 18).Value
'        packaging(1, 2) = thisSheet.Cells(t, 19).Value
'        packaging(1, 3) = thisSheet.Cells(t, 20).Value
    End If
    If ComboBox3.Value <> "" Then
        numGMs = numGMs + 1
        thisSheet.Cells(t, 26).Value = ComboBox3.Value
        thisSheet.Cells(t, 27).Value = retQTY.Value
        thisSheet.Cells(t, 28).Value = "0100"
        thisSheet.Cells(t, 29).Value = 531
'        packaging(2, 0) = thisSheet.Cells(t, 26).Value
'        packaging(2, 1) = thisSheet.Cells(t, 27).Value
'        packaging(2, 2) = thisSheet.Cells(t, 28).Value
'        packaging(2, 3) = thisSheet.Cells(t, 29).Value
    End If
    
    numGoodsToIssue = 1
    'if strapping completed, there's at least two items to issue
    If localSheet.Cells(t, 17).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 2
    'if return item completed there's one more to issue
    If localSheet.Cells(t, 26).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 1
    
    If numGoodsToIssue > 1 Then
        x = 1
        If thisSheet.Cells(t, 13).Value <> "" Then
            packaging(x, 0) = localSheet.Cells(t, 13).Value
            packaging(x, 1) = localSheet.Cells(t, 14).Value
            packaging(x, 2) = localSheet.Cells(t, 15).Value
            packaging(x, 3) = "" 'hamper batch
            packaging(x, 4) = localSheet.Cells(t, 16).Value 'hamper movement type
            x = x + 1
            packaging(x, 0) = localSheet.Cells(t, 17).Value 'strapping item num
            packaging(x, 1) = localSheet.Cells(t, 18).Value 'strapping qty
            packaging(x, 2) = localSheet.Cells(t, 19).Value 'strapping location
            packaging(x, 3) = "" 'strapping batch
            packaging(x, 4) = localSheet.Cells(t, 20).Value 'strapping movement type
            x = x + 1
        End If
        If localSheet.Cells(t, 26).Value <> "" Then
            packaging(x, 0) = localSheet.Cells(t, 26).Value 'return hamper item num
            packaging(x, 1) = localSheet.Cells(t, 27).Value 'return hamper qty
            packaging(x, 2) = localSheet.Cells(t, 28).Value 'return hamper location
            packaging(x, 3) = "" 'strapping batch
            packaging(x, 4) = localSheet.Cells(t, 29).Value 'return hamper movement type
            x = x + 1
        End If
        If MsgBox("Did you add any waffle boards?", vbYesNo) = vbYes Then
            Dim numWaffle As Integer
            numWaffle = InputBox("How many waffle board did you put in the hamper?", "Waffle Question")
            packaging(x, 0) = "124451" 'waffle item num
            packaging(x, 1) = numWaffle 'waffle qty
            packaging(x, 2) = "0100" 'waffle location
            packaging(x, 3) = "" 'waffle batch
            packaging(x, 4) = "261" 'waffle movement type
        End If
    End If
    
    
    updateStatus "Saving a backup"
    saveRemoteData t, t
    updateStatus ""
    If onlineMode Then
        Dim newCOR6 As New clsNewCOR6
        newCOR6.oper = 20
        newCOR6.applyNumMatl = numGMs
        newCOR6.matlArray = packaging
        newCOR6.procOrdNum = thisSheet.Cells(t, 3).Value
        newCOR6.yield = 0
        newCOR6.saveIt
        Set newCOR6 = Nothing
        
        updateStatus "Completing the handling unit"
        doHumo t - numBatches + 1, t
        updateStatus ""
    End If
    MsgBox "This transaction is complete."
    SaveAndCloseWorkbook
   
End Sub
Function saveThisLine(lineNo As Long) As Boolean
    Dim result, is224 As Boolean
    If Not CheckInputs(lineNo) Then Exit Function ' Exit the subroutine if the input validation fails (checkInputs function not defined here)
    LogEvent "Saving finished item " & finItemNum.Value
    ' --- Declare variables ---
    Dim x As Integer ' Loop counters
    Dim startRow As Integer   ' Row counter for the spreadsheet
    result = False
    startRow = nextEmptyRow(thisSheet, "A", 2)
    ReDim procOrdNum(1) ' Resize the procOrdNum array to the number of batches
    If lineNo = 1 Then
        If ComboBox2.Value <> "" Then
            thisSheet.Cells(startRow, 13).Value = ComboBox2.Value
            thisSheet.Cells(startRow, 14).Value = TextBox6.Value
            thisSheet.Cells(startRow, 15).Value = "0100"
            thisSheet.Cells(startRow, 16).Value = 261
            thisSheet.Cells(startRow, 17).Value = "147664"
            thisSheet.Cells(startRow, 18).Value = 24 * TextBox6.Value
            thisSheet.Cells(startRow, 19).Value = "0100"
            thisSheet.Cells(startRow, 20).Value = 261
        End If
        If ComboBox3.Value <> "" Then
            thisSheet.Cells(startRow, 26).Value = ComboBox3.Value
            thisSheet.Cells(startRow, 27).Value = retQTY.Value
            thisSheet.Cells(startRow, 28).Value = "0100"
            thisSheet.Cells(startRow, 29).Value = 531
        End If
    End If
    Select Case TextBox7.Text
        Case "FCE STD CHEMISTRY"
            is224 = False
        Case "FCE LOW ALKALI"
            is224 = True
        Case Else
            is224 = False
    End Select

' --- Loop through batches and save data to spreadsheets ---
    updateStatus "Saving data locally"
    ' --- Apply hamper and strapping to the first batch ---

    SaveOfflineData startRow, lineNo
    thisSheet.Cells(startRow, 2).Select
    updateStatus "Saving a backup"
    saveRemoteData startRow, startRow
    updateStatus ""
    If onlineMode Then
        updateStatus "Saving data to SAP"
        If saveToSAP(startRow) Then result = True
        updateStatus "Printing COA"
        printCOA thisSheet.Cells(startRow, 4).Value, thisSheet.Cells(startRow, 8).Value, is224
        updateStatus ""
    End If
    MsgBox "This transaction is complete."

   saveThisLine = result
End Function
Private Function CheckInputs(ByVal lineNo As Integer) As Boolean
    Dim result As Boolean
'    If TextBox6.Value = "" Then
        'pallet qty
'        MsgBox "Please enter the number of hampers you used.", vbOKOnly
'        TextBox6.SetFocus
'        result = False
'    ElseIf TextBox6.Value = "" Then
        'operator name
'        MsgBox "Please enter the person loading the hamper.", vbOKOnly
'        TextBox5.SetFocus
'        result = False
    If finItemNum.Value = "" Then
        MsgBox "Please enter the finished item number.", vbOKOnly
        finItemNum.SetFocus
        result = False
'    ElseIf ComboBox2.Value = "" Then
'        MsgBox "What hamper were these tubes packed in?", vbOKOnly
'        ComboBox2.SetFocus
'        result = False
    ElseIf Controls("Batch" & lineNo).Value = "" Then
        Controls("Batch" & lineNo).SetFocus
        MsgBox "Please enter a correct batch number.", vbOKOnly
        result = False
    Else
        result = True
    End If
    CheckInputs = result
End Function
Private Function checkHamperInputs() As Boolean
    
    If IsNull(ComboBox2.Value) Then
        If MsgBox("Were these tubes packed in a 'new' hamper?", vbYesNo) = vbYes Then
            ComboBox2.SetFocus
            checkHamperInputs = False
            MsgBox "Please select that hamper from this list.", vbOKOnly
            Exit Function
        Else
            checkHamperInputs = True
        End If
    Else
        If TextBox6.Value = "" Then
            'pallet qty
            MsgBox "Please enter the number of hampers you packed this in.", vbOKOnly
            TextBox6.SetFocus
            checkHamperInputs = False
            Exit Function
        Else
            checkHamperInputs = True
        End If
    End If
    If IsNull(ComboBox3.Value) Then
        If MsgBox("Are you returning an empty hamper to inventory?", vbYesNo) = vbYes Then
            ComboBox3.SetFocus
            checkHamperInputs = False
            MsgBox "Please select that hamper from this list.", vbOKOnly
            Exit Function
        Else
            checkHamperInputs = True
        End If
    Else
        If retQTY.Value = "" Then
            'pallet qty
            MsgBox "Please enter the number of hampers you are returning to inventory.", vbOKOnly
            retQTY.SetFocus
            checkHamperInputs = False
            Exit Function
        Else
            checkHamperInputs = True
        End If
    End If
End Function
Sub checkBatches()
Dim x, y As Integer
For x = 1 To 8
    If Me.Controls("Batch" & x).Value <> "" Then
        y = 8
        Do While y > 1
            If y <> x Then
                If Me.Controls("Batch" & x).Value = Me.Controls("Batch" & y).Value Then
                    MsgBox "Check batches.", vbCritical
                    Me.Controls("Batch" & y).SetFocus
                    Exit Sub
                End If
            End If
            y = y - 1
        Loop
    End If
Next x

End Sub


Private Sub makeDisabled(ByVal x As Integer)
    Me.Controls("Batch" & x).Enabled = False
    Me.Controls("length" & x).Enabled = False
    Me.Controls("bow" & x & "a").Enabled = False
    Me.Controls("bow" & x & "b").Enabled = False
    Me.Controls("Complete" & x).Enabled = False
End Sub
Private Sub makeVisible(ByVal x As Integer)
    Me.Controls("Batch" & x).Visible = True
    Me.Controls("length" & x).Visible = True
    Me.Controls("bow" & x & "a").Visible = True
    Me.Controls("bow" & x & "b").Visible = True
    Me.Controls("Batch" & x).SetFocus
End Sub
Private Sub CommandButton3_Click()
    SaveAndCloseWorkbook
End Sub

Private Sub Complete1_Click()
    
    If saveThisLine(1) Then
        makeDisabled 1
        makeVisible 2
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete2_Click()
    If saveThisLine(2) Then
        makeDisabled 2
        makeVisible 3
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete3_Click()
    If saveThisLine(3) Then
        makeDisabled 3
        makeVisible 4
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete4_Click()
    If saveThisLine(4) Then
        makeDisabled 4
        makeVisible 5
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete5_Click()
    If saveThisLine(5) Then
        makeDisabled 5
        makeVisible 6
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete6_Click()
    If saveThisLine(6) Then
        makeDisabled 6
        makeVisible 7
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete7_Click()
    If saveThisLine(7) Then
        makeDisabled 7
        makeVisible 8
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub Complete8_Click()
    If saveThisLine(8) Then
        makeDisabled 8
        If MsgBox("Is this the last tube for this hamper?", vbYesNo) = vbYes Then makeHamperReady
    End If
End Sub
Private Sub finItemNum_AfterUpdate()
    If onlineMode Then
        updateStatus "Looking up item info"
        Dim thisItem As New clsMaterial
        Label11.Visible = True
        finishedItemDescription.Caption = thisItem.getDescription(finItemNum)
        TextBox7.Text = thisItem.matlGrade
        Set thisItem = Nothing
        updateStatus ""
'        Batch1.SetFocus
    End If
    Label21.Visible = True
    Label13.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    makeVisible 1
    Batch1.SetFocus
End Sub


Private Sub OfflineMode_Click()
    If onlineMode Then
        OfflineMode.Caption = "Offline Mode"
        Me.BackColor = vbYellow
        onlineMode = False
        ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Offline"
    Else
        setStartScreen
        OfflineMode.Caption = "Online Mode"
        Me.BackColor = &H8000000F
        onlineMode = True
        ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Online"
    End If
End Sub



Private Sub UserForm_Activate()
    If ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Offline" Then
        onlineMode = False
    ElseIf ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Online" Then
        onlineMode = True
    End If
    TextBox5.Value = Environ$("UserName")
    Set wb = ThisWorkbook ' Set the workbook object
    Set thisSheet = wb.Sheets("Sheet1") ' Set the worksheet object
    If Not onlineMode Then
        PackForm.BackColor = vbYellow
        OfflineMode.Caption = "Offline Mode"
        OfflineMode.ControlTipText = "Click to switch to offline mode."
    Else
        PackForm.BackColor = &H8000000F
        OfflineMode.Caption = "Online Mode"
        OfflineMode.ControlTipText = "Click to switch to online mode."
    End If
    Dim x As Integer
    For x = 1 To 8
        Me.Controls("Batch" & x).TabIndex = x * 4 - 3
        Me.Controls("length" & x).TabIndex = x * 4 - 2
        Me.Controls("bow" & x & "a").TabIndex = x * 4 - 1
        Me.Controls("bow" & x & "b").TabIndex = x * 4
        
    Next x
End Sub

