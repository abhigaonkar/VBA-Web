VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CurrentMethod 
   Caption         =   "Cut and Wash Feedstock Form"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   11450
   OleObjectBlob   =   "CurrentMethod.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CurrentMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cutItemDescription As String
Private Sub bakedItemNum_AfterUpdate()
If bakedItemNum <> "" And onlineMode Then
    setStatus "Looking up baked item details"
    Dim thisItem As New clsMM
    BakedDesc.Caption = thisItem.getDescription(bakedItemNum.Value)
    ComboBox1.Value = thisItem.ODNominal
    If Left(thisItem.Description, 3) = "219" Then
        ComboBox10.RowSource = updateBoxes("219")
    Else
        ComboBox10.RowSource = updateBoxes(ComboBox1.Text)
    End If
    Set thisItem = Nothing
    setStatus ""
End If
End Sub
Sub setStatus(newStatus As String)
    statusUpdates.Caption = newStatus
End Sub
Private Sub ComboBox1_AfterUpdate()
     ComboBox10.RowSource = updateBoxes(ComboBox1.Text)
End Sub
Private Function updateBoxes(fsOD As String) As String
    Select Case fsOD
        Case 208
            updateBoxes = "FEEDSTOCK208"
        Case 210
            updateBoxes = "FEEDSTOCK210"
        Case 219
            updateBoxes = "FEEDSTOCK219"
        Case 254
            updateBoxes = "FEEDSTOCK254"
        Case 267
            updateBoxes = "FEEDSTOCK267"
        Case 165
            updateBoxes = "FEEDSTOCK165"
        Case 160
            updateBoxes = "FEEDSTOCK160"
        Case 134
            updateBoxes = "FEEDSTOCK134"
        Case 129
            updateBoxes = "FEEDSTOCK129"
        Case 102
            updateBoxes = "FEEDSTOCK102"
        Case 110
            updateBoxes = "FEEDSTOCK110"
        Case 100
            If bakedItemNum = 2123986 Or bakedItemNum = 2104864 Or bakedItemNum = 2104863 Then
                updateBoxes = "FEEDSTOCK100"
            Else
                updateBoxes = "FEEDSTOCK88100"
            End If
        Case 97
            updateBoxes = "FEEDSTOCK97"
        Case 88
            updateBoxes = "FEEDSTOCK88"
        Case 74
            updateBoxes = "FEEDSTOCK74"
        Case 72
            updateBoxes = "FEEDSTOCK72"
        Case 65
            updateBoxes = "FEEDSTOCK65"
        Case 60
            updateBoxes = "FEEDSTOCK60"
        Case 46
            updateBoxes = "FEEDSTOCK46"
   End Select
End Function
Private Sub fillODs(ctlNum As Integer, ctlText As String)
    Dim x As Integer
    x = ctlNum + 1
    
    Do While x <= 9
        Me.Controls("ComboBox" & x).Value = ctlText
        Me.Controls("ComboBox" & x + 9).RowSource = updateBoxes(ctlText)
        x = x + 1
    Loop
End Sub
Private Sub fillLens(ctlNum As Integer, ctlText As String)
    Dim x As Integer
    x = ctlNum + 1
    
    Do While x <= 18
        Me.Controls("ComboBox" & x).Value = ctlText
        x = x + 1
    Loop
End Sub
Private Sub ComboBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 1, ComboBox1.Value
End Sub

Private Sub ComboBox10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 10, ComboBox10.Value
End Sub

Private Sub ComboBox11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 11, ComboBox11.Value
End Sub

Private Sub ComboBox12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 12, ComboBox12.Value
End Sub

Private Sub ComboBox13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 13, ComboBox13.Value
End Sub

Private Sub ComboBox14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 14, ComboBox14.Value
End Sub

Private Sub ComboBox15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 15, ComboBox15.Value
End Sub

Private Sub ComboBox16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 16, ComboBox16.Value
End Sub

Private Sub ComboBox17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillLens 17, ComboBox17.Value
End Sub


Private Sub ComboBox2_AfterUpdate()
    ComboBox11.RowSource = updateBoxes(ComboBox2.Text)
End Sub

Private Sub ComboBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 2, ComboBox2.Value
End Sub

Private Sub ComboBox3_AfterUpdate()
    ComboBox12.RowSource = updateBoxes(ComboBox3.Text)
End Sub

Private Sub ComboBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 3, ComboBox3.Value

End Sub

Private Sub ComboBox4_AfterUpdate()
    ComboBox13.RowSource = updateBoxes(ComboBox4.Text)
End Sub

Private Sub ComboBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 4, ComboBox4.Value

End Sub

Private Sub ComboBox5_AfterUpdate()
    ComboBox14.RowSource = updateBoxes(ComboBox5.Text)
End Sub

Private Sub ComboBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 5, ComboBox5.Value

End Sub

Private Sub ComboBox6_AfterUpdate()
    ComboBox15.RowSource = updateBoxes(ComboBox6.Text)
End Sub

Private Sub ComboBox6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 6, ComboBox6.Value

End Sub

Private Sub ComboBox7_AfterUpdate()
    ComboBox16.RowSource = updateBoxes(ComboBox7.Text)
End Sub

Private Sub ComboBox7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 7, ComboBox7.Value

End Sub

Private Sub ComboBox8_AfterUpdate()
    ComboBox17.RowSource = updateBoxes(ComboBox8.Text)
End Sub

Private Sub ComboBox8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    fillODs 8, ComboBox8.Value

End Sub

Private Sub ComboBox9_AfterUpdate()
    ComboBox18.RowSource = updateBoxes(ComboBox9.Text)
End Sub
Private Function checkInputs() As Boolean
    Dim result As Boolean
    Dim x As Integer
    If bakedItemNum.Value = "" Then
        result = False
        MsgBox "Please enter a baked item number.", vbCritical
        bakedItemNum.SetFocus
    ElseIf bakedBatchNum.Value = "" Then
        result = False
        MsgBox "Please enter a baked batch number.", vbCritical
        bakedBatchNum.SetFocus
    ElseIf ComboBox10.Value = "" Then
        result = False
        MsgBox "Please select the length of the cut/washed item.", vbCritical
        ComboBox10.SetFocus
    Else
        result = True
    End If
    For x = 10 To 18
        If Controls("ComboBox" & x).Value <> "" Then
            If Not IsNumeric(Controls("TextBox" & x).Value) Then
                result = False
                MsgBox "Please enter a positive quantity for this length.", vbCritical
                Controls("TextBox" & x).SetFocus
    '            checkInputs = result
     '           Exit Function
            End If
        End If
    Next x
    checkInputs = result
End Function

Private Sub CommandButton1_Click()
'find out how many inputs
    Dim numBatches, x, recptAdj As Integer
    Dim startHere As Long
    Dim matlArray() As Variant
    Dim cutDown As Boolean
    Dim localWB As Workbook
    Dim localWS As Worksheet
    Dim t As Long
On Error GoTo ErrHandler
    If checkInputs Then
    '    ActiveWindow.WindowState = xlMinimized
        application.ScreenUpdating = False
        cutDown = False
        numBatches = tubeCount
        startHere = nextRow
        recptAdj = totalReceived - bakedQTY.Value
        If numBatches > 1 Then cutDown = True
        For x = 1 To numBatches
            Dim p, bakedItemNumber, processOrderNumber, cutItemNumber As Long
            p = nextRow
            With dataWS
          '      .Cells(p, 1).Select
                .Cells(p, 1).Value = Date
                .Cells(p, 2).Value = Now
                .Cells(p, 5).Value = BakedDesc.Caption
                .Cells(p, 4).Value = bakedItemNum.Value
                .Cells(p, 3).Value = Environ(Expression:="Username")
                .Cells(p, 6).Value = UCase(bakedBatchNum.Value)
                .Cells(p, 16).Value = False
                If recptAdj = 0 Then
                    .Cells(p, 7).Value = Me.Controls("TextBox" & x + 9).Value
                ElseIf recptAdj < 0 Then ' if baked > cut
                    If x = 1 Then
                        .Cells(p, 7).Value = Me.Controls("TextBox" & x + 9).Value - recptAdj
                    Else
                        .Cells(p, 7).Value = Me.Controls("TextBox" & x + 9).Value
                    End If
                ElseIf recptAdj > 0 Then  'if baked < cut
                    If x = 1 Then
                        .Cells(p, 7).Value = Me.Controls("TextBox" & x + 9).Value - recptAdj
                    Else
                        .Cells(p, 7).Value = Me.Controls("TextBox" & x + 9).Value
                    End If
                End If
                .Cells(p, 12).Value = Me.Controls("TextBox" & x + 9).Value
                .Cells(p, 9).Value = Me.Controls("ComboBox" & x + 9).Value
                .Cells(p, 11).Value = UCase(bakedBatchNum)
                .Cells(p, 13).Value = "No"
            End With
            setStatus "Saving data locally"
            ActiveWorkbook.Save
            setStatus ""
            If x = 1 And ComboBox19 <> "" Then
                dataWS.Cells(p, 14).Value = ComboBox19.Value
                dataWS.Cells(p, 15).Value = 1
            End If
            
            If CheckBox1 Then
                fastTrackIt Me.Controls("ComboBox" & x + 9).Value
            End If
        Next x
        setStatus "Saving locally again"
        ActiveWorkbook.Save
        setStatus ""
        For t = startHere To p
            If onlineMode Then
                setStatus "Saving to SAP"
                saveToSAP t
                setStatus ""
            End If
        Next t
        setStatus "Saving the backup"
        saveRemoteData startHere, p
        setStatus ""
        ActiveWindow.WindowState = xlMaximized
        If MsgBox("The transaction was successful. Would you like to clear the form?", vbYesNo) = vbYes Then CommandButton2_Click
        
        application.ScreenUpdating = True
    End If
    Exit Sub
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in CommandButton1_Click of form CurrentMethod"

End Sub
Sub fastTrackIt(ByVal cwItemNum As Long)
    setStatus "Printing Fast Track labels"
    Dim FT As Worksheet
    Set FT = thisWB.Sheets("FastTrack")
    FT.Activate
    FT.Range("e" & 2).Value = Environ(Expression:="Username")
    FT.Range("e" & 28).Value = Environ(Expression:="Username")
    If BakedDesc.Caption <> "" Then
        FT.Range("e" & 3).Value = BakedDesc.Caption
        FT.Range("e" & 29).Value = BakedDesc.Caption
    Else
        FT.Range("e" & 3).Value = "Offline" 'BakedDesc.Caption
        FT.Range("e" & 29).Value = "Offline"
    End If
    FT.Range("e" & 4).Value = bakedItemNum.Value
    FT.Range("e" & 30).Value = bakedItemNum.Value
    If onlineMode Then
        Dim cwDesc As New clsMM
        cwDesc.getDescription cwItemNum
        FT.Range("e" & 5).Value = cwDesc.Description
        FT.Range("e" & 31).Value = cwDesc.Description
        Set cwDesc = Nothing
    Else
        FT.Range("e" & 5).Value = "" 'cutItemDescription
        FT.Range("e" & 31).Value = "" 'cutItemDescription
    End If
    FT.Range("e" & 6).Value = cwItemNum
    FT.Range("e" & 32).Value = cwItemNum
    FT.Range("e" & 7).Value = UCase(bakedBatchNum)
    FT.Range("e" & 33).Value = UCase(bakedBatchNum)
    PrintFastTrackData 2
    setStatus ""

End Sub
Sub PrintFastTrackData(numCopies As Integer)

    ' Specify the range to print
    Dim printRange As Range
    Dim myPrinter As String
    On Error GoTo ErrHandler
    If numCopies > 0 Then
        myPrinter = "HP OfficeJet Pro 8710 (Feedstock) on Ne03:"
        application.ActivePrinter = myPrinter
        Set printRange = ThisWorkbook.Sheets("FastTrack").Range("A1:H51") '
        ' Print the range
        printRange.PrintOut Copies:=numCopies ' Adjust copies as needed
    End If
    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in PrintFastTrackData of form CurrentMethod"

End Sub
Function tubeCount() As Integer
    Dim result As Integer
    Dim keepGoing As Boolean
    On Error GoTo ErrHandler
    keepGoing = True
    result = 0
    Do While keepGoing
        result = result + 1
        If Me.Controls("TextBox" & result + 9).Value = "" Then
            keepGoing = False
            
        End If
    Loop
    tubeCount = result - 1
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in tubeCount function of form CurrentMethod"

End Function
Function tubeCount2() As Integer
    Dim result As Integer
    Dim keepGoing As Boolean
    On Error GoTo ErrHandler
    keepGoing = True
    result = 0
    Do While keepGoing
        result = result + 1
        If Me.Controls("combobox" & result + 9).Value = "" Then
            keepGoing = False
        End If
    Loop
    tubeCount2 = result - 1
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in tubeCount2 function of form CurrentMethod"

End Function
Function totalReceived() As Integer
    Dim result, x As Integer
    On Error GoTo ErrHandler
    Dim keepGoing As Boolean
    keepGoing = True
    result = 0
    x = 10
    Do While keepGoing
        
        result = result + Me.Controls("TextBox" & x).Value
        x = x + 1
        If Me.Controls("TextBox" & x).Value = "" Then
            keepGoing = False
            
        End If
    Loop
    totalReceived = result
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in totalReceived function of form CurrentMethod"

End Function
Private Sub clearForm()
    Dim x As Integer
    For x = 1 To 18
        Me.Controls("comboBox" & x).Value = ""
        If x > 9 Then Me.Controls("textBox" & x).Value = ""
    Next x
    bakedItemNum.Value = ""
    bakedBatchNum.Value = ""
    bakedQTY.Value = ""
    ComboBox19.Value = ""
    TextBox19.Value = ""
    CheckBox1.Value = False
End Sub
Private Sub CommandButton2_Click()
clearForm
bakedItemNum.SetFocus
End Sub

Private Sub CommandButton3_Click()
CurrentMethod.Hide
UserForm1.Show
End Sub

Private Sub CommandButton4_Click()
    If bakedItemNum.Value = "" Then
        MsgBox "Please enter an item number in the baked item number field.", vbOKOnly
        bakedItemNum.SetFocus
    Else
        getProcOrdNumXX bakedItemNum.Value, True, True
    End If
End Sub

Private Sub CommandButton5_Click()
    CurrentMethod.Hide
    frmLogPrint.Show
    ActiveWorkbook.Save
End Sub

Private Sub CommandButton6_Click()
    Me.Hide
    ActiveWorkbook.Save
End Sub

Private Sub CommandButton7_Click()
    Dim p, x As Integer
    p = tubeCount2
    If checkInputs Then
        For x = 1 To p
                fastTrackIt Me.Controls("combobox" & x + 9).Value
        Next x
    End If
    
End Sub

Private Sub UserForm_Activate()
    Set thisWB = ThisWorkbook
    Set dataWS = thisWB.Sheets("Sheet1")
    Set settingsWS = thisWB.Sheets("Sheet2")
    If settingsWS.Cells(1, 19).Value = "Online" Then
        onlineMode = True
    Else
        onlineMode = False
    End If


End Sub

