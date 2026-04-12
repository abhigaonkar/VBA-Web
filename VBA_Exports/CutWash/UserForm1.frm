VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9940.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8690.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_AfterUpdate()
    ComboBox10.RowSource = updateBoxes(ComboBox1.Text)
End Sub
Private Function updateBoxes(fsOD As String) As String
    Select Case fsOD
        Case 208
            updateBoxes = "FEEDSTOCK208"
        Case 210
            updateBoxes = "FEEDSTOCK210"
        Case 254
            updateBoxes = "FEEDSTOCK254"
        Case 267
            updateBoxes = "FEEDSTOCK267"
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


Private Sub CommandButton1_Click()
'find out how many inputs
    Dim numBatches As Integer, startingRow As Integer, x As Integer
'    Dim matlArray() As Variant
    Dim cutDown As Boolean
 '   Dim localWB As Workbook
  '  Dim localWS As Worksheet
   ' Dim BakedDesc As String
On Error GoTo ErrHandler
    numBatches = tubeCount
    'ActiveWindow.WindowState = xlMinimized
    startingRow = nextRow
    For x = 1 To numBatches
        Dim p, bakedItemNumber, processOrderNumber, cutItemNumber As Long
        p = nextRow
        With dataWS
            .Range("A" & p).Value = Date
            .Range("B" & p).Value = Now
            .Range("C" & p).Value = Environ(Expression:="Username")
            '.Range("E" & p).Value = bakedItemNumber
            .Range("F" & p).Value = UCase(Me.Controls("TextBox" & x).Value)
            .Range("G" & p).Value = 1
            .Range("I" & p).Value = Me.Controls("ComboBox" & x + 9).Value
            .Range("K" & p).Value = UCase(Me.Controls("TextBox" & x).Value)
            .Range("L" & p).Value = 1
            .Range("P" & p).Value = True
        End With
    Next x
    ActiveWorkbook.Save
    If onlineMode Then
        For x = startingRow To p
            saveToSAP x
        Next x
    End If
    saveRemoteData startingRow, p
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    MsgBox "This transaction is complete.", vbInformation
    Exit Sub
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in CommandButton1_Click of userForm1"
End Sub
Function tubeCount() As Integer
    Dim result As Integer
    Dim keepGoing As Boolean
    On Error GoTo ErrHandler
    keepGoing = True
    result = 0
    Do While keepGoing
        result = result + 1
        If Me.Controls("TextBox" & result).Value = "" Then
            keepGoing = False
        ElseIf result = 9 Then
            keepGoing = False
        End If
    Loop
    tubeCount = result - 1
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in tubeCount of userForm1"
    
End Function

Private Sub CommandButton3_Click()
UserForm1.Hide
CurrentMethod.Show
End Sub
Private Sub clearForm()
Dim x As Integer
For x = 1 To 19
    Me.Controls("comboBox" & x).Value = ""
    If x < 10 Then Me.Controls("textBox" & x).Value = ""
Next x
End Sub

Private Sub CommandButton4_Click()
clearForm
End Sub


Private Sub CommandButton5_Click()
    UserForm1.Hide
    frmLogPrint.Show
End Sub

Private Sub CommandButton6_Click()
    Me.Hide
    ActiveWorkbook.Save
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim numTubes As Variant
    Dim rightDig, x As Integer
    If TextBox1.Text = "" Then
        TextBox1.Text = InputBox("What's the lowest batch number?")
    End If
    numTubes = InputBox("How many tubes are you cutting?")
    If IsNull(numTubes) Then Exit Sub
    rightDig = CInt(Right(TextBox1.Text, 4))
    For x = 2 To CInt(numTubes)
        Me.Controls("TextBox" & x).Value = Left(TextBox1.Text, 5) & addZeros(rightDig + (x - 1))
    Next x
    TextBox5.SetFocus

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

