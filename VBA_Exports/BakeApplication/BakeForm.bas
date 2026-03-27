VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BakeForm 
   Caption         =   "Vacuum Bake Transactions"
   ClientHeight    =   8235.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9710.001
   OleObjectBlob   =   "BakeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BakeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim wb As Workbook, wbData As Workbook ' Workbook objects
    Dim thisSheet As Worksheet ' Worksheet object for "BakeTrial1.xlsm"


Private Sub BakedItemNum_AfterUpdate()
    If onlineMode Then
        Dim thisMaterial As New clsMaterial
        bakedItemDesc.Caption = thisMaterial.getDescription(BakedItemNum.Value)
        Set thisMaterial = Nothing
    End If
End Sub


Private Sub Batch1_AfterUpdate()
    checkBatches
End Sub

Private Sub Batch1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim numTubes As Variant
    Dim rightDig, x, firstHamperQty As Integer
    If Batch1.Text = "" Then
        Batch1.Text = InputBox("What's the lowest batch number?")
    End If
    numTubes = CInt(InputBox("How many tubes are on this rack?"))
    If IsNull(numTubes) Then Exit Sub
    Select Case numTubes
        Case 2 To 6
            firstHamperQty = numTubes
        Case 7 To 9
            firstHamperQty = 4
        Case 10 To 16
            firstHamperQty = CInt(numTubes / 2)
    End Select
    rightDig = CInt(right(Batch1.Text, 4))
    For x = 2 To CInt(numTubes)
        Me.Controls("Batch" & x).Value = left(Batch1.Text, 5) & addZeros(rightDig + (x - 1))
        Me.Controls("ToggleButton" & x).Visible = True
        If x > firstHamperQty Then clickChange Me.Controls("ToggleButton" & x)
    Next x
    TextBox5.SetFocus
End Sub


Private Sub Batch10_afterupdate()
checkBatches
End Sub

Private Sub Batch11_afterupdate()
checkBatches
End Sub

Private Sub Batch12_afterupdate()
checkBatches
End Sub

Private Sub Batch13_afterupdate()
checkBatches
End Sub

Private Sub Batch14_afterupdate()
checkBatches
End Sub

Private Sub Batch15_afterupdate()
checkBatches
End Sub

Private Sub Batch16_afterupdate()
checkBatches
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

Private Sub Batch9_AfterUpdate()
checkBatches
End Sub
Private Sub SaveOfflineData(ByVal rowNum As Long, ByVal ctrlNum As Integer, ByVal bakeHours As Variant)
    With thisSheet
        .Cells(rowNum, 1).Value = ComboBox1.Value ' Write data to the spreadsheet
        .Cells(rowNum, 2).Value = Now()
        .Cells(rowNum, 3).Value = TextBox1.Value
        .Cells(rowNum, 4).Value = TextBox5.Value
        
     '   .Cells(rowNum, 5).Value = procOrdNum(ctrlNum - 1)
        .Cells(rowNum, 6).Value = CLng(BakedItemNum.Value)
       ' .Cells(rowNum, 7).Value = bakedItemDesc.Caption
        .Cells(rowNum, 8).Value = UCase(Me.Controls("Batch" & ctrlNum).Value)
        .Cells(rowNum, 9).Value = TextBox2.Value
        .Cells(rowNum, 11).Value = Round(bakeHours, 2)
        .Cells(rowNum, 14).Value = 1
        .Cells(rowNum, 15).Value = 1
        .Cells(rowNum, 10).Value = False
        .Cells(rowNum, 26).Value = "Serialized"
        .Cells(rowNum, 33).Value = Me.Controls("ToggleButton" & ctrlNum).Caption
    End With

End Sub



Private Sub CommandButton1_Click()
    ' This subroutine processes data from a form, performs calculations, and saves the data to
    ' multiple locations, including an Excel spreadsheet ("BakeTrial1.xlsm") and potentially SAP.
    ' It interacts with SAP using the clsCOR6 class and handles online and offline modes.

    ' --- Check input values ---
    If Not CheckInputs Then Exit Sub ' Exit the subroutine if the input validation fails (checkInputs function not defined here)
    LogEvent "Saving Serialized Load for pallet card " & TextBox1.Value & " baked item num " & BakedItemNum
    ' --- Declare variables ---
    Dim x As Integer, Y As Integer ' Loop counters
    Dim t As Long, startRow As Integer  ' Row counter for the spreadsheet
    Dim bakeHours As Variant ' Variable to store bake hours

    ' --- Prepare data for hamper and strapping ---
    Dim xtras() As Variant ' Array to store hamper and strapping details
    ReDim xtras(1, 1) ' Initialize the array with 2 rows and 2 columns
    xtras(0, 0) = ComboBox2.Value ' Store the value from ComboBox2
    xtras(0, 1) = TextBox4.Value ' Store the value from TextBox4
    xtras(1, 0) = "147664" ' Store a constant value
    xtras(1, 1) = 24 ' Store a constant value
    printerName = ComboBox5.Value
    ' --- Determine the number of batches ---
    ' Count the number of batches by checking the values in controls named "Batch1", "Batch2", etc.
    For x = 1 To 15
        If Me.Controls("Batch" & x).Value = "" Then
            Y = x - 1 ' Store the number of batches in y
            x = 15 ' Exit the loop
        End If
    Next x
    startRow = nextEmptyRow(thisSheet, 1, 2)
    ReDim procOrdNum(Y) ' Resize the procOrdNum array to the number of batches

    bakeHours = Round(TextBox2.Value / Y, 2) ' Calculate bake hours per batch

    ' --- Loop through batches and save data to spreadsheets ---
    For x = 1 To Y
        
        
        t = nextEmptyRow(thisSheet, 1, 2) ' Get the next blank row in the spreadsheet
        
        ' --- Apply hamper and strapping to the first batch ---
        If x = 1 Then
            thisSheet.Cells(t, 17).Value = ComboBox2.Value
            thisSheet.Cells(t, 18).Value = TextBox4.Value
            thisSheet.Cells(t, 19).Value = "0100"
            thisSheet.Cells(t, 20).Value = 261
            thisSheet.Cells(t, 21).Value = "147664"
            thisSheet.Cells(t, 22).Value = 24 * TextBox4.Value
            thisSheet.Cells(t, 23).Value = "0100"
            thisSheet.Cells(t, 24).Value = 261

        End If
        
        SaveOfflineData t, x, bakeHours
        If onlineMode Then saveToSAP t

    Next x
    If onlineMode Then
        doHumo startRow, t
    End If
    ActiveWorkbook.Save
    thisSheet.Cells(t, 2).Select
    saveRemoteData startRow, t
    ' --- Close the external workbook ---
    If MsgBox("This transaction is complete. Would you like to clear the form?", vbYesNo) = vbYes Then
        SaveAndCloseWorkbook
    End If
End Sub
Private Function CheckInputs() As Boolean
    Dim result As Boolean
    If TextBox1.Value = "" Then
        'pallet card
        MsgBox "Please enter the pallet card number.", vbOKOnly
        TextBox1.SetFocus
        result = False
    ElseIf TextBox2.Value = "" Then
        MsgBox "Please enter the total hours this load baked.", vbOKOnly
        TextBox2.SetFocus
        'baked hours
        result = False
    ElseIf TextBox4.Value = "" Then
        'pallet qty
        MsgBox "Please enter the number of hampers you used.", vbOKOnly
        TextBox4.SetFocus
        result = False
    ElseIf TextBox5.Value = "" Then
        'operator name
        MsgBox "Please enter the person loading the hamper.", vbOKOnly
        TextBox5.SetFocus
        result = False
    ElseIf BakedItemNum.Value = "" Then
        MsgBox "Please enter the baked item number.", vbOKOnly
        BakedItemNum.SetFocus
        result = False
    ElseIf ComboBox1.Value = "" Then
        MsgBox "Please select the oven this baked in.", vbOKOnly
        ComboBox1.SetFocus
        result = False
    ElseIf ComboBox2.Value = "" Then
        MsgBox "What hamper were these tubes packed in?", vbOKOnly
        ComboBox2.SetFocus
        result = False
    ElseIf Batch1.Value = "" Then
        Batch1.SetFocus
        MsgBox "There must be at least one tube on this load.", vbOKOnly
        result = False
    
    Else
        result = True
    End If
    CheckInputs = result
End Function
Sub checkBatches()
Dim x, Y As Integer
For x = 1 To 16
    If Me.Controls("Batch" & x).Value <> "" Then
        Y = 16
        Do While Y > 1
            If Y <> x Then
                If Me.Controls("Batch" & x).Value = Me.Controls("Batch" & Y).Value Then
                    MsgBox "Check batches.", vbCritical
                    Me.Controls("Batch" & Y).SetFocus
                    Exit Sub
                End If
            End If
            Y = Y - 1
        Loop
    End If
Next x

End Sub


Private Sub CommandButton2_Click()
    'clear the form
    Dim x As Integer
    clearBatches
    For x = 1 To 4
        If x = 1 Or x = 2 Or x = 4 Then Me.Controls("TextBox" & x).Value = ""
    Next x
    BakedItemNum.Value = ""
    ComboBox1.Value = ""
    ComboBox2.Value = ""
End Sub

Private Sub CommandButton3_Click()
    Me.Hide
    NotSerialized.Show
End Sub

Private Sub OfflineMode_Click()
    If onlineMode Then
        OfflineMode.Caption = "Offline Mode"
        Me.BackColor = vbYellow
        onlineMode = False
  '      ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Offline"
    Else
  '      setStartScreen
        OfflineMode.Caption = "Online Mode"
        Me.BackColor = &H8000000F
        onlineMode = True
        ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Online"
    End If
End Sub

Private Sub TextBox1_AfterUpdate()
    Dim theseBatches As New clsPalletCardData
    Dim x As Integer
    Dim numTubes As Integer
    Dim firstHamperQty As Integer
    clearBatches
    theseBatches.cardNum = TextBox1.Value
    theseBatches.ReadFilteredDataIntoArray
    numTubes = UBound(theseBatches.batchNumbers) + 1
    Select Case numTubes
        Case 2 To 6
            firstHamperQty = numTubes
        Case 7 To 9
            firstHamperQty = 4
        Case 10 To 16
            firstHamperQty = CInt(numTubes / 2)
    End Select
    For x = 1 To numTubes
        Me.Controls("Batch" & x) = theseBatches.batchNumbers(x - 1)
        Me.Controls("ToggleButton" & x).Visible = True
        If x > firstHamperQty Then clickChange Me.Controls("ToggleButton" & x)
    Next x

    Set theseBatches = Nothing
End Sub

Private Sub clearBatches()
    Dim x As Integer
    For x = 1 To 16
        Me.Controls("Batch" & x).Value = ""
        Me.Controls("ToggleButton" & x).Visible = False
'        If x = 1 Or x = 2 Or x = 4 Or x = 5 Then Me.Controls("TextBox" & x).Value = ""
    Next x

End Sub

Private Sub ToggleButton1_Click()
    clickChange ToggleButton1
End Sub
Private Sub clickChange(thisButton As Control)
    If thisButton.Caption = 1 Then
        thisButton.Caption = 2
        thisButton.BackColor = vbRed
        thisButton.Value = 2
    Else
        thisButton.Caption = 1
        thisButton.BackColor = vbGreen
        thisButton.Value = 1

    End If

End Sub

Private Sub ToggleButton10_Click()
    clickChange ToggleButton10
End Sub

Private Sub ToggleButton11_Click()
    clickChange ToggleButton11
End Sub

Private Sub ToggleButton12_Click()
    clickChange ToggleButton12
End Sub

Private Sub ToggleButton13_Click()
    clickChange ToggleButton13
End Sub

Private Sub ToggleButton14_Click()
    clickChange ToggleButton14
End Sub

Private Sub ToggleButton15_Click()
    clickChange ToggleButton15
End Sub

Private Sub ToggleButton16_Click()
    clickChange ToggleButton16
End Sub

Private Sub ToggleButton2_Click()
    clickChange ToggleButton2
End Sub

Private Sub ToggleButton3_Click()
    clickChange ToggleButton3
End Sub

Private Sub ToggleButton4_Click()
    clickChange ToggleButton4
End Sub

Private Sub ToggleButton5_Click()
    clickChange ToggleButton5
End Sub

Private Sub ToggleButton6_Click()
    clickChange ToggleButton6
End Sub

Private Sub ToggleButton7_Click()
    clickChange ToggleButton7
End Sub

Private Sub ToggleButton8_Click()
    clickChange ToggleButton8
End Sub

Private Sub ToggleButton9_Click()
    clickChange ToggleButton9
End Sub

Private Sub UserForm_Activate()
    Set wb = ThisWorkbook ' Set the workbook object
    Set thisSheet = wb.Sheets("Sheet1") ' Set the worksheet object

    TextBox5.Value = Environ$("UserName")
    If Not onlineMode Then
        BakeForm.BackColor = vbYellow
        OfflineMode.Caption = "Offline Mode"
        OfflineMode.ControlTipText = "Click to switch to online mode."
    Else
        BakeForm.BackColor = &H8000000F
        OfflineMode.Caption = "Online Mode"
        OfflineMode.ControlTipText = "Click to switch to offline mode."
    End If

End Sub

