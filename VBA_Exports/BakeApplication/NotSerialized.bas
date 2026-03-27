VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NotSerialized 
   Caption         =   "Non-Serialized Bake Transaction"
   ClientHeight    =   9225.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10470
   OleObjectBlob   =   "NotSerialized.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NotSerialized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bakedItemNumber As Long
Dim drawItemNum As Long
Dim batchNum As String
Dim bakedQty, unbakedQty, fty As Variant
Dim bakeTime As Variant
Dim oven As String
Dim oper As String
Dim comments As String
Dim hamperNum As Long
Dim hamperQty As Integer
Dim palletCard As Long
Dim procOrdNumStr As String
Dim bakedDesc As String
Dim unBakedDesc As String
Dim hrsPerEach As Variant
Dim inSAP As Boolean
Dim sourceStockLoc As String
Dim UOM As String


Function CheckInputs() As Boolean
    CheckInputs = False
    If TextBox1 = "" Then
        MsgBox "Please enter a pallet card number.", vbOKOnly
        TextBox1.SetFocus
    ElseIf TextBox2 = "" Then
        MsgBox "Please enter the bake hours.", vbOKOnly
        TextBox2.SetFocus
    ElseIf TextBox5 = "" Then
        MsgBox "Please identify yourself.", vbOKOnly
        TextBox5.SetFocus
    ElseIf TextBox6 = "" Then
        MsgBox "Please enter the unbaked batch number.", vbOKOnly
        TextBox6.SetFocus
    ElseIf TextBox7 = "" Then
        MsgBox "Please enter the baked quantity.", vbOKOnly
        TextBox7.SetFocus
    ElseIf TextBox8 = "" Then
        MsgBox "Please enter the unbaked quantity.", vbOKOnly
        TextBox8.SetFocus
    ElseIf procOrdNum = "" Then
        MsgBox "Please enter a baked process order number.", vbOKOnly
        procOrdNum.SetFocus
    ElseIf UnBakedItemNum = "" Then
        MsgBox "Please enter an unbaked item number.", vbOKOnly
        UnBakedItemNum.SetFocus
    ElseIf IsNull(ComboBox1.Value) Then
        MsgBox "Please select the oven.", vbOKOnly
        ComboBox1.SetFocus
    ElseIf IsNull(ComboBox2.Value) Then
        If Not IsNull(ComboBox3.Value) Then
            MsgBox "Please fill out details for hamper one before filling out details for hamper 2.", vbOKOnly
            ComboBox3.Value = ""
            ComboBox2.SetFocus
        ElseIf MsgBox("Are you sure you didn't pack this in a hamper?", vbYesNo) = vbNo Then
            MsgBox "Please select the hamper(s) used.", vbOKOnly
            ComboBox2.SetFocus
        Else
            TextBox4 = ""
        End If
        CheckInputs = True
        assignValues
    Else
        CheckInputs = True
        assignValues
    End If
    
End Function
Private Sub assignValues()
'        bakedItemNumber = BakedItemNum.Value
        drawItemNum = UnBakedItemNum.Value
        batchNum = TextBox6.Value
        bakedQty = TextBox7.Value
        unbakedQty = TextBox8.Value
        bakeTime = TextBox2.Value
        oven = ComboBox1.Value
        oper = TextBox5.Value
        comments = TextBox3.Value
        If ComboBox2.Value <> "" Then
            hamperNum = ComboBox2.Value
            hamperQty = TextBox4.Value
        Else
            hamperQty = 0
            hamperNum = 0
        End If
        palletCard = TextBox1.Value
        hrsPerEach = bakeTime / bakedQty

End Sub
Sub setStatus(newStatus As String)
    StatusWindow.Caption = newStatus
End Sub

Private Sub ComboBox2_AfterUpdate()
    ComboBox3.Visible = True
    TextBox9.Visible = True
    Label24.Visible = True
    Label25.Visible = True
End Sub


Private Sub CommandButton1_Click()
    If CheckInputs Then
        printerName = ComboBox5.Value
        Dim x As Integer 'to hold row num
        If TextBox7 <> TextBox8 Then
            Dim msgStr As String
            msgStr = "Your unbaked quantity is " & TextBox8.Value & "." & vbCrLf _
             & "Your baked quantity is " & TextBox7.Value & "." & vbCrLf _
             & "Is this correct?"
            If MsgBox(msgStr, vbYesNo) = vbNo Then Exit Sub
        End If
        x = nextEmptyRow(localSheet, 1, 2)
         With localSheet
            .Cells(x, 1).Value = oven
            .Cells(x, 2).Value = Now()
            .Cells(x, 3).Value = palletCard
            .Cells(x, 4).Value = oper
            .Cells(x, 5).Value = CLng(procOrdNum)
            .Cells(x, 6).Value = bakedItemNumber
            .Cells(x, 7).Value = bakedDesc
            .Cells(x, 8).Value = UCase(batchNum)
            .Cells(x, 9).Value = bakeTime
            .Cells(x, 10).Value = False
            .Cells(x, 11).Value = hrsPerEach
            .Cells(x, 12).Value = CLng(UnBakedItemNum)
            .Cells(x, 13).Value = unBakedDesc
            .Cells(x, 14).Value = bakedQty
            .Cells(x, 15).Value = unbakedQty
            .Cells(x, 16).Value = comments
            .Cells(x, 17).Value = hamperNum
            .Cells(x, 18).Value = hamperQty
            .Cells(x, 19).Value = "0100"
            .Cells(x, 20).Value = 261
            If hamperNum <> 134974 And hamperNum <> 134975 And hamperNum <> 134976 And hamperNum <> 134977 And hamperNum <> 134978 And hamperNum <> 0 Then
                .Cells(x, 21).Value = "147664"
                .Cells(x, 22).Value = 24 * hamperQty
                .Cells(x, 23).Value = "0100"
                .Cells(x, 24).Value = 261
            Else
                If MsgBox("Did you make this (these) hamper(s) from a bottom and a top?", vbYesNo) = vbYes Then
                    Dim makeHampers As New clsNewCOR6
                    With makeHampers
                        .procOrdNum = getCtnProcOrdNums(hamperNum)(0)
                        .yield = hamperQty
                        .applyMatl = False
                        .saveIt
                    End With
                    Set makeHampers = Nothing
                End If
            End If
            .Cells(x, 26).Value = "Not Serialized"
            .Cells(x, 27).Value = numLbl1
            .Cells(x, 28).Value = QTYperLbl1
            .Cells(x, 29).Value = numLbl2
            .Cells(x, 30).Value = QTYperLbl2
            .Cells(x, 31).Value = numLbl3
            .Cells(x, 32).Value = QTYperLbl3
            If ComboBox3.Value <> "" Then
                .Cells(x, 35).Value = ComboBox3.Value
                .Cells(x, 36).Value = TextBox9.Value
                .Cells(x, 37).Value = "0100"
                .Cells(x, 38).Value = 261
                .Cells(x, 22).Value = .Cells(x, 22).Value * 2
            End If
            If sourceStockLoc = "" Then
                sourceStockLoc = "0020"
            End If
            .Cells(x, 25).Value = sourceStockLoc
            setStatus "Saving locally, please wait."
            LogEvent "Saving pallet card: " & palletCard & " - process order: " & procOrdNum, "Transaction"
            ActiveWorkbook.Save
            setStatus "Saving a backup, please wait a little longer."
            saveRemoteData x, x
            
            If onlineMode Then
                setStatus "Saving SAP data"
                'check the qty of unbaked material against what was entered.
                saveToSAP x
            End If
            setStatus ""
         '   .Cells(x, 1).Select
        End With
        If MsgBox("Would you like to clear the form?", vbYesNo) = vbYes Then
            SaveAndCloseWorkbook
        Else
            TextBox6_AfterUpdate
        End If
    End If
    

End Sub
Private Sub CommandButton3_Click()
    Me.Hide
    BakeForm.Show
End Sub

Private Sub CommandButton4_Click()
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox4.Value = 1
    TextBox5.Value = ""
    TextBox6.Value = ""
    TextBox7.Value = ""
    TextBox8.Value = ""
    UnBakedItemNum.Value = ""
    procOrdNum.Value = ""
    ComboBox1.Value = ""
    ComboBox2.Value = ""
End Sub


Private Sub CommandButton5_Click()
    Dim p As Integer
    printerName = ComboBox5.Value
    If bakedItemNumber <> 0 Then
        For p = 1 To 3
            If Me.Controls("numLbl" & p).Text <> "" And Me.Controls("QTYperLbl" & p).Value <> "" Then
                setStatus "Printing Labels"
                printLabel procOrdNum, bakedItemNumber, batchNum, Me.Controls("numLbl" & p).Value, Me.Controls("QTYperLbl" & p).Value, UOM
                setStatus ""
            End If
        Next p
    Else
        MsgBox "Please enter a process order number, so I can lookup the details.", vbOKOnly
        procOrdNum.SetFocus
    End If
End Sub


Private Sub CommandButton6_Click()
NotSerialized.Hide
SandForm.Show

End Sub

Private Sub CommandButton7_Click()
    Dim bakedItemNumber As Long
    Dim thisUOM As String
    Dim x As Integer
    Dim batchNum As String
    printerName = ComboBox5
    setStatus "Looking up order information"
    Dim thisItem As New clsCOR3
    thisItem.getDetails procOrdNum
    bakedItemNumber = thisItem.itemNum
    batchNum = thisItem.batch
    thisUOM = thisItem.UOM
    setStatus ""
    Set thisItem = Nothing
    For x = 1 To 3
        If Controls("numLbl" & x).Value <> "" And Controls("qtyPerLbl" & x).Value <> "" Then
            If MsgBox("You are about to print " & Controls("numLbl" & x).Value & " labels with QTY " & Controls("qtyPerLbl" & x).Value & " each of " & bakedItemNumber & ". Is this correct?", vbYesNo) = vbYes Then
                setStatus "Printing"
                printLabel procOrdNum, bakedItemNumber, batchNum, Controls("numLbl" & x).Value, Controls("qtyPerLbl" & x).Value, thisUOM
            End If
        End If
    Next x
    setStatus ""
End Sub

Private Sub OfflineModeButton_Click()
    If onlineMode Then
        OfflineModeButton.Caption = "Offline Mode"
        Me.BackColor = vbYellow
        onlineMode = False
'        ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Offline"
    Else
        setStartScreen
        OfflineModeButton.Caption = "Online Mode"
        Me.BackColor = &H8000000F
        onlineMode = True
        ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Online"
    End If
'    ActiveWorkbook.Save

End Sub

Private Sub procOrdNum_AfterUpdate()
    If onlineMode Then
        setStatus "Looking up order information"
        Dim thisItem As New clsCOR3
        thisItem.getDetails procOrdNum
        bakedItemNumber = thisItem.itemNum
        bakedDesc = thisItem.itemDesc
        batchNum = thisItem.batch
        UOM = thisItem.UOM
        If right(batchNum, 1) = "F" Then
            TextBox7.Value = 4
            TextBox8.Value = 4
            setStatus "Fast Track"
        End If
        Label17.Caption = bakedDesc
        If thisItem.POType = "Unbaked" Then
            MsgBox "This process order is a draw process order. Please change it to the baked process order number.", vbCritical
            procOrdNum.SetFocus
'        ElseIf thisItem.POType <> "Baked" Then
'            MsgBox "This doesn't appear to be a baked process order. It appears to be a " & thisItem.POType & " process order. Please check it.", vbCritical
'            procOrdNum.SetFocus
        End If
        Set thisItem = Nothing
        setStatus ""
    End If
End Sub




Private Sub TextBox6_AfterUpdate()
    Dim thisMatl() As Variant
    If UnBakedItemNum.Value <> "" And onlineMode And Len(TextBox6.Value) > 5 Then
        setStatus "Looking up batch information"
        thisMatl = findMaterial(UnBakedItemNum.Value, TextBox6.Value)
        TextBox8.Value = CInt(thisMatl(0, 2))
        fty = TextBox8.Value
        sourceStockLoc = thisMatl(0, 3)
        setStatus ""
    End If

End Sub


Private Sub TextBox7_Change()
QTYperLbl1.Value = TextBox7.Value
End Sub

Private Sub UnBakedItemNum_AfterUpdate()
    If onlineMode Then
        setStatus "Looking up item data. Please wait."
        Dim thisMaterial As New clsMaterial
        unBakedDesc = thisMaterial.getDescription(UnBakedItemNum.Value)
        unbakedItemDesc.Caption = unBakedDesc
        If thisMaterial.isBaked Then
            MsgBox "This is a BAKED item number. Please enter an UNBAKED item number.", vbCritical
            NotSerialized.BackColor = vbRed
        Else
            NotSerialized.BackColor = -2147483633
        End If
        Set thisMaterial = Nothing
        setStatus ""
    End If
End Sub

Private Sub UserForm_Activate()

    If ThisWorkbook.Sheets("Sheet2").Cells(1, 10).Value = "Online" Then
        onlineMode = True
    Else
        onlineMode = False
    End If


    TextBox5.Value = Environ$("UserName")
    TextBox5.Enabled = False
    
    Set localSheet = ThisWorkbook.Sheets("Sheet1")
    
    If Not onlineMode Then
        OfflineModeButton.Value = False
        OfflineModeButton.Caption = "Offline Mode"
        Me.BackColor = vbYellow
    Else
     '   setStartScreen
        OfflineModeButton.Value = True
        OfflineModeButton.Caption = "Online Mode"
        Me.BackColor = &H8000000F
    End If

End Sub

