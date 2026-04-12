VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogPrint 
   Caption         =   "Other Transactions"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   14690
   OleObjectBlob   =   "frmLogPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    printDailyLog TextBox1.Value
    frmLogPrint.Hide
    CurrentMethod.Show
End Sub

Private Sub CommandButton2_Click()
    If checkInputs(False) Then
        MIGO551 itemNum.Value, batch.Value, qty.Value
    End If
End Sub

Private Sub CommandButton3_Click()
    If checkInputs(True) Then
        printLabel procOrdNum.Value, itemNum.Value, batch.Value, qty.Value
    End If
End Sub

Private Sub CommandButton4_Click()
    frmLogPrint.Hide
    CurrentMethod.Show
End Sub

Private Sub UserForm_Activate()
    TextBox1.Value = Format(DateAdd("d", -1, Now()), "mm/dd/yyyy")
    If settingsWS.Cells(1, 19).Value = "Online" Then
        onlineMode = True
    Else
        onlineMode = False
    End If
End Sub

Private Function checkInputs(procOrdNeeded As Boolean) As Boolean
    Dim result As Boolean
    result = True
    If itemNum.Value = "" Then
        MsgBox "Please enter an item number.", vbCritical
        itemNum.SetFocus
        result = False
    ElseIf batch.Value = "" Then
        MsgBox "Please enter a batch number.", vbCritical
        batch.SetFocus
        result = False
    ElseIf qty.Value = "" Then
        MsgBox "Please enter a quantity.", vbCritical
        qty.SetFocus
        result = False
    End If
    If procOrdNeeded Then
        If procOrdNum.Value = "" Then
            MsgBox "One moment while I lookup a process order for this.", vbCritical
            procOrdNum.Value = getProcOrdNumXX(itemNum, False, False)
            If procOrdNum.Value = 9 Then
                MsgBox "I can't find a valid process order number. Please enter any process order here.", vbCritical
                itemNum.SetFocus
                result = False
            End If
        End If
    End If
    checkInputs = result
End Function

