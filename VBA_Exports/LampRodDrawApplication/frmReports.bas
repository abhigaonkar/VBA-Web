VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReports 
   Caption         =   "SAP Reports"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5190
   OleObjectBlob   =   "frmReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CommandButton1_Click()

    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen



    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "mb51"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "311"
    thisSession.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/txtUSNAM-LOW").Text = operator.Value
    thisSession.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/BASIC"
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Set thisSession = Nothing
    
End Sub

Private Sub CommandButton2_Click()
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen



    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "mb51"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtBWART-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/txtUSNAM-LOW").Text = operator.Value
    thisSession.FindById("wnd[0]/usr/radRFLAT_L").Select
    thisSession.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/BASIC"
    
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Set thisSession = Nothing

End Sub

Private Sub CommandButton3_Click()
Me.Hide
UserForm1.Show
End Sub

Private Sub CommandButton4_Click()
'cooispi confirmations
    Dim thisReport As New clsCOOISPI
    thisReport.confirmations
    Set thisReport = Nothing
End Sub

Private Sub CommandButton5_Click()
    Dim thisReport As New clsCOOISPI
    thisReport.goodsMovements
    Set thisReport = Nothing
End Sub

Private Sub UserForm_Initialize()
    operator.Value = Environ$("UserName")
End Sub
