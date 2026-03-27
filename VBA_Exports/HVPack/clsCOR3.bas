VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCOR3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cDesc As String
Dim cBatch As String
Dim cItemNum As Long
Dim cProcOrdNum As String
Dim cRecipe As String

Sub getDetails(procOrdNum As String)  ' given the process order number, return the item number and description
    Dim xSession As GuiSession
    Set xSession = SAPSession.CurrentSession
    cProcOrdNum = procOrdNum
'    setStartScreen
    With xSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncor3"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = cProcOrdNum
        .FindById("wnd[0]").SendVKey 0
        If SAPStatus.CurrentStatus <> "" Then
      '      "Order " & cProcOrdNum & "   not found (check entry)" Then
            cDesc = SAPStatus.CurrentStatus
            MsgBox SAPStatus.CurrentStatus, vbOKOnly
            Exit Sub
        End If
        cItemNum = CLng(.FindById("wnd[0]/usr/txtCAUFVD-MATNR").Text)
        cDesc = .FindById("wnd[0]/usr/txtCAUFVD-MATXT").Text
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
        cBatch = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/ctxtAFPOD-CHARG").Text
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpSLAP").Select
        cRecipe = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpSLAP/ssubSUBSCR_5115:SAPLCOKO:5250/ctxtCAUFVD-PLNNR").Text
        .FindById("wnd[0]").SendVKey 3
        .FindById("wnd[0]").SendVKey 3
    End With
    Set xSession = Nothing
End Sub

Public Property Get itemNum() As Long
On Err GoTo ErrHandler
    itemNum = cItemNum
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get itemNum of clsCOR3."
End Property
Public Property Get itemDesc() As String
On Err GoTo ErrHandler
    itemDesc = cDesc
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get itemDesc of clsCOR3."
End Property
Public Property Get recipe() As String
On Err GoTo ErrHandler
    recipe = cRecipe
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get recipe of clsCOR3."
End Property
Public Property Get batch() As String
On Err GoTo ErrHandler
    batch = cBatch
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get batch of clsCOR3."
End Property

