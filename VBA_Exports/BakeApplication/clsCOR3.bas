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
Dim cUOM As String
Dim cAmt As Variant
Dim cFinished As Variant
Dim cProcOrdNum As String
Dim cPOType As String
Dim cMRPController As String

Sub getDetails(procOrdNum As String)  ' given the process order number, return the item number and description
    Dim xSession As GuiSession
    Set xSession = SAPSession.CurrentSession
    cProcOrdNum = procOrdNum
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
        cUOM = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/ctxtCAUFVD-GMEIN").Text
        cAmt = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").Text
        cFinished = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GWEMG").Text
        cDesc = .FindById("wnd[0]/usr/txtCAUFVD-MATXT").Text
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOAL").Select
        cMRPController = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOAL/ssubSUBSCR_5115:SAPLCOKO:5140/txtT024D-DSNAM").Text
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
        cBatch = .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/ctxtAFPOD-CHARG").Text
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
Public Property Get POType() As String
    Dim totLen As Long
    Dim myLen As Long
    totLen = Len(cMRPController)
    myLen = totLen - InStr(1, cMRPController, "-") - 1
    POType = right(cMRPController, myLen)
End Property
Public Property Get UOM() As String
On Err GoTo ErrHandler
    UOM = cUOM
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get UOM of clsCOR3."
End Property
Public Property Get itemDesc() As String
On Err GoTo ErrHandler
    itemDesc = cDesc
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get itemDesc of clsCOR3."
End Property
Public Property Get batch() As String
On Err GoTo ErrHandler
    batch = cBatch
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get batch of clsCOR3."
End Property

