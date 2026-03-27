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
Dim clsProcessOrderNum As String
Dim clsItemNum As Long
Dim clsOrderQty As Variant
Dim clsCompletedQty As Variant
Dim clsBatch As String
Dim cDescription As String
Dim cUOM As String
Dim cPOType As String
Dim cMRPController As String


Function loadOrderInfo(procOrdNum As String) As Boolean
    Dim result As Boolean
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen
    result = True
    clsProcessOrderNum = procOrdNum
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "COR3"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = clsProcessOrderNum
    thisSession.FindById("wnd[0]").SendVKey 0
    If SAPStatus.CurrentStatus = "Order " & clsProcessOrderNum & "   not found (check entry)" Then
        MsgBox "Order " & clsProcessOrderNum & " not found (check entry)."
        result = False
        Exit Function
    End If
    clsItemNum = thisSession.FindById("wnd[0]/usr/txtCAUFVD-MATNR").Text
    clsOrderQty = thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").Text
    cUOM = thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/ctxtCAUFVD-GMEIN").Text
    clsCompletedQty = thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GWEMG").Text
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOAL").Select
    cMRPController = thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOAL/ssubSUBSCR_5115:SAPLCOKO:5140/txtT024D-DSNAM").Text
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
    clsBatch = thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/ctxtAFPOD-CHARG").Text
    cDescription = thisSession.FindById("wnd[0]/usr/txtCAUFVD-MATXT").Text
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing
    loadOrderInfo = result
End Function
Function detailsAsString() As Variant
    Dim meStr As String
    Dim tempArray As Variant
    ReDim tempArray(6)
    tempArray(0) = clsItemNum
    tempArray(1) = clsBatch
    tempArray(2) = cDescription
    tempArray(3) = clsOrderQty
    tempArray(4) = clsCompletedQty
    meStr = "Item Number:" & vbTab & clsItemNum & vbCrLf
    meStr = meStr & "Batch:" & vbTab & vbTab & clsBatch & vbCrLf
    meStr = meStr & "Description:" & vbTab & cDescription & vbCrLf
    meStr = meStr & "Order Qty:" & vbTab & clsOrderQty & " " & cUOM & vbCrLf
    meStr = meStr & "Completed:" & vbTab & clsCompletedQty
    tempArray(5) = meStr
    detailsAsString = tempArray
    
End Function

Property Let processOrderNum(ByVal procOrdNum As String) 'process order number
On Error GoTo ErrHandler

    clsProcessOrderNum = procOrdNum

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let procOrdNum of clsCOR3."
    
    
End Property
Public Property Get POType() As String
    Dim totLen As Long
    Dim myLen As Long
    totLen = Len(cMRPController)
    myLen = totLen - InStr(1, cMRPController, "-") - 1
    POType = right(cMRPController, myLen)
End Property

Public Property Get itemNum() As Long
    itemNum = clsItemNum
End Property
Public Property Get orderQty() As Variant
    orderQty = Round(CDec(clsOrderQty), 2)
End Property
Public Property Get completedQty() As Variant
    completedQty = Round(CDec(clsCompletedQty), 2)
End Property
Public Property Get Batch() As String
    Batch = clsBatch
End Property
Public Property Get Description() As String
    Description = cDescription
End Property
