VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMB52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cInvArray() As Variant
Dim cSomeThere As Boolean
Dim cItemNum As String
Dim cSLoc As String
Dim cTotAtLoc As Variant


Sub getInventory(ByVal itemNum As String, ByVal sLoc As String)
    Dim thisSession As GuiSession
        Dim b As Integer
        Dim x As Integer
        cItemNum = itemNum
        cSLoc = sLoc
        Set thisSession = SAPSession.CurrentSession
        
        With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52" ' Open transaction MB52
        .FindById("wnd[0]").SendVKey 0 ' Enter
        
        ' Enter the sand item number, batch number, and storage location in SAP
        .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = cItemNum
        .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
        
   '     .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
   '     .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
   '     .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
        .FindById("wnd[0]/usr/radPA_FLT").Select
        .FindById("wnd[0]/usr/chkNOZERO").Selected = True
        
        .FindById("wnd[0]/usr/chkNEGATIV").Selected = False
        .FindById("wnd[0]/usr/chkXMCHB").Selected = True
        .FindById("wnd[0]/usr/chkNOVALUES").Selected = False
        .FindById("wnd[0]/usr/radPA_FLT").Select
        .FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = sLoc
          
        .FindById("wnd[0]").SendVKey 8
        cTotAtLoc = 0
        If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
            cSomeThere = False
            Exit Sub
        End If
        cSomeThere = True
        b = .FindById("wnd[0]/usr").Children.Count / 8 - 1

        ReDim cInvArray(b - 1, 3)
        
        ' Build an array with item availability information
        For x = 3 To 2 + b
            cInvArray(x - 3, 0) = cItemNum
            cInvArray(x - 3, 1) = .FindById("wnd[0]/usr/lbl[53," & x & "]").Text ' batch
            cInvArray(x - 3, 2) = .FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Location
            cInvArray(x - 3, 3) = .FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
            cTotAtLoc = cTotAtLoc + cInvArray(x - 3, 3)
        Next x
    End With ' thisSession
    Set thisSession = Nothing
End Sub
Property Get inventoryArray() As Variant
    inventoryArray = cInvArray
End Property
Property Get totalAtLocation() As Variant
    totalAtLocation = cTotAtLoc
    
End Property
Property Get isSomeThere() As Boolean
    isSomeThere = cSomeThere
End Property

