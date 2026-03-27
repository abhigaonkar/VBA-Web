VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private cProcOrdNum As String
Private cHUNum As String
Private cBatch As String
Private cItemNum As Long
Private cHUBatch() As Variant
Sub cop1()
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOP1"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtPLAPPLDATA-AUFNR").Text = cProcOrdNum
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0052/subAREA2:SAPLVHUDIAL2:0061/subPACKDIALOG:SAPLVHUSUBSC:0100/btnCREATE_HUS").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisSession = Nothing
End Sub
Function cowbhuwe() As String
'GR for this order
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncowbhuwe"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = cProcOrdNum
        
        .FindById("wnd[0]/tbar[1]/btn[25]").Press
        .FindById("wnd[0]/usr/chkAFPO-ELIKZ").Selected = True
        .FindById("wnd[0]/usr/subHULIST:SAPLVHURMSUB:1000/subHULIST_TC:SAPLVHURMSUB:2100/tblSAPLVHURMSUBTC_HULIST_FAUF/ctxtVHURMHUD-CHARG[9,0]").Text = cBatch
        .FindById("wnd[0]/usr/subHULIST:SAPLVHURMSUB:1000/subHULIST_TC:SAPLVHURMSUB:2100/tblSAPLVHURMSUBTC_HULIST_FAUF/ctxtVHURMHUD-PRODDATE[11,0]").Text = Format(Date, "mm/dd/yyyy")
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        
        If printYes Then
        .FindById("wnd[1]/usr/btnDY_VAROPTION1").Press
            .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName
            
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
        Else
            .FindById("wnd[1]/usr/btnCANCEL").Press
        End If
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisSession = Nothing

End Function
Sub humo4Gr()
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    Dim x As Integer
    Dim scrollCount As Integer
    Dim rowNum As Integer
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nHUMO"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200").Select
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELWERK-LOW").Text = "q105"
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELMATNR-LOW").Text = cItemNum
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/btn%_SELCHARG_%_APP_%-VALU_PUSH").Press
        scrollCount = 1
        rowNum = 0
        For x = 0 To UBound(cHUBatch)
            If x > 7 Then
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").VerticalScrollbar.Position = scrollCount
                scrollCount = scrollCount + 1
                rowNum = rowNum - 1
            End If
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1," & rowNum & "]").Text = cHUBatch(x)
            rowNum = rowNum + 1
        Next x
        
        .FindById("wnd[1]/tbar[0]/btn[8]").Press
        
        .FindById("wnd[0]").SendVKey 8
        cHUNum = .FindById("wnd[0]/usr/lbl[6,4]").Text
'        .FindById("wnd[0]/usr/chk[1,4]").Selected = True
'        .FindById("wnd[0]/usr/chk[1,6]").Selected = True
        
        .FindById("wnd[0]/tbar[1]/btn[5]").Press
        .FindById("wnd[0]/tbar[1]/btn[18]").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT014").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").GetAbsoluteRow(0).Selected = False
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT003").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").GetAbsoluteRow(0).Selected = True
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_002").GetAbsoluteRow(0).Selected = False
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    
    If printYes Then printHULabel cHUNum, 2
    
    Set thisSession = Nothing

End Sub
Private Function lookupHU() As String

    Dim thisSession1 As GuiSession
    Dim result As String
    Set thisSession1 = SAPSession.CurrentSession
    With thisSession1
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nhumo"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200").Select
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELMATNR-LOW").Text = cItemNum
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/txtSELCHARG-LOW").Text = cBatch
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELWERK-LOW").Text = "q105"
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELLGORT-LOW").Text = ""
        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        .FindById("wnd[0]/usr/lbl[3,4]").SetFocus
        .FindById("wnd[0]").SendVKey 2
        cHUNum = .FindById("wnd[0]/usr/lbl[6,4]").Text
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    
    Set thisSession1 = Nothing
    lookupHU = cHUNum
End Function
Sub huGI(ByVal procOrdNum As String, ByVal itemNum As String, ByVal batchNum As String)
    Dim thisSession As GuiSession
    Dim x As Integer
    Dim t As Integer
    Set thisSession = SAPSession.CurrentSession
 '   setStartScreen
    x = 0
    cProcOrdNum = procOrdNum
    cItemNum = itemNum
    cBatch = UCase(batchNum)
    lookupHU
    With thisSession
    
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncor2"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = cProcOrdNum
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/ctxtAFPOD-CHARG").Text = cBatch
        .FindById("wnd[0]").SendVKey 0
        If SAPStatus.IsThereAModalWindow Then .FindById("wnd[1]/usr/btnSPOP-VAROPTION1").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncopawa"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = cProcOrdNum
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsSTRIP/tabpHUZU").Select
        .FindById("wnd[0]/usr/tabsSTRIP/tabpHUZU/ssubSUB1:SAPLCOMI:0103/ctxtEXIDV_ASS").Text = cHUNum

        .FindById("wnd[0]/usr/tabsSTRIP/tabpHUZU/ssubSUB1:SAPLCOMI:0103/btn%#AUTOTEXT001").Press
        Do While SAPStatus.IsThereAModalWindow
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
            x = x + 1
        Loop

        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncowbhuwa"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/subHEADER:SAPLCOWBHU:0120/ctxtCAUFVD-AUFNR").Text = cProcOrdNum

        .FindById("wnd[0]/tbar[1]/btn[25]").Press
        .FindById("wnd[0]/usr/subHULIST:SAPLCOWBHU:0130/btnTEIL_WA").Press
        For t = 0 To (x / 2) - 1
            If .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-CHARG[1," & t & "]").Text = cBatch Then
                .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-MENGE[2," & t & "]").Text = "1"
            Else
                .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-MENGE[2," & t & "]").Text = "0"
            End If

        Next t
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        PauseMe 1, False
        If x > 2 Then
            .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncopawa"
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE").Select
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE/ssubSUB1:SAPLCOMI:0104/ctxtEXIDV_DEASS").Text = cHUNum
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE/ssubSUB1:SAPLCOMI:0104/btn%#AUTOTEXT001").Press
            .FindById("wnd[0]/tbar[0]/btn[11]").Press
            If printYes Then printHULabel cHUNum, 2
        End If

    End With
    
    Set thisSession = Nothing
End Sub
Public Property Let batchNum(ByVal newData As String)
    cBatch = newData
End Property
Public Property Let procOrdNum(ByVal sProcOrdNum As String)
    cProcOrdNum = sProcOrdNum
End Property

Public Property Let itemNum(ByVal lItemNum As Long)
    cItemNum = lItemNum
End Property


Public Property Let batchArray(ByVal newData As Variant)
    cHUBatch = newData
End Property


