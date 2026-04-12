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
Dim cBatchNum As String
Dim cProcOrdNum As String
Dim cHUNum As String
Dim cItemNum As String

'COR6 to receive FERT(remove GI lines)

'copawa procOrdNum HUNum (to reserve HU for this po)

'cowbhuwa procOrdNum (issue material to po)

'copawa (to delete reservation)
Private Function lookupHU() As String

    Dim thisSession1 As GuiSession
    Dim result As String
    Set thisSession1 = SAPSession.CurrentSession
    With thisSession1
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nhumo"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200").Select
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/ctxtSELMATNR-LOW").Text = cItemNum
        .FindById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpTEXT-200/ssub%_SUBSCREEN_ORDER_CRITERIA:RHU_HELP:2020/txtSELCHARG-LOW").Text = cBatchNum
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
    cBatchNum = UCase(batchNum)
    lookupHU
    With thisSession
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
            If .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-CHARG[1," & t & "]").Text = cBatchNum Then
                .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-MENGE[2," & t & "]").Text = "1"
            Else
                .FindById("wnd[1]/usr/tblSAPLCOMITC_PG_ISSUE/txtVHUMI_PGI-MENGE[2," & t & "]").Text = "0"
            End If

        Next t
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        If x > 2 Then
            .FindById("wnd[0]/tbar[0]/okcd").Text = "copawa"
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE").Select
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE/ssubSUB1:SAPLCOMI:0104/ctxtEXIDV_DEASS").Text = cHUNum
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/tabsSTRIP/tabpHUDE/ssubSUB1:SAPLCOMI:0104/btn%#AUTOTEXT001").Press
            .FindById("wnd[0]/tbar[0]/btn[11]").Press
            printHULabel cHUNum, 2
        End If

    End With
    
    Set thisSession = Nothing
End Sub
