VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCOOISPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cProcOrdNum As String

Sub goodsMovements()
    Dim procOrdNum As String
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen
    cProcOrdNum = InputBox("What is the process order number?", , cProcOrdNum)
    ' --- Start interacting with the COOISPI screen in SAP ---
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "COOISPI" ' Enter transaction code COOISPI
    thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOD000"
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/jjbrown"
    ' Select checkboxes for "Order" and "Material"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = False
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = False
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = cProcOrdNum ' Clear order number field
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = "" ' Enter the provided item number
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105" ' Set plant to q105
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "" ' Set system status 1 to ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "" ' Set system status 2 to ""
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
    ' --- End of COOISPI screen interaction ---


    ' Check if there is data for the selection
    If SAPStatus.CurrentStatus = "There is no data for the selection" Then
        MsgBox "There are no process orders with this number - " & procOrdNum, vbOKOnly
        Exit Sub
    End If
    Set thisSession = Nothing
End Sub
Sub confirmations()
    Dim procOrdNum As String
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen
    cProcOrdNum = InputBox("What is the process order number?", , cProcOrdNum)
    ' --- Start interacting with the COOISPI screen in SAP ---
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "COOISPI" ' Enter transaction code COOISPI
    thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOR000"
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/jrp"
' Select checkboxes for "Order" and "Material"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = False
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = False
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = cProcOrdNum ' Clear order number field
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = "" ' Enter the provided item number
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105" ' Set plant to q105
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "" ' Set system status 1 to ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "" ' Set system status 2 to ""
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
    ' --- End of COOISPI screen interaction ---


    ' Check if there is data for the selection
    If SAPStatus.CurrentStatus = "There is no data for the selection" Then
        MsgBox "There are no process orders with this number - " & procOrdNum, vbOKOnly
        Exit Sub
    End If
    Set thisSession = Nothing
End Sub
