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
Dim cItemNum As String
Dim cProcOrdNum As String
Dim cNumOrders As Integer

Function getProcOrdNum(ByVal itemNum As String, ByVal unused As Boolean, ByVal batchNum As String) As String
    ' This function retrieves the process order number for the given item number from SAP.

    Dim thisSession As GuiSession
    Dim x  As Integer, p As Integer
    Dim result As String
On Error GoTo ErrHandler
    cItemNum = itemNum
        Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
        With thisSession
            ' --- Start interacting with the COOISPI screen in SAP ---
            .FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOOISPI" ' Enter transaction code COOISPI
            .FindById("wnd[0]").SendVKey 0 ' Press Enter
            ' Select checkboxes for "Order" and "Material"
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = "" ' Clear order number field
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = cItemNum ' Enter the provided item number
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105" ' Set plant to q105
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").Text = "SAPCEM1"
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco" ' Set system status 1 to "teco"
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "cnf"
            If unused Then
                .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "gmps" ' Set system status 2 to "gmps"
            
                .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYOP1").SetFocus  ' makes sure we're only seeing released orders
                .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYOP1").CaretPosition = 3
                .FindById("wnd[0]").SendVKey 4
                .FindById("wnd[1]/usr/lbl[1,4]").SetFocus
                .FindById("wnd[1]/usr/lbl[1,4]").CaretPosition = 0
                .FindById("wnd[1]").SendVKey 2
            End If
            .FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
            ' --- End of COOISPI screen interaction ---
            cNumOrders = .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
            
   '         create process orders
            
            x = 0 ' Initialize row counter
            ' Check if there is data for the selection
            If SAPStatus.CurrentStatus = "There is no data for the selection" Then
                MsgBox "There are no open process orders for " & itemNum, vbOKOnly
                Exit Function
            End If
            For p = 0 To cNumOrders - 1
                If batchNum = .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(p, "CHARG") Then
                    cProcOrdNum = .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(p, "AUFNR")
                    getProcOrdNum = cProcOrdNum
                    Exit Function
                End If
            Next p
            ' Navigate to the first cell in the results and double-click to open the order
            .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellRow = x
            .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellColumn = "AUFNR"
            .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
            result = .FindById("wnd[0]/usr/txtCAUFVD-AUFNR").Text
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
            
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
        End With
        Set thisSession = Nothing
    
    cProcOrdNum = result
    getProcOrdNum = result
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getProcOrdNum function of UserForm1"
    
End Function
Sub goodsMovements()
    ' Declare variables
    Dim procOrdNum As String
    Dim thisSession As Object ' Changed to Object for late binding, more flexible with error handling
    Dim SAPGuarding As Boolean ' To track if SAP GUI is actively being guarded
    Dim obj As Object ' Generic object for checking SAP GUI elements

    ' Initialize SAP GUI session
    On Error GoTo ErrorHandler
    Set thisSession = SAPSession.CurrentSession

    ' Check if session was successfully obtained
    If thisSession Is Nothing Then
        MsgBox "Could not establish SAP GUI session. Please ensure SAP GUI is open and a session is active.", vbCritical
        Exit Sub
    End If

    ' Turn on SAP GUI scripting error guarding
    ' This helps catch errors when SAP GUI objects aren't found
    On Error GoTo SAPGuiErrorHandler
    SAPGuarding = True
    thisSession.FindById("wnd[0]").LockControls = True ' Optional: Lock controls for faster execution, but can hide issues

    ' Call a function to set the initial screen (ensure this function also has error handling if it interacts with SAP)
'    setStartScreen

    ' Get process order number from user
    procOrdNum = InputBox("What is the process order number?", , procOrdNum) ' Using procOrdNum directly for input

    ' --- Start interacting with the COOISPI screen in SAP ---
    ' Navigate to COOISPI transaction
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOOISPI"
    thisSession.FindById("wnd[0]").SendVKey 0

    ' Set list type and ALV variant
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOD000"
    thisSession.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/jjbrown"

    ' Select/Deselect checkboxes for "Order" and "Material"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = False
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = False

    ' Enter selection criteria
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = procOrdNum ' Use procOrdNum
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = ""

    ' Execute report (likely F8)
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press

    ' Turn off SAP GUI scripting error guarding and unlock controls
    SAPGuarding = False
    thisSession.FindById("wnd[0]").LockControls = False

    ' --- End of COOISPI screen interaction ---

    ' Check if there is data for the selection
    ' This specific check assumes SAPStatus.CurrentStatus is reliable.
    ' A more robust check might involve looking for specific text on the SAP GUI screen itself
    ' or checking for the presence of a table control.
    If SAPStatus.CurrentStatus = "There is no data for the selection" Then
        MsgBox "There are no process orders with this number - " & procOrdNum, vbExclamation, "No Data Found"
        GoTo CleanUp
    End If

CleanUp:
    ' Release object references
    Set thisSession = Nothing
    Exit Sub

SAPGuiErrorHandler:
    ' Handles errors specifically related to SAP GUI object access
    If Err.Number = -2147352567 Or Err.Number = 424 Then ' Specific error numbers for "Object not found" or "Object required"
        MsgBox "SAP GUI element not found or accessible. This usually means the screen layout changed, or the script is out of sync with SAP GUI. Error: " & Err.Description, vbCritical
    Else
        ' For any other SAP GUI related errors
        MsgBox "An unexpected SAP GUI scripting error occurred: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    End If
    Resume CleanUp ' Go to cleanup before exiting

ErrorHandler:
    ' General error handler for other VBA errors
    If SAPGuarding Then
        ' If we were in the middle of SAP GUI interaction, try to turn off guarding and unlock controls
        On Error Resume Next ' Temporarily disable error handling for this line
        thisSession.FindById("wnd[0]").LockControls = False
        On Error GoTo ErrorHandler ' Re-enable error handling
    End If

    MsgBox "An unexpected error occurred: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Resume CleanUp ' Go to cleanup before exiting
End Sub
Sub confirmations()
    Dim procOrdNum As String
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
'    setStartScreen ' Call a function to set the initial screen
    cProcOrdNum = InputBox("What is the process order number?", , cProcOrdNum)
    ' --- Start interacting with the COOISPI screen in SAP ---
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOOISPI" ' Enter transaction code COOISPI
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
Property Let itemNum(ByVal vNewValue As String)
On Error GoTo ErrHandler

    cItemNum = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let itemNum of clsCOOISPI."
    
End Property
Property Let processOrderNum(ByVal vNewValue As String)
On Error GoTo ErrHandler

    cProcOrdNum = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let processOrderNum of clsCOOISPI."
    
End Property
Property Get numOrders(ByVal itemNum As String) As Integer
    Dim thisSession As GuiSession
    cItemNum = itemNum
    'If onlineMode Then
        Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
 '       setStartScreen ' Call a function to set the initial screen
        With thisSession
            ' --- Start interacting with the COOISPI screen in SAP ---
            .FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOOISPI" ' Enter transaction code COOISPI
            .FindById("wnd[0]").SendVKey 0 ' Press Enter
            ' Select checkboxes for "Order" and "Material"
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = "" ' Clear order number field
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = cItemNum ' Enter the provided item number
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105" ' Set plant to q105
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").Text = "SAPCEM1"
            .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco" ' Set system status 1 to "teco"
 '           .FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "gmps" ' Set system status 2 to "gmps"
            .FindById("wnd[0]").SendVKey 4
            .FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
            ' --- End of COOISPI screen interaction ---
            cNumOrders = .FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
        End With
        Set thisSession = Nothing
    numOrders = cNumOrders
End Property

