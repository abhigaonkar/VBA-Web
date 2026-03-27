Attribute VB_Name = "FIFO"
Option Explicit
Private msngcSandAtFurnace As Single
Function useMB53(sloc As String) As Variant
' This function finds the sand available at the specified furnace location (sloc) in SAP.
' It uses the MB52 transaction to retrieve the data.
' Returns an array containing the item number, batch number, and quantity of sand.
' If multiple batches are found, it calls the useMB51 function to get FIFO (First-In, First-Out) information.
    Dim thisGUI As GuiSession
    On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession ' Initialize SAP GUI session

    setStartScreen ' Call a function to set the initial screen

    Dim x As Integer
    Dim sandAtLoc() As Variant
    ReDim sandAtLoc(2, 3) ' Array to store sand data (item, batch, qty)
    msngcSandAtFurnace = 0
    ' --- Start interacting with the MB52 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "mb52" ' Enter transaction code MB52
    thisGUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "" ' Clear material number field
    thisGUI.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "Q105" ' Set plant to Q105
    thisGUI.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = "" ' Clear batch number field
    thisGUI.FindById("wnd[0]/usr/chkNOZERO").Selected = True
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = sloc ' Set storage location to the provided value
    thisGUI.FindById("wnd[0]/usr/chkNEGATIV").Selected = False
    thisGUI.FindById("wnd[0]/usr/chkXMCHB").Selected = True
    thisGUI.FindById("wnd[0]/usr/chkNOVALUES").Selected = False
    thisGUI.FindById("wnd[0]/usr/radPA_FLT").Select
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/PARKINSONJ" ' Set a variant (a custom view)

    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (F8)
    ' --- End of MB52 screen interaction ---

    ' Determine the number of data items returned
    Dim b As Integer
    thisGUI.FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status")
    If productionMode Then
        b = thisGUI.FindById("wnd[1]/usr/lbl[17,10]").Text ' Extract the number of entries from the status window
    Else
        b = thisGUI.FindById("wnd[1]/usr/lbl[17,8]").Text ' Extract the number of entries from the status window
    End If
    thisGUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window

    ' Loop through the results and extract data
    For x = 3 To 2 + b
        ' Extract item number, batch number, and quantity from the screen
        sandAtLoc(x - 3, 0) = thisGUI.FindById("wnd[0]/usr/lbl[1," & x & "]").Text ' Item number
        sandAtLoc(x - 3, 1) = thisGUI.FindById("wnd[0]/usr/lbl[53," & x & "]").Text ' Batch number
        sandAtLoc(x - 3, 2) = thisGUI.FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
        msngcSandAtFurnace = msngcSandAtFurnace + CSng(sandAtLoc(x - 3, 2))
    Next x

    Set thisGUI = Nothing ' Release the GUI session object

    ' If more than one batch is found, call useMB51 to get FIFO details
    If b = 1 Then
        useMB53 = sandAtLoc() ' Return the sand data
    Else
        useMB53 = useMB51(sloc, sandAtLoc(), b) ' Call useMB51 to get FIFO data
    End If

Exit Function

ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB53 of FIFO Module" ' Call a custom error handler

End Function
Public Property Get cSandAtFurnace() As Single
    cSandAtFurnace = msngcSandAtFurnace
End Property


Function useMB51(furnace As String, sandDetails() As Variant, numItems As Integer) As Variant
' This function finds the document number of the 311 transaction (goods movement)
' to enable FIFO (First-In, First-Out) processing of sand batches.
' It uses the MB51 transaction and filters by the provided sandDetails.
' Returns an array containing item number, batch number, quantity, and document number.

    Dim thisGUI As GuiSession
    On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession ' Initialize SAP GUI session

    Dim newArray() As Variant
    ReDim newArray(2, 3) ' Array to store the results (item, batch, qty, docnum)
    Dim x, y As Integer

    setStartScreen ' Call a function to set the initial screen

    ' --- Start interacting with the MB51 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "mb51" ' Enter transaction code MB51
    thisGUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "" ' Clear material number field
    thisGUI.FindById("wnd[0]/usr/ctxtLIFNR-LOW").Text = "" ' Clear vendor number field
    thisGUI.FindById("wnd[0]/usr/ctxtKUNNR-LOW").Text = "" ' Clear customer number field
    thisGUI.FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "311" ' Set movement type to 311 (goods movement)
    thisGUI.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = "" ' Clear order number field
    thisGUI.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = "" ' Clear posting date field
    thisGUI.FindById("wnd[0]/usr/txtUSNAM-LOW").Text = "" ' Clear username field
    thisGUI.FindById("wnd[0]/usr/txtMBLNR-LOW").Text = "" ' Clear material document field
    thisGUI.FindById("wnd[0]/usr/radRFLAT_L").Select ' Select a radio button
    thisGUI.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/JAVIEW" ' Set a layout variant (a custom view)
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = furnace ' Set storage location to the provided value
    thisGUI.FindById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").Press ' Open batch selection dialog

    ' Enter batch numbers from sandDetails into the selection dialog
    For x = 0 To 2
        thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & x & "]").Text = sandDetails(x, 1)
    Next x

    thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,3]").SetFocus ' Set focus to a button in the dialog
    thisGUI.FindById("wnd[1]/tbar[0]/btn[8]").Press ' Execute (F8) in the dialog
    thisGUI.FindById("wnd[0]/usr/txtMBLNR-LOW").Text = "" ' Clear material document field
    thisGUI.FindById("wnd[0]/usr/radRFLAT_L").SetFocus ' Set focus back to the radio button
    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (F8) in the main window
    ' --- End of MB51 screen interaction ---

    ' Sort the results by material document number (MBLNR)
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell -1, "MBLNR" ' Set focus to the MBLNR column
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleColumn = "BUDAT" ' Make the posting date column visible
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectColumn "MBLNR" ' Select the MBLNR column
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = ""
    thisGUI.FindById("wnd[0]/tbar[1]/btn[41]").Press
    
    'figure out the item number, batch, and qty available here
    Dim b As Integer
    Dim batchesMatch As Boolean
    batchesMatch = True
    b = 0
        newArray(0, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MATNR")
        newArray(0, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "CHARG")
        newArray(0, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MBLNR")
        b = b + 1
        Do While batchesMatch
            If thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "CHARG") = newArray(0, 1) Then
                b = b + 1
            Else
                batchesMatch = False
            End If
        Loop
        newArray(1, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "MATNR")
        newArray(1, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "CHARG")
        newArray(1, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "MBLNR")
        If numItems > 2 Then
            b = b + 1
            newArray(2, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "MATNR")
            newArray(2, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "CHARG")
            newArray(2, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(b, "MBLNR")
        End If
    For y = 0 To 2
        For x = 0 To 2
            If sandDetails(x, 0) = newArray(y, 0) Then
                If sandDetails(x, 1) = newArray(y, 1) Then
                    newArray(y, 2) = sandDetails(x, 2)
                End If
            End If
        Next x
    Next y

    thisGUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisGUI = Nothing
    useMB51 = newArray
Exit Function
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB51 of FIFO Module"

End Function
    

