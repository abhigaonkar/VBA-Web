Attribute VB_Name = "FIFO"
Option Explicit

Function findMaterial(itemNum As String, batch As String, Optional sloc As String) As Variant
    ' This function finds the sand available at the specified storage location (sloc) in SAP.
    ' It uses the MB52 transaction to retrieve the data.
    ' Returns an array containing the item number, batch number, and quantity of sand for up to 3 different batches.
    ' If multiple batches are found, it calls the useMB51 function to get FIFO (First-In, First-Out) information.

    Dim thisGUI As GuiSession ' Declare a variable to hold the SAP GUI session object
    On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession ' Initialize the SAP GUI session
'    setStartScreen ' Call a function to navigate to the SAP Easy Access screen
    Dim thisLoc As String
    If sloc = "" Then
        thisLoc = ""
    Else
        thisLoc = sloc
    End If
    Dim x As Integer ' Loop counter
    Dim sandAtLoc() As Variant ' Array to store sand data (item number, batch number, quantity)
    ReDim sandAtLoc(2, 3) ' Initialize the array with a size of 3 rows and 4 columns (0 to 3)

    ' --- Start interacting with the MB52 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52" ' Enter transaction code MB52
    thisGUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = itemNum ' material number field
    thisGUI.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "Q105" ' Set plant to Q105
    thisGUI.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = batch '  batch number field
    thisGUI.FindById("wnd[0]/usr/chkNOZERO").Selected = True
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = thisLoc ' Set storage location to the provided value
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/PARKINSONJ" ' Set a variant (likely a custom view)
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").SetFocus ' Set focus to the variant field
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").CaretPosition = 11 ' Set caret position (not clear why)
    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
    ' --- End of MB52 screen interaction ---
    
    If SAPStatus.CurrentStatus <> "No stock exists for specified data" Then
        ' --- Determine the number of data items returned ---
        Dim b As Integer ' Variable to store the number of data items
        thisGUI.FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status") to get the number of entries
'        b = thisGUI.FindById("wnd[1]/usr/lbl[17,10]").Text ' Extract the number of entries from the status window
        b = thisGUI.FindById("wnd[1]/usr/lbl[17,8]").Text ' Extract the number of entries from the status window
        thisGUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window
    
        ' --- Loop through the results and extract data ---
        For x = 3 To 2 + b ' Loop through the rows in the MB52 result list
            sandAtLoc(x - 3, 0) = thisGUI.FindById("wnd[0]/usr/lbl[1," & x & "]").Text ' Extract item number
            sandAtLoc(x - 3, 3) = thisGUI.FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Extract location
            sandAtLoc(x - 3, 1) = thisGUI.FindById("wnd[0]/usr/lbl[53," & x & "]").Text ' Extract batch number
            sandAtLoc(x - 3, 2) = thisGUI.FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Extract quantity
        Next x
    Else
        
    End If
    Set thisGUI = Nothing ' Release the SAP GUI session object

    ' --- Call useMB51 to get FIFO details if multiple batches are found ---
    If b = 1 Then
       ' findMaterial = sandAtLoc() ' Return the sand data array if only one batch is found
    Else
        MsgBox "This item/batch combination has nothing in inventory.", vbOKOnly
        sandAtLoc(0, 2) = 0
        sandAtLoc(0, 3) = ""
    End If
    findMaterial = sandAtLoc()
Exit Function ' Exit the function

ErrHandler:
    ' --- Error handling ---
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB52 of FIFO Module" ' Call the standard error handler
End Function

Function useMB51(furnace As String, sandDetails() As Variant, numItems As Integer) As Variant
    ' This function finds the document number of the 311 transaction (goods movement) for the sand batches
    ' to enable FIFO (First-In, First-Out) processing.
    ' It uses the MB51 transaction and filters by the provided sandDetails.
    ' Returns an array containing item number, batch number, quantity, and document number for up to 3 batches.

    Dim thisGUI As GuiSession ' Declare a variable to hold the SAP GUI session object
    On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession ' Initialize the SAP GUI session

    Dim newArray() As Variant ' Array to store the results (item number, batch number, quantity, document number)
    ReDim newArray(2, 3) ' Initialize the array with a size of 3 rows and 4 columns (0 to 3)
    Dim x As Integer, y As Integer ' Loop counters

    setStartScreen ' Call a function to navigate to the SAP Easy Access screen

    ' --- Start interacting with the MB51 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb51" ' Enter transaction code MB51
    thisGUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "" ' Clear material number field
    thisGUI.FindById("wnd[0]/usr/ctxtLIFNR-LOW").Text = "" ' Clear vendor number field
    thisGUI.FindById("wnd[0]/usr/ctxtKUNNR-LOW").Text = "" ' Clear customer number field
    thisGUI.FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "311" ' Set movement type to 311 (goods receipt)
    thisGUI.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = "" ' Clear order number field
    thisGUI.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = "" ' Clear posting date field
    thisGUI.FindById("wnd[0]/usr/txtUSNAM-LOW").Text = "" ' Clear username field
    thisGUI.FindById("wnd[0]/usr/txtMBLNR-LOW").Text = "" ' Clear material document field
    thisGUI.FindById("wnd[0]/usr/radRFLAT_L").Select ' Select a radio button (not clear which one)
    thisGUI.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/JAVIEW" ' Set a layout variant (likely a custom view)
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = furnace ' Set storage location to the provided value
    thisGUI.FindById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").Press ' Open batch selection dialog

    ' --- Enter batch numbers from sandDetails into the selection dialog ---
    For x = 0 To 2 ' Loop through the first 3 batches in the sandDetails array
        thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & x & "]").Text = sandDetails(x, 1) ' Enter the batch number
    Next x

    thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,3]").SetFocus ' Set focus to a button in the dialog
    thisGUI.FindById("wnd[1]/tbar[0]/btn[8]").Press ' Execute (likely F8) in the dialog
    thisGUI.FindById("wnd[0]/usr/txtMBLNR-LOW").Text = "" ' Clear material document field
    thisGUI.FindById("wnd[0]/usr/radRFLAT_L").SetFocus ' Set focus back to the radio button
    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8) in the main window

    ' --- Sort the results by material document number (MBLNR) ---
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell -1, "MBLNR" ' Set focus to the MBLNR column
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleColumn = "BUDAT" ' Make the posting date column visible (likely for sorting purposes)
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectColumn "MBLNR" ' Select the MBLNR column
    thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = "" ' Clear any selected rows
    thisGUI.FindById("wnd[0]/tbar[1]/btn[41]").Press ' Sort ascending by MBLNR

    ' --- Extract data from the sorted results ---
    ' Extract item number, batch number, and document number from the first 3 rows of the sorted results
    newArray(0, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MATNR") ' Item number
    newArray(0, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "CHARG") ' Batch number
    newArray(0, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MBLNR") ' Document number
    newArray(1, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(1, "MATNR") ' Item number
    newArray(1, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(1, "CHARG") ' Batch number
    newArray(1, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(1, "MBLNR") ' Document number
    If numItems > 2 Then ' Check if there are more than 2 items
        newArray(2, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(2, "MATNR") ' Item number
        newArray(2, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(2, "CHARG") ' Batch number
        newArray(2, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(2, "MBLNR") ' Document number
    End If

    ' --- Match the extracted data with the input sandDetails and update the quantity ---
    For y = 0 To 2 ' Loop through the first 3 rows of the newArray
        For x = 0 To 2 ' Loop through the first 3 rows of the sandDetails array
            If sandDetails(x, 0) = newArray(y, 0) Then ' Check if the item numbers match
                If sandDetails(x, 1) = newArray(y, 1) Then ' Check if the batch numbers match
                    newArray(y, 2) = sandDetails(x, 2) ' If both match, update the quantity in the newArray with the quantity from sandDetails
                End If
            End If
        Next x
    Next y

    thisGUI.FindById("wnd[0]/tbar[0]/btn[3]").Press ' Press the "Back" button to exit MB51
    Set thisGUI = Nothing ' Release the SAP GUI session object
    useMB51 = newArray ' Return the updated newArray

Exit Function ' Exit the function

ErrHandler:
    ' --- Error handling ---
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB51 of FIFO Module" ' Call the standard error handler
End Function
