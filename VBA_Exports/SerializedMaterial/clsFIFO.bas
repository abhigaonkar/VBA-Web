VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFIFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private msngcSandAtFurnace As Single
Function useMB52(sLoc As String) As Variant
' This function finds the sand available at the specified furnace location (sloc) in SAP.
' It uses the MB52 transaction to retrieve the data.
' Returns an array containing the item number, batch number, and quantity of sand.
' If multiple batches are found, it calls the useMB51 function to get FIFO (First-In, First-Out) information.
    Dim thisGUI As GuiSession
    On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession ' Initialize SAP GUI session

'    setStartScreen ' Call a function to set the initial screen

    Dim x As Integer
    Dim sandAtLoc() As Variant
    msngcSandAtFurnace = 0
    ' --- Start interacting with the MB52 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52" ' Enter transaction code MB52
    thisGUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "" ' Clear material number field
    thisGUI.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "Q105" ' Set plant to Q105
    thisGUI.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = "" ' Clear batch number field
    thisGUI.FindById("wnd[0]/usr/radPA_FLT").Select
    thisGUI.FindById("wnd[0]/usr/chkNOZERO").Selected = True
    
    thisGUI.FindById("wnd[0]/usr/chkNEGATIV").Selected = False
    thisGUI.FindById("wnd[0]/usr/chkXMCHB").Selected = True
    thisGUI.FindById("wnd[0]/usr/chkNOVALUES").Selected = False
    thisGUI.FindById("wnd[0]/usr/radPA_FLT").Select
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = sLoc ' Set storage location to the provided value
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/PARKINSONJ" ' Set a variant (a custom view)

    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (F8)
    ' --- End of MB52 screen interaction ---
    If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
        MsgBox "Furnace " & sLoc & " has no sand on it in SAP.", vbCritical
        Exit Function
        
    End If
    ' Determine the number of data items returned
    Dim b As Integer
    b = thisGUI.FindById("wnd[0]/usr").Children.Count / 8 - 1
    

    ' Loop through the results and extract data
    ReDim sandAtLoc(b - 1, 3) ' Array to store sand data (item, batch, qty)

    For x = 3 To 2 + b
        ' Extract item number, batch number, and quantity from the screen
        sandAtLoc(x - 3, 0) = thisGUI.FindById("wnd[0]/usr/lbl[1," & x & "]").Text ' Item number
        sandAtLoc(x - 3, 1) = thisGUI.FindById("wnd[0]/usr/lbl[53," & x & "]").Text ' Batch number
        sandAtLoc(x - 3, 2) = thisGUI.FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
        msngcSandAtFurnace = msngcSandAtFurnace + CSng(sandAtLoc(x - 3, 2))
    Next x
    
'   Save Sand Level
    Dim thisWB As Workbook
       
    Dim thisSheet As Worksheet
    Set thisWB = ThisWorkbook
    Set thisSheet = thisWB.Sheets("Sheet2")
    thisSheet.Cells(2, 27).Value = msngcSandAtFurnace
    Set thisSheet = Nothing
    Set thisWB = Nothing
    
    Set thisGUI = Nothing ' Release the GUI session object

    ' If more than one batch is found, call useMB51 to get FIFO details
    If b = 1 Then
        useMB52 = sandAtLoc() ' Return the sand data
    Else
        useMB52 = useMB51(sLoc, sandAtLoc(), b) ' Call useMB51 to get FIFO data
    End If

Exit Function

ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB52 of FIFO Module" ' Call a custom error handler

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
    Dim x, y As Integer

 '   setStartScreen ' Call a function to set the initial screen

    ' --- Start interacting with the MB51 screen ---
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb51" ' Enter transaction code MB51
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
    x = 0
    Dim p As Integer
    ' Enter batch numbers from sandDetails into the selection dialog
    For p = 0 To numItems - 1
        thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & x & "]").Text = sandDetails(p, 1)
        If x = 7 Then
            thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").VerticalScrollbar.Position = 8
            x = 0
        End If
        x = x + 1
    Next p

'    thisGUI.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,3]").SetFocus ' Set focus to a button in the dialog
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
    Dim uniqueItems As Integer
    uniqueItems = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount - 1
    ReDim newArray(uniqueItems, 3) ' Array to store the results (item, batch, qty, docnum)
    x = 0
    For x = 0 To uniqueItems
        newArray(x, 0) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "MATNR")
        newArray(x, 1) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "CHARG")
        newArray(x, 3) = thisGUI.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "MBLNR")
    Next x
    Dim keepGoing As Boolean
    Dim t As Integer
    y = 0
    For x = 0 To UBound(sandDetails)
        keepGoing = True
        Do While keepGoing
'            Debug.Print "Sand: " & sandDetails(x, 0) & " = " & newArray(y, 0) & vbCrLf & "Batch: " & sandDetails(x, 1) & " = " & newArray(y, 1)
            If sandDetails(x, 0) = newArray(y, 0) And sandDetails(x, 1) = newArray(y, 1) Then
                sandDetails(x, 3) = newArray(y, 3)
                y = 0
                keepGoing = False
            Else
                y = y + 1
            End If
        Loop
    Next x
    
    thisGUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisGUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisGUI = Nothing
    useMB51 = sort_2D(sandDetails, 0, UBound(sandDetails), 3)
Exit Function
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB51 of FIFO Module"

End Function
    



