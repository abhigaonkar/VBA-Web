Attribute VB_Name = "Module2"
Option Explicit

Function saveToSAP(ByVal lineNum As Long) As Boolean
    On Error GoTo ErrHandler   ' Global error handler for this function

    '=================================================================
    ' Purpose:
    '   - For a given row (lineNum) on Sheet1:
    '       1) Determine process order based on item & batch
    '       2) Perform GI of baked item to process order
    '       3) Perform GR of finished item
    '       4) Enter quality dimensions (bow, length)
    '       5) Issue packaging materials & confirm order (COR6)
    '       6) Optionally do additional confirmations based on recipe
    '
    ' Requirements:
    '   - Helper classes:
    '       clsCOOISPI, clsHU, clsQualityEntry, clsNewCOR6, clsCOR3
    '   - Functions:
    '       findItemDetailsFromBatch, updateStatus, inspectPackOperation
    '   - Worksheet:
    '       ThisWorkbook.Sheets("Sheet1") with specific layout
    '=================================================================

    Dim goodsIssued() As Variant
    Dim numGoodsToIssue As Integer
    Dim processOrdNum As String
    Dim orderCheck As clsCOOISPI
    Dim itemNum As String
    Dim bakedItemNum As String
    Dim batchNum As String
    Dim thisGI_GR As clsHU
    Dim thisQM As clsQualityEntry
    Dim thisCOR6 As clsNewCOR6
    Dim thisOrder As clsCOR3
    Dim masterRecipe As String
    Dim localSheet As Worksheet
    Dim x As Integer, j As Integer
    Dim doOtherConfirmations As Boolean
    Dim numWaffle As Integer

    '-----------------------------------------------------------------
    ' Initialize objects & variables
    '-----------------------------------------------------------------
    Set localSheet = ThisWorkbook.Sheets("Sheet1")
    Set orderCheck = New clsCOOISPI
    Set thisGI_GR = New clsHU

    ' Basic flags
    doOtherConfirmations = False
    saveToSAP = False    ' default if we exit early

    '-----------------------------------------------------------------
    ' Read basic data from the worksheet
    '   Col 4: Item number
    '   Col 8: Batch number
    '-----------------------------------------------------------------
    itemNum = CStr(localSheet.Cells(lineNum, 4).Value)
    batchNum = CStr(localSheet.Cells(lineNum, 8).Value)

    ' Get baked item from batch (function expected to return array)
    bakedItemNum = CStr(findItemDetailsFromBatch(batchNum)(0))

    ' If no baked item was found, exit function
    If Len(bakedItemNum) = 0 Or bakedItemNum = "0" Then
        updateStatus "No baked item found for batch " & batchNum
        GoTo CleanExit
    End If

    ' Write baked item to sheet (col 7)
    localSheet.Cells(lineNum, 7).Value = bakedItemNum

    '-----------------------------------------------------------------
    ' Get process order for this item and baked item
    '-----------------------------------------------------------------
    processOrdNum = orderCheck.getProcOrdNum(itemNum, bakedItemNum)

    ' Show process order on the user form (if open)
    PackForm.TextBox7.Value = processOrdNum

    ' If no open orders found, alert user and exit
    If orderCheck.numOrders = 0 Or Len(processOrdNum) = 0 Then
        MsgBox "There aren't any open process orders for " & itemNum & _
               " using " & bakedItemNum & " for source stock.", vbCritical
        GoTo CleanExit
    End If

    ' Store process order back into sheet (col 3)
    localSheet.Cells(lineNum, 3).Value = processOrdNum

    ' No longer need orderCheck
    Set orderCheck = Nothing

    '-----------------------------------------------------------------
    ' 1) GOODS ISSUE (GI) of baked item to process order
    '-----------------------------------------------------------------
    updateStatus "Issuing " & bakedItemNum & " to process order " & processOrdNum
    thisGI_GR.huGI processOrdNum, bakedItemNum, batchNum

    '-----------------------------------------------------------------
    ' 2) GOODS RECEIPT (GR) of finished item
    '   - Typically COP1 (order confirmation/posting), COWBHUWE (HU)
    '-----------------------------------------------------------------
    updateStatus "Receiving " & itemNum
    thisGI_GR.cop1
    thisGI_GR.cowbhuwe

    Set thisGI_GR = Nothing

    '-----------------------------------------------------------------
    ' 3) QUALITY MEASUREMENT ENTRY (bow, length)
    '   Col 23 & 24: Bow measurements – take the greater of the two
    '   Col 22: Length
    '   Col 25: Calculated bow (returned from class)
    '-----------------------------------------------------------------
    updateStatus "Entering dimensions in SAP."
    Set thisQM = New clsQualityEntry

    thisQM.batchNum = batchNum

    ' Use the larger of the two bow measurements
    If localSheet.Cells(lineNum, 23).Value >= localSheet.Cells(lineNum, 24).Value Then
        thisQM.bowValue = localSheet.Cells(lineNum, 23).Value
    Else
        thisQM.bowValue = localSheet.Cells(lineNum, 24).Value
    End If

    thisQM.itemNum = CStr(localSheet.Cells(lineNum, 4).Value)
    thisQM.length = localSheet.Cells(lineNum, 22).Value
    thisQM.acceptable = True

    ' Perform the QM characteristic entry in SAP
    thisQM.characteristicEntry

    ' Write calculated bow back to sheet (col 25)
    localSheet.Cells(lineNum, 25).Value = thisQM.calculatedBow

    Set thisQM = Nothing

    '-----------------------------------------------------------------
    ' 4) PACKAGING ISSUE + ORDER CONFIRMATION (COR6)
    '   - Issue sleeving, hamper, strapping, waffle boards, return items
    '   - Close order and record labor
    '-----------------------------------------------------------------
    updateStatus "Issuing packaging and doing confirmation of order."

    Set thisCOR6 = New clsNewCOR6

    With thisCOR6
        .procOrdNum = processOrdNum
        .closeOrder = True
        .applyMatl = True
    End With

    ' Determine how many materials to issue:
    '   Start with 1 for sleeving
    '   +2 if strapping completed
    '   +1 if return item completed
    numGoodsToIssue = 1   ' sleeving

    If CStr(localSheet.Cells(lineNum, 17).Value) <> "" Then
        numGoodsToIssue = numGoodsToIssue + 2   ' hamper + strapping
    End If

    If CStr(localSheet.Cells(lineNum, 26).Value) <> "" Then
        numGoodsToIssue = numGoodsToIssue + 1   ' return item
    End If

    ' Dimension array: (0 To numGoodsToIssue, 0 To 4)
    '   Col 0: Item number
    '   Col 1: Quantity
    '   Col 2: Location
    '   Col 3: Batch (if any)
    '   Col 4: Movement type
    ReDim goodsIssued(numGoodsToIssue, 4)
    thisCOR6.applyNumMatl = numGoodsToIssue

    '-------------------------------------------------------------
    ' 4a) Sleeving
    '   - Hardcoded item: 137235
    '   - Quantity: tube length + 762 mm, converted to feet (304.8)
    '   - Location: 0100
    '   - Movement type: 261
    '-------------------------------------------------------------
    goodsIssued(0, 0) = "137235"  ' Sleeving item number
    goodsIssued(0, 1) = Round((localSheet.Cells(lineNum, 22).Value + 762) / 304.8, 0)
    goodsIssued(0, 2) = "0100"    ' Sleeving location
    goodsIssued(0, 3) = ""        ' Sleeving batch (none)
    goodsIssued(0, 4) = "261"     ' Sleeving movement type

    '-------------------------------------------------------------
    ' 4b) Hamper, Strapping, Return Hamper, Waffle Board (optional)
    '-------------------------------------------------------------
    If numGoodsToIssue > 1 Then
        x = 1

        ' Hamper + Strapping (if hamper item is present)
        If CStr(localSheet.Cells(lineNum, 13).Value) <> "" Then
            ' Hamper
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 13).Value   ' Hamper item
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 14).Value   ' Hamper qty
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 15).Value   ' Hamper location
            goodsIssued(x, 3) = ""                                    ' Hamper batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 16).Value   ' Hamper movement type
            x = x + 1

            ' Strapping
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 17).Value   ' Strapping item
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 18).Value   ' Strapping qty
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 19).Value   ' Strapping location
            goodsIssued(x, 3) = ""                                    ' Strapping batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 20).Value   ' Strapping movement type
            x = x + 1
        End If

        ' Return Hamper (if present)
        If CStr(localSheet.Cells(lineNum, 26).Value) <> "" Then
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 26).Value   ' Return hamper item
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 27).Value   ' Return hamper qty
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 28).Value   ' Return hamper location
            goodsIssued(x, 3) = ""                                    ' Return hamper batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 29).Value   ' Return hamper movement type
            x = x + 1
        End If

        ' Optional waffle boards
        If MsgBox("Did you add any waffle boards?", vbYesNo) = vbYes Then
            numWaffle = CInt(InputBox("How many waffle boards did you put in the hamper?", _
                                      "Waffle Question"))
            goodsIssued(x, 0) = "124451"  ' Waffle item
            goodsIssued(x, 1) = numWaffle ' Waffle qty
            goodsIssued(x, 2) = "0100"    ' Waffle location
            goodsIssued(x, 3) = ""        ' Waffle batch
            goodsIssued(x, 4) = "261"     ' Waffle movement type
        End If
    End If

    '-----------------------------------------------------------------
    ' 5) Record labor & determine operation based on recipe
    '-----------------------------------------------------------------
    ' Default operation from helper function (inspect/pack operation)
'    thisCOR6.oper = inspectPackOperation(localSheet.Cells(lineNum, 4).Value)
    thisCOR6.matlArray = goodsIssued

    ' Get recipe via COR3 helper
    Set thisOrder = New clsCOR3
    With thisOrder
        .getDetails processOrdNum
        masterRecipe = .recipe
    End With
    Set thisOrder = Nothing

    ' Adjust operation & additional confirmations based on recipe
    Select Case masterRecipe
        Case "Q5HVL000"
            thisCOR6.oper = 20
            doOtherConfirmations = False
        Case "Q5HVC000"
            thisCOR6.oper = 60
            doOtherConfirmations = True
        ' Add more recipes here if needed
    End Select

    ' Yield, comments, and labor time
    thisCOR6.yield = 1
    thisCOR6.comment = Left(CStr(localSheet.Cells(lineNum, 12).Value), 40)
    thisCOR6.numHrs = laborTime * localSheet.Cells(lineNum, 11).Value

    ' Save COR6 and write results back to sheet
    localSheet.Cells(lineNum, 6).Value = thisCOR6.saveIt               ' Success flag / message?
    localSheet.Cells(lineNum, 5).Value = thisCOR6.FinishedItemDesc
    localSheet.Cells(lineNum, 4).Value = thisCOR6.FinishedItem

    ' Set function return based on cell value
    saveToSAP = CBool(localSheet.Cells(lineNum, 6).Value)

    Set thisCOR6 = Nothing

    '-----------------------------------------------------------------
    ' 6) Additional confirmations (cut/wash) if required
    '   - For recipe Q5HVC000, do two more operations (20 and 40)
    '-----------------------------------------------------------------
    If doOtherConfirmations Then
        updateStatus "Entering cut/wash details"

        For j = 1 To 2
            Set thisCOR6 = New clsNewCOR6
            With thisCOR6
                .applyMatl = False
                .oper = 20 * j       ' 20, 40, ...
                .numHrs = 0.25       ' Fixed 0.25 hrs per operation
                .procOrdNum = processOrdNum
                .saveIt
            End With
            Set thisCOR6 = Nothing
        Next j
    End If

    updateStatus ""

CleanExit:
    ' Clean up objects if still set
    On Error Resume Next
    Set orderCheck = Nothing
    Set thisGI_GR = Nothing
    Set thisQM = Nothing
    Set thisCOR6 = Nothing
    Set thisOrder = Nothing
    On Error GoTo 0
    Exit Function

ErrHandler:
    '-----------------------------------------------------------------
    ' Centralized error handling
    '-----------------------------------------------------------------
    MsgBox "Error in saveToSAP at line " & lineNum & vbCrLf & _
           "Description: " & Err.description & vbCrLf & _
           "Number: " & Err.Number, vbCritical, "saveToSAP Error"

    updateStatus "Error during SAP processing. Please check the log/worksheet."

    ' Return False on error
    saveToSAP = False
    Resume CleanExit
End Function
Function OLDsaveToSAP(ByVal lineNum As Long) As Boolean
' this can be deleted after the new version has proven to work
    Dim goodsIssued() As Variant
    Dim numGoodsToIssue As Integer
    Dim processOrdNum As String
    Dim orderCheck As New clsCOOISPI
    Dim itemNum As String
    Dim bakedItemNum As String
    Dim batchNum As String
    Set localSheet = ThisWorkbook.Sheets("Sheet1")
    Dim thisGI_GR As New clsHU
    Dim x As Integer, j As Integer
    Dim doOtherConfirmations As Boolean
    
    
    itemNum = localSheet.Cells(lineNum, 4).Value
    batchNum = localSheet.Cells(lineNum, 8)
    localSheet.Cells(lineNum, 3).Value = processOrdNum
    bakedItemNum = findItemDetailsFromBatch(batchNum)(0)
    If bakedItemNum = 0 Then Exit Function
    localSheet.Cells(lineNum, 7).Value = bakedItemNum
    processOrdNum = orderCheck.getProcOrdNum(itemNum, bakedItemNum)
    PackForm.TextBox7.Value = processOrdNum
    If orderCheck.numOrders = 0 Then
        MsgBox "There aren't any open process orders for " & itemNum & " using " & bakedItemNum & " for source stock.", vbCritical
        Exit Function
    End If
    Set orderCheck = Nothing
    ' do GI
    updateStatus "Issuing " & bakedItemNum & " to process order " & processOrdNum
    thisGI_GR.huGI processOrdNum, bakedItemNum, batchNum
'    thisGI_GR.procOrdNum = processOrdNum
'    thisGI_GR.batchNum = batchnum.Value
    ' now do GR
    updateStatus "Receiving " & itemNum
    thisGI_GR.cop1
    thisGI_GR.cowbhuwe
    Set thisGI_GR = Nothing
    
    updateStatus "Entering dimensions in SAP."
    Dim thisQM As New clsQualityEntry
    thisQM.batchNum = batchNum
    If localSheet.Cells(lineNum, 23).Value >= localSheet.Cells(lineNum, 24).Value Then
        thisQM.bowValue = localSheet.Cells(lineNum, 23).Value
    Else
        thisQM.bowValue = localSheet.Cells(lineNum, 24).Value
    End If
    thisQM.itemNum = localSheet.Cells(lineNum, 4).Value
    thisQM.length = localSheet.Cells(lineNum, 22).Value
    thisQM.acceptable = True
    thisQM.characteristicEntry
    localSheet.Cells(lineNum, 25).Value = thisQM.calculatedBow
    Set thisQM = Nothing

    updateStatus "Issuing packaging and doing confirmation of order."
    Dim thisCOR6 As New clsNewCOR6
    thisCOR6.procOrdNum = processOrdNum
    thisCOR6.closeOrder = True
    numGoodsToIssue = 1
    'if strapping completed, there's at least two items to issue
    If localSheet.Cells(lineNum, 17).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 2
    'if return item completed there's one more to issue
    If localSheet.Cells(lineNum, 26).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 1
    thisCOR6.applyMatl = True
    ReDim goodsIssued(numGoodsToIssue, 4)
    thisCOR6.applyNumMatl = numGoodsToIssue
    goodsIssued(0, 0) = "137235" 'sleeving item num
    goodsIssued(0, 1) = Round((localSheet.Cells(lineNum, 22).Value + 762) / 304.8, 0) 'sleeving qty is 3' longer than tube
    goodsIssued(0, 2) = "0100" 'sleeving location
    goodsIssued(0, 3) = "" 'sleeving batch
    goodsIssued(0, 4) = "261" 'sleeving movement type

    If numGoodsToIssue > 1 Then
        x = 1
        If localSheet.Cells(lineNum, 13).Value <> "" Then
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 13).Value
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 14).Value
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 15).Value
            goodsIssued(x, 3) = "" 'hamper batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 16).Value 'hamper movement type
            x = x + 1
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 17).Value 'strapping item num
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 18).Value 'strapping qty
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 19).Value 'strapping location
            goodsIssued(x, 3) = "" 'strapping batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 20).Value 'strapping movement type
            x = x + 1
        End If
        If localSheet.Cells(lineNum, 26).Value <> "" Then
            goodsIssued(x, 0) = localSheet.Cells(lineNum, 26).Value 'return hamper item num
            goodsIssued(x, 1) = localSheet.Cells(lineNum, 27).Value 'return hamper qty
            goodsIssued(x, 2) = localSheet.Cells(lineNum, 28).Value 'return hamper location
            goodsIssued(x, 3) = "" 'strapping batch
            goodsIssued(x, 4) = localSheet.Cells(lineNum, 29).Value 'return hamper movement type
            x = x + 1
        End If
        If MsgBox("Did you add any waffle boards?", vbYesNo) = vbYes Then
            Dim numWaffle As Integer
            numWaffle = InputBox("How many waffle board did you put in the hamper?", "Waffle Question")
            goodsIssued(x, 0) = "124451" 'waffle item num
            goodsIssued(x, 1) = numWaffle 'waffle qty
            goodsIssued(x, 2) = "0100" 'waffle location
            goodsIssued(x, 3) = "" 'waffle batch
            goodsIssued(x, 4) = "261" 'waffle movement type
        End If
    End If
    ' record cut/wash labor for LD. Needed here because we wouldn't know how to share the process order with each role.
    thisCOR6.oper = inspectPackOperation(localSheet.Cells(lineNum, 4).Value)
    thisCOR6.matlArray = goodsIssued
    ' determine what transactions should be ran from recipe
    Dim thisOrder As New clsCOR3
    Dim masterRecipe As String
    With thisOrder
        .getDetails processOrdNum
        masterRecipe = .recipe
    End With
    Set thisOrder = Nothing
    If masterRecipe = "Q5HVL000" Then
        thisCOR6.oper = 20
        doOtherConfirmations = False
    ElseIf masterRecipe = "Q5HVC000" Then
        thisCOR6.oper = 60
        doOtherConfirmations = True
    End If
    
    thisCOR6.yield = 1
    thisCOR6.comment = Left(localSheet.Cells(lineNum, 12).Value, 40)
    thisCOR6.numHrs = laborTime * localSheet.Cells(lineNum, 11).Value
    localSheet.Cells(lineNum, 6).Value = thisCOR6.saveIt
    localSheet.Cells(lineNum, 5).Value = thisCOR6.FinishedItemDesc
    localSheet.Cells(lineNum, 4).Value = thisCOR6.FinishedItem
    Set thisCOR6 = Nothing
    OLDsaveToSAP = localSheet.Cells(lineNum, 6).Value
    If doOtherConfirmations Then
        updateStatus "Entering cut/wash details"
        For j = 1 To 2
            Set thisCOR6 = New clsNewCOR6
            With thisCOR6
                .applyMatl = False
                .oper = 20 * j
                .numHrs = 0.25
                .procOrdNum = processOrdNum
                .saveIt
            End With
            Set thisCOR6 = Nothing
        Next j
    End If
    updateStatus ""
End Function
Sub doHumo(ByVal startRow As Long, t As Long)
    Dim r As Integer 'array position
    Set localSheet = ThisWorkbook.Sheets("Sheet1")

    Dim p As Long 'spreadsheet row we're in
    Dim hampCheck As Boolean 'does this hamper have anything in it
    hampCheck = False
    Dim thisArray() As Variant
    ReDim thisArray(t - startRow)
    r = 0
    For p = startRow To t
        thisArray(r) = localSheet.Cells(p, 8).Value
        hampCheck = True 'this hamper has something in it
        r = r + 1
        
    Next p
    If hampCheck Then
        Dim thisHU As New clsHU
        thisHU.itemNum = localSheet.Cells(startRow, 4).Value
        thisHU.batchArray = removeEmptyElementsFromArray(thisArray)
        thisHU.humo4Gr
        Set thisHU = Nothing
    End If
    
End Sub
Function printCOA(ByVal itemNum As String, ByVal batchNum As String, ByVal lowAlkali As Boolean) As Boolean
    Dim thisSession As GuiSession
    Dim inspPlan As String
    Set thisSession = SAPSession.CurrentSession
    If lowAlkali Then
        inspPlan = "Q_224HV-003"
    Else
        inspPlan = "Q_214HV-003"
    End If
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nqc22"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtP_MATNR").Text = itemNum
        .FindById("wnd[0]/usr/ctxtP_CHARG").Text = batchNum
        .FindById("wnd[0]/usr/ctxtP_SNDWRK").Text = "Q105"
        .FindById("wnd[0]/usr/btnG_V_TEXT").Press
        .FindById("wnd[0]/usr/ctxtP_QCVNR").Text = inspPlan
        .FindById("wnd[0]/usr/ctxtP_QCVKT").Text = "e31b"
        .FindById("wnd[0]/usr/ctxtP_QCVKV").Text = "1"
        .FindById("wnd[0]").SendVKey 8
        .FindById("wnd[1]/usr/radQCPARTYPE-X_NONE").Select
        .FindById("wnd[1]/usr/radQCPARTYPE-X_NONE").SetFocus
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
    End With
End Function



'====================================================================
' Function: findItemDetailsFromBatch
'
' Purpose:
'   Look up a batch in MB51 and return:
'       result(0) = Material number (MATNR)
'       result(1) = Material description (MAKTX)
'       result(2) = Storage location (LGORT)
'
' Input:
'   batch (String) - Batch number to search for in MB51
'
' Return:
'   Variant array (0 To 2)
'       If not found, result(0) will be 0 and a message box is shown.
'
' SAP Transaction:
'   MB51 - Material Document List
'
' Notes:
'   - Filters are set for movement type 101 and the given batch.
'   - Uses the first row of the returned ALV grid.
'   - Assumes an SAP GUI session is already connected.
'====================================================================
Function findItemDetailsFromBatch(ByVal batch As String) As Variant
    On Error GoTo ErrHandler

    Dim thisSession As GuiSession
    Dim result() As Variant

    ' Pre-size result: 0 = MATNR, 1 = MAKTX, 2 = LGORT
    ReDim result(2)

    ' Get current SAP GUI session
    Set thisSession = SAPSession.CurrentSession
    If thisSession Is Nothing Then
        MsgBox "No active SAP session found. Please log in to SAP and try again.", _
               vbCritical, "SAP Session Error"
        result(0) = 0
        findItemDetailsFromBatch = result
        Exit Function
    End If

    With thisSession
        '------------------------------------------------------------
        ' Navigate to MB51
        '------------------------------------------------------------
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb51"
        .FindById("wnd[0]").SendVKey 0

        '------------------------------------------------------------
        ' Set selection criteria:
        '   - MATNR      (material)     : blank
        '   - LGORT      (storage loc.) : blank
        '   - CHARG-LOW  (batch)        : provided batch
        '   - BWART-LOW  (movement)     : 101 (goods receipt from order)
        '   - AUFNR, BUDAT, USNAM, MBLNR: blank
        '   - ALV_DEF    : /S4P_DEFAULT (ALV layout)
        '------------------------------------------------------------
        .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = batch
        .FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "101"
        .FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = ""
        .FindById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
        .FindById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/S4P_DEFAULT"

        ' Optional: ensure "Document list" radio button is selected
        .FindById("wnd[0]/usr/radRFLAT_L").SetFocus

        ' Execute the report
        .FindById("wnd[0]/tbar[1]/btn[8]").Press    ' F8

        '------------------------------------------------------------
        ' Check if any rows returned in ALV grid
        '   - ALV Grid ID: "wnd[0]/usr/cntlGRID1/shellcont/shell"
        '------------------------------------------------------------
        Dim oGrid As Object
        Set oGrid = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")

        If oGrid.RowCount > 0 Then
            ' Optional: select a column (e.g., MBLNR) to avoid "no column" issues
            oGrid.SelectColumn "MBLNR"

            ' Go to line item (if grid detail is needed)
            .FindById("wnd[0]/tbar[1]/btn[40]").Press

            '--------------------------------------------------------
            ' Read first row values:
            '   MATNR - material number
            '   MAKTX - material description
            '   LGORT - storage location
            '--------------------------------------------------------
            result(0) = oGrid.GetCellValue(0, "MATNR")
            result(1) = oGrid.GetCellValue(0, "MAKTX")
            result(2) = oGrid.GetCellValue(0, "LGORT")
        Else
            ' No rows returned for this batch
            MsgBox "I couldn't find batch number " & batch & " (movement 101).", _
                   vbCritical, "Batch Not Found"
            result(0) = 0
        End If

        '------------------------------------------------------------
        ' Back out of MB51 to return to the previous screen
        '   - First back: from item/detail to list
        '   - Second back: out of MB51
        '------------------------------------------------------------
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With

CleanExit:
    ' Return result and clean up
    findItemDetailsFromBatch = result
    Set thisSession = Nothing
    Exit Function

ErrHandler:
    ' Central error handler
    MsgBox "Error in findItemDetailsFromBatch for batch: " & batch & vbCrLf & _
           "Description: " & Err.description & vbCrLf & _
           "Number: " & Err.Number, _
           vbCritical, "SAP Automation Error"

    ' On error, ensure at least result(0) = 0 so caller can test
    On Error Resume Next
    result(0) = 0
    findItemDetailsFromBatch = result
    Set thisSession = Nothing
End Function




' this can be deleted after new one runs well

Function OLDfindItemDetailsFromBatch(ByVal batch As String) As Variant
    Dim thisSession As GuiSession
    Dim result() As Variant
    ReDim result(2)

    Set thisSession = SAPSession.CurrentSession
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb51"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = batch
        .FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "101"
        .FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = ""
        .FindById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
        .FindById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/S4P_DEFAULT"
        .FindById("wnd[0]/usr/radRFLAT_L").SetFocus
        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        If .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount > 0 Then
            .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectColumn "MBLNR"
            .FindById("wnd[0]/tbar[1]/btn[40]").Press
            result(0) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MATNR")
            result(1) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MAKTX")
            result(2) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "LGORT")
        Else
            MsgBox "I couldn't find batch number " & batch & ".", vbCritical
            result(0) = 0
        End If
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press

    End With
    Set thisSession = Nothing
    OLDfindItemDetailsFromBatch = result
End Function


'====================================================================
' Function: lookupRecipeGroup
'
' Purpose:
'   For a given material (itemNum), call C223 to find the
'   recipe group / group number (PLNNR) from the production
'   version list.
'
' Input:
'   itemNum (String) - Finished good material number.
'
' Return:
'   String - Recipe group (PLNNR) for the material
'            or "" if not found / not run (offline or error).
'
' SAP Transaction:
'   C223 - Display Production Version
'
' Assumptions:
'   - SAP GUI scripting is enabled and an active session exists.
'   - Plant "Q105" is the correct plant for this lookup.
'   - The desired recipe group is on the first row of the grid.
'
' Side conditions:
'   - Only runs if global flag onlineMode = True.
'====================================================================
Function lookupRecipeGroup(ByVal itemNum As String) As String
    On Error GoTo ErrHandler

    Dim thisSession As GuiSession
    Dim recipeGroup As String

    ' Default return value
    recipeGroup = ""

    '------------------------------------------------------------
    ' Only run if onlineMode is True
    '   (Assumes onlineMode is a module-level/global Boolean)
    '------------------------------------------------------------
    If Not onlineMode Then
        lookupRecipeGroup = ""     ' Explicitly return empty if offline
        Exit Function
    End If

    '------------------------------------------------------------
    ' Get active SAP session
    '------------------------------------------------------------
    Set thisSession = SAPSession.CurrentSession
    If thisSession Is Nothing Then
        MsgBox "No active SAP session found. Please log in to SAP " & _
               "before running lookupRecipeGroup.", _
               vbCritical, "SAP Session Error"
        lookupRecipeGroup = ""
        Exit Function
    End If

    With thisSession
        '--------------------------------------------------------
        ' Navigate to C223 - Display Production Version
        '--------------------------------------------------------
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nc223"
        .FindById("wnd[0]").SendVKey 0     ' Enter

        '--------------------------------------------------------
        ' Selection Screen Fields (SAPLCMFV:1100)
        '
        '   MKAL-WERKS : Plant             ? "Q105"
        '   MKAL-MATNR : Material number   ? itemNum
        '   MKAL_ADMIN-DISPO : MRP controller   (blank)
        '   MKAL_ADMIN-STTAG : Valid-from date (blank)
        '   MKAL-PLNNR : Recipe group           (blank; we want to read it)
        '   MKAL-MDV01 : Planning indicator     (blank)
        '--------------------------------------------------------
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-WERKS").Text = "Q105"
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MATNR").Text = itemNum
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL_ADMIN-DISPO").Text = ""
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL_ADMIN-STTAG").Text = ""
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-PLNNR").Text = ""
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").Text = ""

        ' Set focus before executing (helps avoid some SAP quirks)
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").SetFocus
        .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").CaretPosition = 0

        ' Execute (e.g. F8 or Enter depending on screen behavior)
        .FindById("wnd[0]").SendVKey 0

        '--------------------------------------------------------
        ' Grid on result screen (SAPLCMFV:1200)
        '   Table control: SAPLCMFVT_MKAL
        '   Column: MKAL_EXPAND-PLNNR (recipe group)
        '
        '   Here you are reading the first row (row 0)
        '   from column index 18 in the table control.
        '
        '   Note: If there is a chance of no rows,
        '         a RowCount check could be added here.
        '--------------------------------------------------------
        Dim oTbl As Object
        Set oTbl = .FindById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL")

        ' Make sure at least one row exists before reading
        If oTbl.RowCount > 0 Then
            oTbl.GetCell(0, "MKAL_EXPAND-PLNNR").SetFocus
            recipeGroup = .FindById( _
                "wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/" & _
                "tblSAPLCMFVT_MKAL/ctxtMKAL_EXPAND-PLNNR[18,0]").Text
        Else
            ' No production versions found for this material
            recipeGroup = ""
        End If

        '--------------------------------------------------------
        ' Back out of C223 to leave SAP in a clean state
        '   - First back out from list
        '   - Second back out of transaction
        '--------------------------------------------------------
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With

    ' Set function return value
    lookupRecipeGroup = recipeGroup

CleanExit:
    On Error Resume Next
    Set thisSession = Nothing
    On Error GoTo 0
    Exit Function

ErrHandler:
    MsgBox "Error in lookupRecipeGroup for item: " & itemNum & vbCrLf & _
           "Description: " & Err.description & vbCrLf & _
           "Number: " & Err.Number, _
           vbCritical, "SAP Automation Error"

    lookupRecipeGroup = ""
    Resume CleanExit
End Function


Function OLDlookupRecipeGroup(ByVal itemNum As String) As String
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    With thisSession
        If onlineMode Then
            .FindById("wnd[0]/tbar[0]/okcd").Text = "/nc223"
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-WERKS").Text = "q105"
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MATNR").Text = itemNum
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL_ADMIN-DISPO").Text = ""
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL_ADMIN-STTAG").Text = ""
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-PLNNR").Text = ""
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").Text = ""
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").SetFocus
            .FindById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MDV01").CaretPosition = 0
            .FindById("wnd[0]").SendVKey 0
            
            .FindById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL/ctxtMKAL_EXPAND-PLNNR[18,0]").SetFocus
            OLDlookupRecipeGroup = .FindById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL/ctxtMKAL_EXPAND-PLNNR[18,0]").Text
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
        End If
    End With
    Set thisSession = Nothing
End Function
Function inspectPackOperation(ByVal itemNum As String) As String
    Dim recipeGroup As String
    recipeGroup = lookupRecipeGroup(itemNum)
    Dim result As String
    If recipeGroup = "Q5HVL000" Then
        result = "20"
    ElseIf recipeGroup = "Q5HVC000" Then
        result = "60"
    Else
        MsgBox "Recipe Group Not Found."
        result = "40"
    End If
    inspectPackOperation = result
    
End Function
Sub printHULabel(ByVal huNum As String, ByVal numLabels As Integer)
    Dim thisLabel As GuiSession
    Set thisLabel = SAPSession.CurrentSession
    With thisLabel
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nzpplabel"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/btn%P150004_1000").Press
        .FindById("wnd[0]/usr/ctxtS_HUNO-LOW").Text = huNum
        .FindById("wnd[0]/usr/cmbP_ORIGIN").Key = "USA"
        .FindById("wnd[0]/usr/txtP_NUM").Text = numLabels
        .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName

        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisLabel = Nothing
End Sub
Function removeEmptyElementsFromArray(originalArray As Variant) As Variant
    
    Dim nonBlankCount As Long
    nonBlankCount = 0
    Dim i As Long
    
    For i = LBound(originalArray) To UBound(originalArray)
        If originalArray(i) <> "" Then
            nonBlankCount = nonBlankCount + 1
        End If
    Next i
    
    ' 3. Create a new, temporary array with the correct size.
    Dim tempArray() As Variant
    ReDim tempArray(0 To nonBlankCount - 1)
    
    ' 4. Populate the new array with only the non-empty values.
    Dim j As Long
    j = 0
    
    For i = LBound(originalArray) To UBound(originalArray)
        If originalArray(i) <> "" Then
            tempArray(j) = originalArray(i)
            j = j + 1
        End If
    Next i
    
    ' 5. Replace the original array with the new one.
    '    Note: This is an important step to "remove" the elements.
    removeEmptyElementsFromArray = tempArray
    
    
End Function
Sub updateStatus(ByVal statusTxt As String)
    PackForm.Label20.Caption = statusTxt
End Sub
Sub LogEvent(message As String, Optional messageType As String = "Transaction")
    Dim filePath As String
    ' Use the correct path separator
    If Not productionMode Then
        filePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Furnace Deck\TransactionLog.csv"
    Else
        filePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\TransactionLog.csv"
    End If
        
    ' Check if the file exists
    Dim fileNum As Integer
    fileNum = FreeFile()
    On Error GoTo ErrorHandler ' Add error handling
    Open filePath For Append As #fileNum
    
    Print #fileNum, Format(Date, "mm/dd/yyyy") & "," & Time() & "," & Environ("computername") & ",HV Pack Application," & Environ$("UserName") & "," & messageType & "," & message
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.description
End Sub
Sub SaveAndCloseWorkbook()
    LogEvent "Closing Workbook", "Application"
    ' Save the active workbook
    ActiveWorkbook.Save

    ' Close the active workbook
    ActiveWorkbook.Close

End Sub
Function FindLastOccurrence(searchString As String, searchColumn As Range) As Range
    '
    ' This function finds the last cell containing a given text string within a specified column.
    '
    ' Parameters:
    '   searchString: The text string to search for.
    '   searchColumn: The column range to search within (e.g., Range("A:A")).
    '
    ' Returns:
    '   A Range object representing the found cell. Returns Nothing if the string is not found.
    '

    ' Declare a variable to hold the found range.
    Dim foundCell As Range

    ' Use the Find method. We start from the last cell in the column and search backwards.
    ' xlPrevious specifies the search direction.
    Set foundCell = searchColumn.Find(What:=searchString, _
                                      After:=searchColumn.Cells(1, 1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlPart, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlPrevious, _
                                      MatchCase:=False) ' Set MatchCase to True for case-sensitive search

    ' Return the found cell object.
    Set FindLastOccurrence = foundCell
    
End Function

