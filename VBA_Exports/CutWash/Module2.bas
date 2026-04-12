Attribute VB_Name = "Module2"
Option Explicit
Sub saveToSAP(ByVal lineNum As Long)
    ' Declare object variables for SAP interaction and workbook/worksheet manipulation
    Dim thisCOR6 As Object 'Changed to Object for late binding if clsCOR6 is not always available or for broader compatibility
    Dim thisWB As Workbook
    Dim itemNum As String
    Dim bakedItemNum As Long
    Dim dataWS As Worksheet
    Dim settingsWS As Worksheet
    Dim thisItem As Object 'Changed to Object for late binding if clsMM is not always available
    Dim orderCheck As New clsCOOISPI
    Dim batchNum As String
    ' Declare other variables
    Dim processOrderNumber As String
    Dim matlArray() As Variant ' Array to hold material data for SAP
    Dim fsArray() As Variant   ' Array to hold data from useMB52 function
    Dim strErrorMessage As String ' To store a user-friendly error message
    Dim isSerialized As Boolean
    Dim x As Integer
    Dim newOrders As Integer
    ' --- Error Handling Setup ---
    On Error GoTo ErrorHandler

    ' --- Initialize Workbook and Worksheet Objects ---
    Set thisWB = ThisWorkbook
    ' Attempt to set worksheet objects, will error if sheets don't exist
    Set dataWS = thisWB.Sheets("Sheet1")
    Set settingsWS = thisWB.Sheets("Sheet2")
    itemNum = dataWS.Cells(lineNum, 9)
    If orderCheck.numOrders(itemNum) = 0 Then
        newOrders = InputBox("There aren't any open process orders for " & dataWS.Cells(lineNum, 9) & ". Please wait while I create some. By default I will create 10. If you want to change that count, enter it below.", "Create Process Orders", 10)
        ActiveWindow.WindowState = xlMinimized
        For x = 1 To newOrders
            createProcessOrder CLng(dataWS.Cells(lineNum, 9))
        Next x
        ActiveWindow.WindowState = xlNormal
    End If
    processOrderNumber = orderCheck.getProcOrdNum(itemNum, isSerialized, dataWS.Cells(lineNum, 6))
    bakedItemNum = dataWS.Cells(lineNum, 4).Value
    batchNum = dataWS.Cells(lineNum, 6).Value
    isSerialized = dataWS.Cells(lineNum, 16).Value
    
    ' --- Instantiate SAP Interaction Class ---
    ' clsCOR6 is a custom class for interacting with SAP transaction COR6
    Set thisCOR6 = New clsCOR6
    thisCOR6.finalConfirmation = True
    thisCOR6.serialized = isSerialized
    ' --- Determine Material Array Size based on presence of a second material ---
    ' Check if a second material (column 14, "N") is specified
    If dataWS.Cells(lineNum, 14).Value <> "" Then
        ' If second material exists, dimension matlArray to hold two materials (index 0 and 1)
        ' Each material has 5 properties (0 to 4)
        ReDim matlArray(1, 4)
        ' Populate details for the second material
        matlArray(1, 0) = dataWS.Cells(lineNum, 14).Value ' Material Number
        matlArray(1, 1) = ""                              ' Batch (assuming blank for this material)
        matlArray(1, 2) = 1                               ' Quantity
        matlArray(1, 3) = "0100"                          ' Storage Location (example value)
        matlArray(1, 4) = "531"                           ' Movement Type (goods receipt for by-product)
        thisCOR6.applyNumMatl = 2 ' Inform clsCOR6 object there are two materials
    Else
        ' If no second material, dimension matlArray for one material (index 0)
        ReDim matlArray(0, 4)
        thisCOR6.applyNumMatl = 1 ' Inform clsCOR6 object there is one material
    End If

    ' --- Populate First Material Details ---
    ' Retrieve batch stock information using a function (useMB52)
    ' This function queries SAP (e.g., MB52 transaction) for stock details
    fsArray = useMB52(batchNum, bakedItemNum)
    If isSerialized Then
        dataWS.Cells(lineNum, 4) = fsArray(1)
        dataWS.Cells(lineNum, 5) = fsArray(2)
    End If
    ' Populate details for the first material (component)
    matlArray(0, 0) = bakedItemNum  ' Material Number (Component)
    matlArray(0, 1) = batchNum  ' Batch (Component)
    matlArray(0, 2) = dataWS.Cells(lineNum, 7).Value  ' Quantity (Component)
    ' fsArray(0) returns the storage location from useMB52
    If UBound(fsArray) >= 0 And Not IsEmpty(fsArray(0)) Then
        matlArray(0, 3) = fsArray(0) ' Storage Location (Component)
    Else
        ' Handle case where fsArray might not be populated as expected
        strErrorMessage = "Could not retrieve storage location for material: " & bakedItemNum & ", Batch: " & batchNum
        GoTo ErrorHandler ' Trigger the error handler with a custom message
    End If
    matlArray(0, 4) = "261"                           ' Movement Type (goods issue for order)

    ' --- Set Yield and Process Order Number ---
    thisCOR6.yield = dataWS.Cells(lineNum, 12).Value ' Production yield quantity
    
    ' --- Update Item Description if Missing ---
    ' If item description (column 5, "E") is missing, fetch it
    If dataWS.Cells(lineNum, 5).Value = "" Then
        Set thisItem = New clsMM ' clsMM is likely a custom class for Material Master (MM) data
        ' Call a method to get material description
        dataWS.Cells(lineNum, 5).Value = thisItem.getDescription(bakedItemNum)
        Set thisItem = Nothing ' Release the object
    End If

    ' --- Update Spreadsheet with Process Order Number ---
    dataWS.Cells(lineNum, 8).Value = processOrderNumber ' Column "H"
    
    ' --- Set Properties for SAP COR6 Interaction ---
    thisCOR6.procOrdNum = processOrderNumber
    thisCOR6.tubeSerialNumber = batchNum ' Batch number used as serial
    ' Calculate labor hours (example: 0.25 hours per unit of yield)
    thisCOR6.numHrs = 0.25 * dataWS.Cells(lineNum, 12).Value
    thisCOR6.matlArray = matlArray ' Assign the populated materials array
    thisCOR6.applyMatl = True      ' Flag to indicate materials should be processed

    ' --- Perform SAP Save Operation and Update Spreadsheet ---
    ' Call the saveIt method of the clsCOR6 object, which presumably performs the SAP posting
    ' The saveIt method should return True on success, False on failure
    If thisCOR6.saveIt Then
        dataWS.Cells(lineNum, 13).Value = "Yes" ' Mark as "Saved to SAP" in column "M"
    Else
        ' If saveIt returns False, it implies a failure in SAP posting handled within the class
        ' You might want to add more specific error feedback here if clsCOR6 provides it
        strErrorMessage = "Failed to save data to SAP for Process Order: " & processOrderNumber & ". Check SAP or clsCOR6 logs."
        GoTo ErrorHandler ' Trigger the error handler
    End If

    ' Update the finished item description in the spreadsheet (column "J")
    dataWS.Cells(lineNum, 10).Value = thisCOR6.FinishedItemDesc
    bakedItemNum = dataWS.Cells(lineNum, 4).Value
    If dataWS.Cells(lineNum, 16).Value = True Then
        Dim thisHU As New clsHU
        thisHU.huGI processOrderNumber, bakedItemNum, batchNum
        Set thisHU = Nothing
    End If
    ActiveWorkbook.Save
CleanExit:
    ' --- Release Object Variables ---
    Set thisCOR6 = Nothing
    Set orderCheck = Nothing
    Set dataWS = Nothing
    Set settingsWS = Nothing
    Set thisWB = Nothing
    Set thisItem = Nothing ' Ensure thisItem is also cleared here
    Exit Sub ' Normal exit

ErrorHandler:
    ' --- Error Handling Block ---
    Dim strFullErrorMessage As String
    If strErrorMessage <> "" Then
        ' Display a custom error message if one was set
        strFullErrorMessage = strErrorMessage
    Else
        ' Display the standard VBA error message
        strFullErrorMessage = "An error occurred in 'saveToSAP'." & vbCrLf & _
                              "Error Number: " & Err.Number & vbCrLf & _
                              "Error Description: " & Err.Description & vbCrLf & _
                              "Processing Line: " & lineNum
    End If

    MsgBox strFullErrorMessage, vbCritical, "VBA Runtime Error"

    ' It's good practice to log errors to a file or a specific sheet for later review,
    ' especially for unattended processes. Example:
    ' On Error Resume Next ' Temporarily ignore errors during error logging
    ' Dim logWS As Worksheet
    ' Set logWS = ThisWorkbook.Sheets("ErrorLog") ' Assuming an "ErrorLog" sheet exists
    ' If Not logWS Is Nothing Then
    '     Dim nextLogRow As Long
    '     nextLogRow = logWS.Cells(Rows.Count, 1).End(xlUp).Row + 1
    '     logWS.Cells(nextLogRow, 1).Value = Now
    '     logWS.Cells(nextLogRow, 2).Value = "saveToSAP"
    '     logWS.Cells(nextLogRow, 3).Value = lineNum
    '     logWS.Cells(nextLogRow, 4).Value = Err.Number
    '     logWS.Cells(nextLogRow, 5).Value = Err.Description
    '     If strErrorMessage <> "" Then logWS.Cells(nextLogRow, 6).Value = strErrorMessage
    ' End If
    ' On Error GoTo 0 ' Reset error handling

    ' Clear the error object
    Err.Clear

    ' Go to CleanExit to ensure all objects are released
    GoTo CleanExit
End Sub

Public Sub catchUpSAP()
    Dim x As Integer
    Dim startingRow As Integer
    Dim keepGoing As Boolean
    keepGoing = True
    x = InputBox("What row would you like to start with?")
    startingRow = x
    If dataWS.Cells(x, 13).Value = "Yes" Then
        keepGoing = False
        saveRemoteData startingRow, x
    Else
        saveToSAP x
        x = x + 1
    End If
End Sub
Function doesPOwithBatchItemExist(ByVal itemNum As Long, ByVal batchNum As String) As String
'this returns the process order number for the batch/item combo if it exists
End Function
Function getProcOrdNumXX(itemNum As String, unused As Boolean, displayOnly As Boolean) As String
    Dim thisSession As GuiSession
    Dim x As Integer
    Dim result As String
On Error GoTo ErrHandler
    Set thisSession = SAPSession.CurrentSession
'    setStartScreen
    unused = False
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nCOOISPI"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = itemNum
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco"
    If unused Then
        thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "gmps" ' Set system status 2 to "gmps"
    
        thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYOP1").SetFocus  ' makes sure we're only seeing released orders
        thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYOP1").CaretPosition = 3
        thisSession.FindById("wnd[0]").SendVKey 4
        thisSession.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
        thisSession.FindById("wnd[1]/usr/lbl[1,4]").CaretPosition = 0
        thisSession.FindById("wnd[1]").SendVKey 2
    End If
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    
    x = 0
    If SAPStatus.CurrentStatus = "There is no data for the selection" Then
        MsgBox "There are no open process orders for " & itemNum, vbOKOnly
        getProcOrdNumXX = 9
        Exit Function
    End If
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellRow = x
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellColumn = "AUFNR"
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
    result = thisSession.FindById("wnd[0]/usr/txtCAUFVD-AUFNR").Text
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    If Not displayOnly Then
        thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
        thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Else
        ActiveWindow.WindowState = xlMinimized
    End If
    Set thisSession = Nothing
    getProcOrdNumXX = result
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getProcOrdNum of Module2"
End Function

Sub printDailyLog(userInput As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterDate As Date
    Dim headerDateString As String
    Dim FieldNum As Long
    Dim visibleRowsCheck As Range

    ' --- Configuration ---
    Const SHEET_NAME As String = "Sheet1"  ' Name of the worksheet
    Const TABLE_NAME As String = "Table1"  ' Name of the table
    Const COLUMN_TO_FILTER As Long = 1     ' Column A is the 1st column
    ' -------------------

    ' Set the worksheet object
    On Error Resume Next ' In case the sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0 ' Turn error handling back on
    If ws Is Nothing Then
        MsgBox "Error: Worksheet named '" & SHEET_NAME & "' not found.", vbCritical
        Exit Sub
    End If

    ' Set the table object
    On Error Resume Next ' In case the table doesn't exist
    Set tbl = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0 ' Turn error handling back on
    If tbl Is Nothing Then
        MsgBox "Error: Table named '" & TABLE_NAME & "' not found on sheet '" & SHEET_NAME & "'.", vbCritical
        Set ws = Nothing ' Clean up
        Exit Sub
    End If

    ' --- Get Date Input from User ---
   ' userInput = application.InputBox("Enter the date to filter by (e.g., MM/DD/YYYY):", "Filter Date", Type:=2) ' Type 2 for Text input

    ' Check if user cancelled
    If userInput = "False" Then ' InputBox returns string "False" if cancelled
        MsgBox "Operation cancelled by user.", vbInformation
        GoTo CleanUp ' Jump to cleanup code
    End If

    ' Check if the input is a valid date
    If Not IsDate(userInput) Then
        MsgBox "Invalid date entered: '" & userInput & "'. Please enter a valid date format (e.g., MM/DD/YYYY).", vbExclamation
        GoTo CleanUp ' Jump to cleanup code
    End If

    ' Convert the valid input string to a Date type
    filterDate = CDate(userInput)
    headerDateString = Format(filterDate, "MM/DD/YYYY") ' Format for header

    ' --- Apply Filter ---
    FieldNum = COLUMN_TO_FILTER

    ' Clear any existing filters on the table first
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    Else
         ' If autofilter was somehow turned off for the table, ensure it's on
         If Not ws.AutoFilterMode Then ' Avoid error if filter arrows aren't showing
            tbl.Range.AutoFilter
         End If
    End If

    ' Apply the date filter.
    ' To robustly filter for a specific day (ignoring time), filter for >= Date and < Date+1
    tbl.Range.AutoFilter Field:=FieldNum, _
                         Criteria1:=">=" & CDbl(filterDate), _
                         Operator:=xlAnd, _
                         Criteria2:="<" & CDbl(filterDate + 1)


    ' --- Check if any data is visible after filtering ---
    On Error Resume Next ' Need error handling in case NO data rows are visible
    Set visibleRowsCheck = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' Turn error handling back on

    If visibleRowsCheck Is Nothing Then
        MsgBox "No records found for the date: " & headerDateString & "." & vbCrLf & "Nothing will be printed.", vbInformation
        ' Optional: Clear the filter again if nothing found
         tbl.AutoFilter.ShowAllData
         Range("A" & nextRow).Select
    Else
        ' --- Set Print Header ---
        With ws.PageSetup
            .LeftHeader = "" ' Clear potentially existing headers
            .RightHeader = ""
            ' Set Center Header with Font Size 14 and the date
            .CenterHeader = "&14" & headerDateString ' &14 sets font size
            ' You might want to set other print settings here, e.g.:
             .Orientation = xlLandscape
            ' .Zoom = False
            ' .FitToPagesWide = 1
            ' .FitToPagesTall = False
        End With

        ' --- Print the Worksheet (Excel inherently prints only visible rows when filtered) ---
        ws.PrintOut
        'clear the filter
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=1
        'go to end of data
        Range("A" & nextRow).Select
        ' Optional: Inform user printing is done
        ' MsgBox "Filtered data for " & headerDateString & " sent to printer.", vbInformation

        ' Note: Printing the whole sheet (ws.PrintOut) is common when filtering.
        ' If you ONLY want the table itself printed, potentially without other sheet elements,
        ' you could use: tbl.Range.PrintOut
    End If


CleanUp:
    ' Release object variables
    Set visibleRowsCheck = Nothing
    Set tbl = Nothing
    Set ws = Nothing

End Sub

Sub printLabel(ByVal procOrdNum As String, ByVal itemNum As Long, ByVal batch As String, numTickets As Integer)
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    setStartScreen
    
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "zpplabel"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/btn%P010002_1000").Press
    thisSession.FindById("wnd[0]/usr/ctxtP_ORDER").Text = procOrdNum
    thisSession.FindById("wnd[0]/usr/ctxtP_MATNR").Text = itemNum
    thisSession.FindById("wnd[0]/usr/ctxtP_CHARG").Text = batch
    thisSession.FindById("wnd[0]/usr/txtP_IGMNG").Text = "1"
    thisSession.FindById("wnd[0]/usr/ctxtP_GMEIN").Text = "ea"
    thisSession.FindById("wnd[0]/usr/cmbP_ORIGIN").Key = "USA"
    thisSession.FindById("wnd[0]/usr/txtP_NUM").Text = numTickets
    thisSession.FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    If SAPStatus.IsThereAModalWindow Then thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
'    thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing

End Sub
Sub MIGO551(ByVal itemNum As Long, ByVal batch As String, qty As Integer)
    Dim moveType As Integer
    Dim uom As String
    Dim reason As String
    Dim location As String
    Dim costCenter As String
    Dim thisGUI As GuiSession
    Set thisGUI = SAPSession.CurrentSession
    setStartScreen
    With thisGUI
        .FindById("wnd[0]/tbar[0]/okcd").Text = "migo"
        .FindById("wnd[0]").SendVKey 0
    
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        
            reason = "LD End User Request"
            costCenter = "NW4455"
            uom = "EA"
            
            .FindById("wnd[0]/tbar[0]/okcd").Text = "migo"
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").Text = Left(reason, 25)
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").Text = itemNum
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").SetFocus
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").CaretPosition = 6
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").Select
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").Text = Abs(qty)
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/ctxtGOITEM-ERFME").Text = uom
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/ctxtGOITEM-ERFME").SetFocus
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/ctxtGOITEM-ERFME").CaretPosition = 2
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").Select
            If qty < 0 Then
                moveType = 551
            Else
                moveType = 552
            End If
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").Text = moveType
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").Text = "Q105"
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").Text = "HVWS"
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").SetFocus
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").CaretPosition = 4
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").Text = reason
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_BATCH").Select
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_BATCH/ssubSUB_TS_GOITEM_BATCH:SAPLMIGO:0335/ctxtGOITEM-CHARG").Text = batch
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_BATCH/ssubSUB_TS_GOITEM_BATCH:SAPLMIGO:0335/ctxtGOITEM-CHARG").SetFocus
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_BATCH/ssubSUB_TS_GOITEM_BATCH:SAPLMIGO:0335/ctxtGOITEM-CHARG").CaretPosition = 9
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT").Select
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = costCenter
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").SetFocus
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").CaretPosition = 6
            .FindById("wnd[0]").SendVKey 0
    
    
    .FindById("wnd[0]/tbar[1]/btn[23]").Press
    End With
End Sub
Function IsWorkBookOpen(ByVal FilePath As String) As Boolean

  Dim wb As Workbook

  ' Check if a file path was provided
  If Len(FilePath) > 0 Then
    ' Loop through all open workbooks
    For Each wb In Workbooks
      ' Compare the full file paths, handling potential case differences
        If LCase(wb.FullName) = LCase(FilePath) Then
            IsWorkBookOpen = True
            Exit Function ' Workbook found, exit the function
        End If
    Next wb
  End If

  ' If no match was found, the workbook is not open
  IsWorkBookOpen = False

End Function

Function nextRow() As Long
    Dim x As Long
    On Error GoTo ErrHandler
    x = 1
    Dim keepGoing As Boolean
    keepGoing = True
    Do While keepGoing
        x = x + 1
        If Sheet1.Range("A" & x).Value = "" Then keepGoing = False
    Loop
    nextRow = x
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextRow of Module2"
End Function

Public Sub changeUserDefaults()
    ' sets the number format, date, and time formats
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "su3"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
        .FindById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key = "X"
        .FindById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").Key = "2"
        .FindById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-TIMEFM").Key = "0"
        .FindById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-TIMEFM").SetFocus
        .FindById("wnd[0]/tbar[0]/btn[11]").Press

    End With
    Set thisSession = Nothing
    
End Sub
Sub saveRemoteData(ByVal startRowNum As Long, ByVal endRowNum As Long)
    Dim col As Integer
    Dim rw As Integer
    Dim remoteWB As Workbook
    
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
On Error GoTo ErrHandler
        
        
                ' --- Open the external workbook ---
  '      If IsWorkBookOpen(remFile) Then ' Check if the workbook is already open
            'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
  '      Else
            Set remoteWB = Workbooks.Open(fileName:=remFile, ReadOnly:=False, ignorereadonlyrecommended:=True)  ' Open the workbook if it's not already open
  '      End If
    
        
        Set thisSheet = thisWB.Sheets("Sheet1")
        
        
        Set remoteSheet = remoteWB.Sheets("Sheet1")
        For rw = startRowNum To endRowNum
            For col = 1 To 15
                remoteSheet.Cells(rw, col).Value = thisSheet.Cells(rw, col)
    '            remoteSheet.Cells(rw, col).Select
            Next col
        Next rw
        
        remoteWB.Close SaveChanges:=True

    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteData of Module2"
End Sub


Function useMB52(ByVal sBatch As String, ByVal itemNum As String) As Variant
'this finds the location from the item and batch
    Dim thisGUI As GuiSession
On Error GoTo ErrHandler

    Set thisGUI = SAPSession.CurrentSession
    setStartScreen
    Dim x As Integer
    Dim b As Integer
    Dim xDetails() As Variant
    ReDim xDetails(2)
    'location   itemNumber
    thisGUI.FindById("wnd[0]/tbar[0]/okcd").Text = "mb52"
    thisGUI.FindById("wnd[0]").SendVKey 0
    If itemNum = 0 Then itemNum = ""
    thisGUI.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = itemNum
    thisGUI.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "Q105"
    thisGUI.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = sBatch
    thisGUI.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    thisGUI.FindById("wnd[0]/usr/chkNOZERO").Selected = True
    thisGUI.FindById("wnd[0]/usr/radPA_FLT").Select
    thisGUI.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/PARKINSONJ"
    thisGUI.FindById("wnd[0]/tbar[1]/btn[8]").Press
    If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
        MsgBox "There is 0 in inventory for batch " & sBatch & ".", vbOKOnly
        Err.Raise 0
    End If
    b = thisGUI.FindById("wnd[0]/usr").Children.Count / 8 - 1
    ' I need to figure out how many pieces of data there are
          
    x = 3
    'itemnum
    xDetails(1) = thisGUI.FindById("wnd[0]/usr/lbl[1," & x & "]").Text
    'location
    xDetails(0) = thisGUI.FindById("wnd[0]/usr/lbl[48," & x & "]").Text
    'description
    xDetails(2) = thisGUI.FindById("wnd[0]/usr/lbl[17," & x & "]").Text
    Set thisGUI = Nothing
    useMB52 = xDetails()

Exit Function
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in useMB52 of Module2"

End Function
Function createProcessOrder(itemNum As Long) As String
    Dim cor1GUI As GuiSession
    Set cor1GUI = SAPSession.CurrentSession
'    setStartScreen
    With cor1GUI
       
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/ncor1"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = itemNum
        .FindById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = "Q105"
        .FindById("wnd[0]/usr/ctxtAUFPAR-PI_AUFART").Text = "ZQ01"
        .FindById("wnd[0]").SendVKey 0
        
   '     .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").Text = 1
        .FindById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/ctxtCAUFVD-GSTRP").Text = Format(Date, "mm/dd/yyyy")
        
        .FindById("wnd[0]/tbar[1]/btn[30]").Press
        If Left(SAPStatus.CurrentStatus, 14) = "Order quantity" Then .FindById("wnd[0]/tbar[1]/btn[30]").Press
        If SAPStatus.IsThereAModalWindow Then
            If SAPStatus.PopUpDialogText = "Classification of the batch is not complete. Do you want to save the batch anyway?" Then
                .FindById("wnd[1]/usr/btnPBUTTON01").Press
            Else
                .FindById("wnd[1]/tbar[0]/btn[0]").Press
                .FindById("wnd[1]/usr/btnSPOP-VAROPTION1").Press
            End If
        End If
        If SAPStatus.IsThereAModalWindow Then
            If SAPStatus.PopUpDialogText = "Classification of the batch is not complete. Do you want to save the batch anyway?" Then
                .FindById("wnd[1]/usr/btnPBUTTON01").Press
            End If
        End If
        
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
'        .FindById("wnd[1]/usr/btnSPOP-OPTION1").Press
        createProcessOrder = SAPStatus.CurrentStatus
        
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set cor1GUI = Nothing
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
