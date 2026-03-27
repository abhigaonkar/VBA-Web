Attribute VB_Name = "Module2"
Option Explicit
Sub testNewMB52()
    Dim mb52 As New clsFIFO
    mb52.useMB52 ("nf31")
    Debug.Print mb52.cSandAtFurnace
    Set mb52 = Nothing
End Sub

Sub SaveAndCloseWorkbook()
    LogEvent "Closing workbook", "Application"
    ' Save the active workbook
    ActiveWorkbook.Save

    ' Close the active workbook
    ActiveWorkbook.Close

End Sub
' This subroutine updates a visual sand level indicator on a user form.
' It reads the amount of sand from a cell in "Sheet2" and displays
' a corresponding number of illuminated "lights" on the form.
Function getSandLevel() As Variant

    ' Declare variables
    If onlineMode Then
        Dim thisSandCheck As New clsFIFO
        thisSandCheck.useMB52 furnName
        getSandLevel = thisSandCheck.cSandAtFurnace
        Set thisSandCheck = Nothing
     End If

End Function


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
Sub ketchup()
    Dim y As Long, numTransactions As Integer
    y = InputBox("What line shall we start on?")
    numTransactions = InputBox("How many transactions?")
    enterSavedData y, numTransactions
    Dim t As Long
    For t = y To y + numTransactions
        saveRemoteData t, True
    Next t

End Sub

Sub rowAlert(p As Long)
Attribute rowAlert.VB_Description = "This will turn a row red to alert us to a possible issue."
Attribute rowAlert.VB_ProcData.VB_Invoke_Func = " \n14"
'
' rowAlert Macro
' This will turn a row red to alert us to a possible issue.
'

'
    Range("A" & p, "O" & p).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub updateFurnaceData(whatSheet As Worksheet, ByVal processOrderNumber As String, ByVal orderData As Variant, itemData As String, ByVal furnNum As Integer)
    With whatSheet
        .Cells(2, -10 + furnNum) = processOrderNumber
        .Cells(3, -10 + furnNum) = orderData(0)
        .Cells(4, -10 + furnNum) = orderData(1)
        .Cells(5, -10 + furnNum) = orderData(2)
        .Cells(6, -10 + furnNum) = orderData(3)
        .Cells(7, -10 + furnNum) = orderData(4)
        .Cells(8, -10 + furnNum) = orderData(5)
        .Cells(9, -10 + furnNum) = orderData(6)
    End With
    'ActiveWorkbook.Save
End Sub
Function enterSavedData(whereStart As Long, numTrans As Integer) As Boolean
    Dim p As Long, y As Integer
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")

    y = whereStart + numTrans - 1
    For p = whereStart To y
        statusUpdate "Checking available sand"
        Dim availableSand As Variant
        Dim sand As New clsFIFO
        availableSand = sand.useMB52(localWS.Cells(p, 1).Value)
        sSandQty = sand.cSandAtFurnace
        Set sand = Nothing
        'how much sand in SAP before transaction
        localWS.Cells(p, 16).Value = sSandQty
        statusUpdate "Doing COR6"
        Dim thisEntry As New clsCOR6
        thisEntry.procOrdNum = localWS.Cells(p, 5).Value
        thisEntry.oper = "20"
        thisEntry.yield = localWS.Cells(p, 9).Value
        thisEntry.numHrs = localWS.Cells(p, 10).Value
        thisEntry.equipmentUsed = "FR" & right(localWS.Cells(p, 1).Value, 2)
        thisEntry.location = localWS.Cells(p, 1).Value
        If localWS.Cells(p, 13).Value <> "" Then
            thisEntry.scrapCode = localWS.Cells(p, 13).Value
            thisEntry.scrapQty = localWS.Cells(p, 14).Value
            If localWS.Cells(p, 18).Value = "Scrap D" Then
'                Dim scrapDGR As New clsCOR6
'                With scrapDGR
'                    .procOrdNum = localWS.Cells(p, 5).Value
'                    .oper = "20"
'                    .equipmentUsed = "FR" & right(localWS.Cells(p, 1).Value, 2)
'                    .location = localWS.Cells(p, 1).Value
'                    .byProduct = True
'                    .byProductBatch = scrapDBatch
'                    .addResNum = "2100094"
'                End With
                thisEntry.byProduct = True
                thisEntry.byProductBatch = scrapDBatch
                thisEntry.byProductFactor = localWS.Cells(p, 19).Value 'this is saving the proper percentage of by product
                thisEntry.addResNum = "2100094"
            End If
        End If
        thisEntry.availableSand = availableSand
        thisEntry.applyMatl = True
        thisEntry.saveIt
        
        'how much sand in SAP after transaction
        
        availableSand = sand.useMB52(localWS.Cells(p, 1).Value)
        sSandQty = sand.cSandAtFurnace
        Set sand = Nothing
        localWS.Cells(p, 17).Value = sSandQty

        localWS.Cells(p, 6).Value = thisEntry.FinishedItem
        localWS.Cells(p, 7).Value = thisEntry.FinishedItemDesc
        localWS.Cells(p, 8).Value = thisEntry.thisBatch
        localWS.Cells(p, 11).Value = True
        statusUpdate "Saving data updates locally"
        ActiveWorkbook.Save
        Set thisEntry = Nothing
    Next p


End Function
Function scrapDBatch() As String
scrapDBatch = right(Year(Now), 2) & monthLetter() & "NQA001"
End Function
Sub statusUpdate(msg As String)
    UserForm1.Label6.Caption = msg
    UserForm1.Label6.TextAlign = fmTextAlignCenter
End Sub

Sub bulkSandTransfer(startingRow As Long)
    Dim result As Boolean, keepGoing As Boolean
    Dim msgBoxTxt As String, furnaceName As String
    Dim p As Long
    keepGoing = True
    p = startingRow
    result = False
    If startingRow = 0 Then keepGoing = False
    setStartScreen ' Initialize SAP screen
    Do While keepGoing
        Set localWB = ThisWorkbook
        If localWB.Sheets("SandTrans").Cells(p, 6).Value = "False" Then
            Dim thisSession As GuiSession
            Set thisSession = SAPSession.CurrentSession
            With thisSession
                .FindById("wnd[0]/tbar[0]/okcd").Text = "mb52" ' Open transaction MB52
                .FindById("wnd[0]").SendVKey 0 ' Enter
                
                ' Enter the sand item number, batch number, and storage location in SAP
                Dim sandItemNumber As Long
                Dim sandBatchNumber As String
                sandItemNumber = localWB.Sheets("SandTrans").Cells(p, 1).Value
                sandBatchNumber = localWB.Sheets("SandTrans").Cells(p, 2).Value
                .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = sandItemNumber
                .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = sandBatchNumber
                .FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "q105"
                .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
                '.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = furnName
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "0400"
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT").Select
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT/ssubSCREEN_HEADER:SAPLALDB:3040/tblSAPLALDBINTERVAL_E/ctxtRSCSEL_255-ILOW_E[1,0]").Text = "nf21"
                .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT/ssubSCREEN_HEADER:SAPLALDB:3040/tblSAPLALDBINTERVAL_E/ctxtRSCSEL_255-IHIGH_E[2,0]").Text = "nf38"
                .FindById("wnd[0]/usr/radPA_FLT").Select
                .FindById("wnd[0]/usr/chkNEGATIV").Selected = False
                .FindById("wnd[0]/usr/chkXMCHB").Selected = True
                .FindById("wnd[0]/usr/chkNOZERO").Selected = True
                .FindById("wnd[0]/usr/chkNOVALUES").Selected = False
                .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
                
                .FindById("wnd[1]/tbar[0]/btn[8]").Press
                .FindById("wnd[0]").SendVKey 8
                
                If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
                    MsgBox "There is no sand in inventory with this combination. " & vbCrLf & localWB.Sheets("SandTrans").Cells(p, 1).Value & vbCrLf & localWB.Sheets("SandTrans").Cells(p, 2).Value
                    p = p + 1
                Else
                'removed here
                    .FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status")
                    Dim b As Integer, x As Integer
                    b = .FindById("wnd[1]/usr/lbl[17,10]").Text ' Extract the number of entries from the status window
                    .FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window
                    ' Build a string with sand availability information
                    For x = 3 To 2 + b
                        Dim xLoc As String
                        Dim xAmt As Variant
                        xLoc = .FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Location
                        xAmt = .FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
                    Next x
                    Dim fromLoc As String
                    ' Get the transfer amount and source location from the user
                    fromLoc = xLoc
                    Dim transferAmt As Variant
                    transferAmt = localWB.Sheets("SandTrans").Cells(p, 3).Value
                    furnaceName = localWB.Sheets("SandTrans").Cells(p, 5).Value
                        ' Perform the transfer (record in spreadsheets and update SAP if online)
                    If transferMaterial(sandItemNumber, sandBatchNumber, transferAmt, fromLoc, furnaceName, False) Then
                        localWB.Sheets("SandTrans").Cells(p, 6).Value = True
                        saveRemoteSandTransfers p
                        result = True
                    Else
                        msgBoxTxt = "You tried to transfer item: " & vbTab & sandItemNumber & vbCrLf _
                         & "Batch number: " & vbTab & vbTab & vbTab & sandBatchNumber & vbCrLf _
                         & "Quantity: " & vbTab & vbTab & vbTab & transferAmt & vbCrLf _
                         & "to " & vbTab & vbTab & vbTab & vbTab & furnName & vbCrLf _
                         & "It failed. Please transfer it manually."
                        MsgBox msgBoxTxt, vbOKOnly
                    End If 'for actual transaction
                    Set thisSession = Nothing ' Release the SAP session object
                    p = p + 1
                End If 'for having inventory
            End With
        Else
            keepGoing = False
        End If
    Loop
End Sub
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
Sub saveRemoteData(ByVal rowNum As Long, update As Boolean)
    Dim col As Integer
    Dim rw As Integer
    Dim thisWB As Workbook
    Dim remoteWB As Workbook
    Dim newRow As Integer
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
    Dim thisSandSheet As Worksheet
    Dim remoteSandSheet As Worksheet
    Dim filePath As String
    On Error GoTo ErrHandler
        
        If productionMode Then
            filePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\LampRodDrawLog.xlsx"
        Else
            filePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\LampRodDrawLog.xlsx"
        End If
                ' --- Open the external workbook ---
        If IsWorkBookOpen(filePath) Then ' Check if the workbook is already open
            'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
        Else
            Workbooks.Open filename:=filePath, ReadOnly:=False, ignorereadonlyrecommended:=True  ' Open the workbook if it's not already open
        End If
    
        
        Set thisWB = ThisWorkbook
        Set thisSheet = thisWB.Sheets("Sheet1")
        
        Set remoteWB = Workbooks("LampRodDrawLog.xlsx")
        Set remoteSheet = remoteWB.Sheets("Sheet1")
'        If Not update Then
 '           newRow = nextEmptyRow(remoteSheet, 1, 2)
  '      Else
   '         newRow = rowNum
    '    End If
        For col = 1 To 19
            remoteSheet.Cells(rowNum, col).Value = thisSheet.Cells(rowNum, col)
        Next col
        'save furnace data
        col = CInt(right(thisSheet.Cells(rowNum, 1), 2)) - 10
        For rw = 2 To 10
            remoteWB.Sheets("Sheet2").Cells(rw, col) = thisWB.Sheets("Sheet2").Cells(rw, col)
        Next rw
        
        remoteWB.Close SaveChanges:=True

    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteData of Module 1"
End Sub
Function getProcOrdNum(itemNum As String) As String
    ' This function retrieves the process order number for the given item number from SAP.

    Dim thisSession As GuiSession
    Dim x  As Integer
    Dim result As String
On Error GoTo ErrHandler
    Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
    setStartScreen ' Call a function to set the initial screen

    ' --- Start interacting with the COOISPI screen in SAP ---
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "COOISPI" ' Enter transaction code COOISPI
    thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
    ' Select checkboxes for "Order" and "Material"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = "" ' Clear order number field
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = itemNum ' Enter the provided item number
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105" ' Set plant to q105
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco" ' Set system status 1 to "teco"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "gmps" ' Set system status 2 to "gmps"
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Execute (likely F8)
    ' --- End of COOISPI screen interaction ---

    x = 0 ' Initialize row counter

    ' Check if there is data for the selection
    If SAPStatus.CurrentStatus = "There is no data for the selection" Then
        MsgBox "There are no open process orders for " & itemNum, vbOKOnly
        Exit Function
    End If

    ' Navigate to the first cell in the results and double-click to open the order
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellRow = x
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellColumn = "AUFNR"
    thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
    result = thisSession.FindById("wnd[0]/usr/txtCAUFVD-AUFNR").Text
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing
    getProcOrdNum = result
    Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getProcOrdNum function of Module1"
    
End Function


' This subroutine saves sand transfer data from the local "SandTrans" sheet
' to a backup spreadsheet stored on SharePoint.
'
' Args:
'   myRowNum: The row number in the local "SandTrans" sheet containing the data to save.
Sub saveRemoteSandTransfers(myRowNum As Long)

    ' Declare variables
    Dim col As Long ' Use Long for consistency with row numbers
    Dim thisWB As Workbook
    Dim remoteWB As Workbook
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
    Dim t As Long
    Dim filePath As String
    On Error GoTo ErrHandler
        
        If productionMode Then
            filePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\LampRodDrawLog.xlsx"
        Else
            filePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\LampRodDrawLog.xlsx"
        End If

    ' Open the remote workbook if it's not already open
    If Not IsWorkBookOpen(filePath) Then
      Workbooks.Open filename:=filePath, ignorereadonlyrecommended:=True
    End If
    
    ' Set workbook and worksheet objects
    Set thisWB = ThisWorkbook
    Set thisSheet = thisWB.Sheets("SandTrans")
    Set remoteWB = Workbooks("LampRodDrawLog.xlsx")
    Set remoteSheet = remoteWB.Sheets("SandTrans")
    t = nextEmptyRow(remoteSheet, 1, 2)
  ' Copy data from the local sheet to the remote sheet
    For col = 1 To 6
      remoteSheet.Cells(t, col).Value = thisSheet.Cells(myRowNum, col).Value
    Next col

  ' Close the remote workbook and save changes
  remoteWB.Close SaveChanges:=True

  ' Clean up object references
  Set thisWB = Nothing
  Set thisSheet = Nothing
  Set remoteWB = Nothing
  Set remoteSheet = Nothing

  Exit Sub ' Normal exit

ErrHandler:
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteSandTransfers of Module 1"
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
    
    Print #fileNum, Format(Date, "mm/dd/yyyy") & "," & Time() & "," & Environ("computername") & ",Draw Application," & Environ$("UserName") & "," & messageType & "," & message
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
Function IsWorkBookOpen(ByVal filePath As String) As Boolean

  Dim wb As Workbook

  ' Check if a file path was provided
  If Len(filePath) > 0 Then
    ' Loop through all open workbooks
    For Each wb In Workbooks
      ' Compare the full file paths, handling potential case differences
        If LCase(wb.FullName) = LCase(filePath) Then
            IsWorkBookOpen = True
            Exit Function ' Workbook found, exit the function
        End If
    Next wb
  End If

  ' If no match was found, the workbook is not open
  IsWorkBookOpen = False

End Function

'-----------------------------------------------------------------------------------
' Subroutine: Sort_2D (Recursive Quicksort helper)
' Purpose:    Recursively sorts a 2D array.
' Parameters:
'   - arr: The 2D array being sorted.
'   - left: The lower bound (start index) of the current partition.
'   - right: The upper bound (end index) of the current partition.
'   - col: The column number to sort by.
'-----------------------------------------------------------------------------------
 Function sort_2D(ByRef arr As Variant, ByVal left As Long, ByVal right As Long, ByVal col As Long) As Variant
    Dim i As Long, j As Long
    Dim pivot As Variant
    
    i = left
    j = right
    
    ' Select the pivot element (in this case, the middle element).
    pivot = arr((left + right) \ 2, col)
    
    ' Partition the array.
    Do While i <= j
        ' Find an element on the left side that is greater than the pivot.
        Do While arr(i, col) < pivot And i < right
            i = i + 1
        Loop
        
        ' Find an an element on the right side that is less than the pivot.
        Do While pivot < arr(j, col) And j > left
            j = j - 1
        Loop
        
        ' If the indices haven't crossed, swap the elements.
        If i <= j Then
            ' Swap the entire rows.
            Dim tempRow() As Variant
            ReDim tempRow(LBound(arr, 2) To UBound(arr, 2))
            
            Dim k As Long
            For k = LBound(arr, 2) To UBound(arr, 2)
                tempRow(k) = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = tempRow(k)
            Next k
            
            i = i + 1
            j = j - 1
        End If
    Loop
    
    ' Recursively sort the sub-arrays.
    If left < j Then
        sort_2D arr, left, j, col
    End If
    
    If i < right Then
        sort_2D arr, i, right, col
    End If
    sort_2D = arr
End Function
Sub printLabels(processOrder As String, itemNum As String, batchNum As String, QTY As Variant, uom As String, copies As Long, printer As String)

    Dim thisGUI As GuiSession
    Set thisGUI = SAPSession.CurrentSession
    With thisGUI
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nZPPLABEL"
        .FindById("wnd[0]").SendVKey 0
        
        .FindById("wnd[0]/usr/btn%P010002_1000").Press
        .FindById("wnd[0]/usr/ctxtP_ORDER").Text = processOrder
            
        .FindById("wnd[0]/usr/ctxtP_MATNR").Text = itemNum
        .FindById("wnd[0]/usr/ctxtP_CHARG").Text = batchNum
        .FindById("wnd[0]/usr/txtP_IGMNG").Text = QTY
        .FindById("wnd[0]/usr/ctxtP_GMEIN").Text = uom
        
        
        
        .FindById("wnd[0]/usr/cmbP_ORIGIN").Key = "USA"
        
        .FindById("wnd[0]/usr/txtP_NUM").Text = copies
        .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printer
        
        .FindById("wnd[0]/usr/ctxtP_PRINT").SetFocus
        .FindById("wnd[0]/usr/ctxtP_PRINT").CaretPosition = 10
        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        
        If SAPStatus.SAPWindowName = "Print Confirmation!" Then
            .FindById("wnd[0]").SendVKey 0
        End If
        .FindById("wnd[0]").SendVKey 0
        PauseMe 1, False
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisGUI = Nothing
End Sub

