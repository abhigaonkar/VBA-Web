Attribute VB_Name = "Module2"
Option Explicit
Function onlineProcess(ByVal rowNum As Long) As Boolean
    ' This subroutine performs the main processing for creating a new tube entry.
    ' It gathers data from the user interface, updates the spreadsheet, and runs the COR6 transaction in SAP.
    ' It also handles optional dimension entry in SAP.
    Dim x As Long ' Row number in the spreadsheet
    ' Declare variables to store tube data
    Dim pCard, oper, procOrder, tubeNote, scrapShort, scrapReason As String
    Dim ODAC, ODBD, ODLabel, ODEnd, thickness1, thickness2, thickness3, thickness4, deflection As Single
    Dim availableSand As Variant ' Array to store sand data
    Dim newStr As String
    Dim byProduct As Boolean
    
    Dim orderCheck As New clsCOOISPI
    
    
On Error GoTo ErrHandler
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")
    byProduct = False
    statusUpdate "Processing"
    localWS.Activate
    x = rowNum    ' Get the next available row number in the spreadsheet (using the furnName variable)
    pCard = localWS.Cells(x, 3) ' Assign the provided pCardNum to the pCard variable
    oper = localWS.Cells(x, 4) ' Get the operator name from a control named "operator"
    procOrder = orderCheck.getProcOrdNum(localWS.Cells(x, 6), True) ' Get the process order number for the given item number
    If orderCheck.numOrders = 0 Then
        MsgBox "There are no process orders open for item " & localWS.Cells(x, 6) & ".", vbCritical
        onlineProcess = False
        Exit Function
    Else
        localWS.Cells(x, 5).Value = procOrder
    End If
    LogEvent "Saving process order " & procOrder & " to SAP"
    Dim itemNum As Long
    Dim thisItem As New clsMaterial
    itemNum = localWS.Cells(x, 6)
    thisItem.getDescription CStr(itemNum)
    localWS.Cells(x, 37) = thisItem.LENGTH / 25.4
    
    Set thisItem = Nothing
    ' --- Initialize and populate a COR6 object ---
    Dim thisCOR6 As New clsCOR6 ' Create an instance of the clsCOR6 class
    thisCOR6.procOrdNum = procOrder ' Set the process order number
    thisCOR6.oper = "20" ' Set the operation number (hardcoded to "20")
    thisCOR6.totalTubeWeight = localWS.Cells(x, 38)
    thisCOR6.equipmentUsed = equipName ' Set the equipment used (using the equipName variable)
 '   thisCOR6.actualTubeWeightDelta = Round(calcActualTubeWeightDelta(localWS.Cells(x, 36), localWS.Cells(x, 37)), 5)
    tubeNote = left(localWS.Cells(x, 23), 40) ' Get the tube notes (truncated to 40 characters)
    ' --- Check if the tube is scrapped ---
    If Not localWS.Cells(x, 24) Then  ' If ToggleButton1 is not pressed (meaning the tube is NOT scrapped)
        thisCOR6.yield = 1 ' Set yield to 1 (full yield)
    Else ' If ToggleButton1 is pressed (meaning the tube IS scrapped)
        scrapShort = localWS.Cells(x, 25) ' Get the scrap code from ComboBox1
       ' scrapReason = ComboBox1.Column(0) ' Get the scrap reason from ComboBox1
        thisCOR6.scrapCode = scrapShort ' Set the scrap code in the COR6 object
        ' NEED TO HANDLE BYPRODUCTS
        If tubeNote <> "trash" Then
            byProduct = True
            thisCOR6.addResNum = tubeNote
        End If
        thisCOR6.scrapQty = 1
        
    End If
    ' --- End of scrap handling ---

    ' --- Get hours from user input ---
    thisCOR6.numHrs = localWS.Cells(x, 26)

    statusUpdate "Finding Sand"

    ' --- Get available sand using FIFO ---
    Dim sand As New clsFIFO
    availableSand = sand.useMB52(localWS.Cells(x, 1))
    sSandQty = sand.cSandAtFurnace
    Set sand = Nothing
    'availableSand = useMB52(furnName) ' Call the useMB52 function to get sand data
    ' --- End of getting sand data ---

    ' --- Run the COR6 transaction in SAP ---
    thisCOR6.tubeSerialNumber = localWS.Range("H" & x).Value ' Set the tube serial number (from the spreadsheet)
    thisCOR6.comment = tubeNote ' Set the comment (same as tube notes)

   ' thisCOR6.actualTubeWeightDelta = Round(calcActualTubeWeightDelta(localWS.Cells(x, 36)), 5)
    thisCOR6.applyMatl = True ' Set applyMatl to True (meaning apply material)
    thisCOR6.availableSand = availableSand ' Set the available sand data
    statusUpdate "Saving confirmation"
    If thisCOR6.saveIt Then ' Check if the COR6 transaction was saved successfully
        localWS.Range("V" & x).Value = True ' Mark the COR6 transaction as successful in the spreadsheet
    Else
        localWS.Range("V" & x).Value = True ' Mark the COR6 transaction as failed in the spreadsheet
    End If

    ' --- End of running COR6 ---

    localWS.Range("G" & x).Value = thisCOR6.FinishedItemDesc  ' Store the value from Label70 in the spreadsheet
    ThisWorkbook.Sheets("Sheet2").Cells(24, 11).Value = localWS.Range("G" & x).Value

    Set thisCOR6 = Nothing ' Release the COR6 object

    ' --- Optional dimension entry in SAP ---
    ' need to move dimesional calculations to here so that it's stored when SAP is down?
    
    statusUpdate "Saving dimensions in SAP"

    If Not localWS.Cells(x, 24) Then   ' If this isn't scrap then dimensions should be entered in SAP
        Dim theseDims As New clsDimEntry
        theseDims.saveDimensions x  ' Call a method to save the dimensions in SAP
        Set theseDims = Nothing ' Release the clsDimEntry object
    End If
    If byProduct And (tubeNote = "2307812" Or tubeNote = "2307813" Or tubeNote = "2307814" Or tubeNote = "2307815" Or tubeNote = "2307816") Then 'to handle HU for Webb
            statusUpdate "Saving Handling Unit for Webb"
        ' see how much is in 0300
        Dim thisInvCheck As New clsMB52
        Dim y As Integer
        thisInvCheck.getInventory tubeNote, "0300"
        If thisInvCheck.isSomeThere Then
            If thisInvCheck.totalAtLocation > 200 Then
                If MsgBox("The Webb hamper now has " & thisInvCheck.totalAtLocation & " KGs. Would you like to close it?", vbYesNo) = vbYes Then
                    'take care of HU
                     'do migo for delivery
                    migo311Array thisInvCheck.inventoryArray()
                    
                     'do vl32n for receiving to H300
                     vl32N thisInvCheck.inventoryArray()
                     'save HUNum in column 38
                     
                     'when hamper full do HU4GR to complete it
'                     Dim thisHU As New clsHU
'                     Dim p As Integer
'                     Dim newArray() As Variant
'                     ReDim newArray(UBound(thisInvCheck.inventoryArray))
'                     For p = 0 To UBound(thisInvCheck.inventoryArray)
'                        newArray(p) = thisInvCheck.inventoryArray(p, 1)
'                     Next p
'                     thisHU.itemNum = tubeNote
'                     thisHU.batchArray = newArray
'                     thisHU.humo4Gr
'                     localWS.Cells(x, 38).Value = thisHU.HUNum
'                     Set thisHU = Nothing
                End If

            End If
        End If
    statusUpdate ""

    End If
    ' --- End of dimension entry ---
    onlineProcess = True
Exit Function
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in onlineProcess Function of Module1"
End Function
' This subroutine updates a visual sand level indicator on a user form.
' It reads the amount of sand from a cell in "Sheet2" and displays
' a corresponding number of illuminated "lights" on the form.
Sub setSandMeter()

    ' Declare variables
    Dim amtSand As Variant
    Dim x As Integer, lightsLit As Integer
    If onlineMode Then
        statusUpdate "Setting Sand Meter"

        Dim thisSandCheck As New clsFIFO
        thisSandCheck.useMB52 furnName
        Set thisSandCheck = Nothing
        ' Get the amount of sand from cell AA2 in "Sheet2"
        amtSand = ThisWorkbook.Sheets("Sheet2").Range("AA2").Value
        
        ' Determine the number of lights to illuminate based on the amount of sand
        Select Case amtSand
            Case Is < 140: lightsLit = 1
            Case Is < 200: lightsLit = 2
            Case Is < 275: lightsLit = 3
            Case Is < 350: lightsLit = 4
            Case Is < 425: lightsLit = 5
            Case Is < 500: lightsLit = 6
            Case Is < 575: lightsLit = 7
            Case Else:      lightsLit = 8
        End Select
        
        ' Turn on the appropriate number of lights
        For x = 1 To 8 ' Loop through all lights
            If x <= lightsLit Then
                UserForm1.Controls("Sand" & x).Visible = True    ' Make the light visible
                UserForm1.Controls("Sand" & x).ControlTipText = Round(amtSand, 3) & " KGS" ' Set tooltip
            Else
                UserForm1.Controls("Sand" & x).Visible = False   ' Hide the light
            End If
        Next x
    End If
    statusUpdate ""
End Sub

Function calcActualTubeWeightDelta(actTubeLength As Variant, docTubeLength As Variant) As Variant
    If actTubeLength > docTubeLength Then
        calcActualTubeWeightDelta = 1 + (Abs(actTubeLength - docTubeLength) / docTubeLength)
    ElseIf actTubeLength < docTubeLength Then
        calcActualTubeWeightDelta = 1 - (Abs(actTubeLength - docTubeLength) / docTubeLength)
    Else
        calcActualTubeWeightDelta = 1
    End If
End Function

Public Sub changeUserDefaults()
    ' sets the number format, date, and time formats
    If onlineMode Then
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
    End If
End Sub
Sub saveRemoteData(rowNum As Long, update As Boolean)
    Dim col As Integer
    Dim thisWB As Workbook
    Dim remoteWB As Workbook
    Dim remoteRow As Long
    Dim thisSheet As Worksheet
    Dim remoteSheet As Worksheet
On Error GoTo ErrHandler
        statusUpdate "Saving backup"
                ' --- Open the external workbook ---
        If IsWorkBookOpen(filePath) Then ' Check if the workbook is already open
            'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
        Else
            Workbooks.Open filename:=filePath, ignorereadonlyrecommended:=True   ' Open the workbook if it's not already open
        End If
    
        
        Set thisWB = ThisWorkbook
        Set thisSheet = thisWB.Sheets("Sheet1")
        
        Set remoteWB = Workbooks("SerializedTubeLog.xlsx")
        Set remoteSheet = remoteWB.Sheets("Sheet1")
        If Not update Then
            remoteRow = nextEmptyRow(remoteSheet, "H", 2)
        Else
            remoteRow = FindLastOccurrence(thisSheet.Cells(rowNum, 8), remoteSheet, Range("H:H")).Row
        End If
        For col = 1 To 37
            remoteSheet.Cells(remoteRow, col).Value = thisSheet.Cells(rowNum, col)
        Next col
        Set remoteSheet = remoteWB.Sheets("SandQTY")
        If sSandQty <> 0 Then remoteSheet.Cells(2, sandQtyCol).Value = sSandQty      'update remote spreadsheet with sand available at the furnace
        remoteSheet.Cells(3, sandQtyCol).Value = Now
 '       Set thisWB = Nothing
 '       Set thisSheet = Nothing
        
        remoteWB.Close SaveChanges:=True
  '      Set remoteWB = Nothing
  '      Set remoteSheet = Nothing
    statusUpdate ""

    Exit Sub
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteData of Module 2"
End Sub

Public Function FindLastOccurrence(searchString As String, ws As Worksheet, searchColumn As Range) As Range
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
    ws.Activate
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


' This subroutine saves sand transfer data from the local "SandTrans" sheet
' to a backup spreadsheet stored on SharePoint.
'
' Args:
'   myRowNum: The row number in the local "SandTrans" sheet containing the data to save.
Sub saveRemoteSandTransfers(myRowNum As Long)

  ' Declare variables
  Dim filePath As String
  Dim col As Long ' Use Long for consistency with row numbers
  Dim thisWB As Workbook
  Dim remoteWB As Workbook
  Dim thisSheet As Worksheet
  Dim remoteSheet As Worksheet

  On Error GoTo ErrHandler ' Error handling

  ' Open the remote workbook if it's not already open
  If Not IsWorkBookOpen(filePath) Then
    Workbooks.Open filename:=filePath, ignorereadonlyrecommended:=True
  End If

  ' Set workbook and worksheet objects
  Set thisWB = ThisWorkbook
  Set thisSheet = thisWB.Sheets("SandTrans")
  Set remoteWB = Workbooks("FurnaceDeckRunData.xlsx")
  Set remoteSheet = remoteWB.Sheets("SandTrans")

  ' Copy data from the local sheet to the remote sheet
  For col = 1 To 6
    remoteSheet.Cells(nextEmptyRow(remoteSheet, "A", 2), col).Value = _
      thisSheet.Cells(myRowNum, col).Value
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
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveRemoteSandTransfers of Module 2"
End Sub

Sub SaveAndCloseWorkbook()
    LogEvent "Closing Workbook", "Application"
    ' Save the active workbook
    ActiveWorkbook.Save

    ' Close the active workbook
    ActiveWorkbook.Close

End Sub
' This function finds the next empty row in a specified column of a worksheet.
'
' Args:
'   shet: The worksheet to search in.
'   col: The column letter (e.g., "A", "B", "C") to search in.
'   rowStart: (Optional) The row number to start the search from.
'             Defaults to 6 if not provided.
'
' Returns:
'   The row number of the first empty cell in the specified column.
Function nextEmptyRow(shet As Worksheet, col As String, Optional rowStart As Long = 2) As Long

  On Error GoTo ErrHandler ' Error handling

  ' Initialize the starting row
  Dim x As Long
  x = rowStart

  ' Loop through the cells in the specified column
  Do While Not IsEmpty(shet.Cells(x, col).Value)
    x = x + 1
  Loop

  ' Return the row number of the first empty cell
  nextEmptyRow = x

  Exit Function ' Normal exit

ErrHandler:
  ' Call a custom error handling subroutine (StdErrorHandler)
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextEmptyRow Function of Module1"
End Function
Function getSandLoc(ByVal itemNum As Long, ByVal batchNum As String) As String
    If onlineMode Then
        Dim thisSession As GuiSession
        Dim b As Integer
        Dim x As Integer
        Dim fromLoc As String
        Dim sandStatement As String
        Set thisSession = SAPSession.CurrentSession
        setStartScreen ' Initialize SAP screen
        
        With thisSession
            .FindById("wnd[0]/tbar[0]/okcd").Text = "mb52" ' Open transaction MB52
            .FindById("wnd[0]").SendVKey 0 ' Enter
            
            ' Enter the sand item number, batch number, and storage location in SAP
            .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = itemNum
            .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = batchNum
            .FindById("wnd[0]/usr/ctxtP_VARI").Text = "/parkinsonj"
            
            .FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = ""
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = furnName
            .FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "0400"
            .FindById("wnd[1]/tbar[0]/btn[8]").Press
              
            .FindById("wnd[0]").SendVKey 8
            If SAPStatus.CurrentStatus = "No stock exists for specified data" Then
                MsgBox "There is no sand in inventory with this item number/batch combination.", vbOKOnly
                fromLoc = ""
            Else
                .FindById("wnd[0]/mbar/menu[3]/menu[8]").Select ' Select a menu item (e.g., "System" -> "Status")
                b = .FindById("wnd[1]/usr/lbl[17,8]").Text ' Extract the number of entries from the status window
                .FindById("wnd[1]/tbar[0]/btn[0]").Press ' Close the status window
                ' Build a string with sand availability information
                sandStatement = "Sand Availability:"
                For x = 3 To 2 + b
                     Dim xLoc As String
                     Dim xAMT As Variant
                    sandStatement = sandStatement & vbCrLf
                    xLoc = .FindById("wnd[0]/usr/lbl[48," & x & "]").Text ' Location
                    xAMT = .FindById("wnd[0]/usr/lbl[68," & x & "]").Text ' Quantity
                    sandStatement = sandStatement & "Location: " & xLoc
                    sandStatement = sandStatement & vbTab & "Amount: " & xAMT
                Next x
                
                ' Get the transfer amount and source location from the user
                If b = 1 Then
                    fromLoc = xLoc
                    sandStatement = sandStatement & vbCrLf & vbCrLf
                Else
                    fromLoc = InputBox(sandStatement & vbCrLf & vbCrLf & "Which location would you like to transfer from?", "Sand Location Query")
                End If
           End If
        End With ' thisSession
        Set thisSession = Nothing ' Release the SAP session object
    End If
    getSandLoc = fromLoc
End Function

Function transferMatFromSS(startingRow As Long) As Boolean
' Declare variables
Dim x As Long
Dim itemNum As Long
Dim batch As String
Dim fromLoc As String
Dim toLoc As String
Dim onl As Boolean
Dim thisMove As GuiSession
Dim qty As Variant
Dim result As Boolean

On Error GoTo ErrHandler ' Error handling
    If onlineMode Then
        result = False
      ' Record the transfer in the local "SandTrans" sheet
        x = startingRow
        With ThisWorkbook.Sheets("SandTrans")
            itemNum = .Cells(x, 1).Value
            batch = .Cells(x, 2).Value
            qty = .Cells(x, 3).Value
            fromLoc = .Cells(x, 4).Value
            toLoc = .Cells(x, 5).Value
            onl = .Cells(x, 6).Value
        End With
      
      
        If fromLoc = "" Then
            fromLoc = getSandLoc(itemNum, batch)
            If fromLoc = "" Then Exit Function
        End If
        
        
      ' Save the transfer data to the backup spreadsheet
      'saveRemoteSandTransfers x
    
      ' If SAP is online, perform the transfer in SAP
        setStartScreen ' Initialize SAP screen (assuming this is a custom function)
        Set thisMove = SAPSession.CurrentSession ' Get the current SAP session
    
        With thisMove
            ' Navigate to the relevant transaction in SAP (ZMMS2S)
            .FindById("wnd[0]/tbar[0]/okcd").Text = "ZMMS2S"
            .FindById("wnd[0]").SendVKey 0 ' Enter
            
            ' Populate the required fields in SAP
            .FindById("wnd[0]/usr/txtMCHB-LGORT").Text = fromLoc
            .FindById("wnd[0]/usr/txtV_BATCH").Text = batch
            .FindById("wnd[0]/usr/ctxtMARA-MATNR").Text = itemNum
            .FindById("wnd[0]").SendVKey 5 ' Enter
            
            .FindById("wnd[0]/usr/txtLAGP-RKAPV").Text = CDec(qty) ' Quantity as decimal
            .FindById("wnd[0]/usr/txtT001L-LGORT").Text = toLoc
            .FindById("wnd[0]").SendVKey 5 ' Enter
            
            ' Check for success message in SAP
            If .FindById("wnd[0]/usr/txtV_MESSAGE3").Text = "Created." Then
              result = True
              ThisWorkbook.Sheets("SandTrans").Cells(x, 6).Value = "True"
            End If
            
            ' Close the transaction in SAP
            .FindById("wnd[0]").SendVKey 8 ' F8 (possibly to save)
            .FindById("wnd[0]/tbar[0]/btn[3]").Press ' Back button
        End With
    
        Set thisMove = Nothing ' Release the SAP session object
    End If
    transferMatFromSS = result
  Exit Function ' Normal exit

ErrHandler:
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in transferMaterialFromSS Function of Module1"

End Function

Public Sub ketchup()
    'Make sure sand is transferred first
    'Ask what row to begin with
    'Enter SAP transactions
    'Switch to Online mode
    If MsgBox("Have you made all sand transfers?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Dim keepGoing As Boolean
    Dim rowN As Integer
    keepGoing = True
    rowN = InputBox("What row would you like to begin with?")
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")
    localWS.Activate
    Do While keepGoing
        If localWS.Cells(rowN, 1).Value <> "" Then
            If localWS.Cells(rowN, 22).Value = False Then
'                furnName = localWS.Cells(rowN, 1).Value
                onlineProcess rowN
                
            End If
            rowN = rowN + 1
        Else
            If MsgBox("SAP entry is complete. Would you like to go back to online mode?", vbYesNo) = vbYes Then
                ThisWorkbook.Sheets("Sheet2").Range("AA1").Value = "Online"
                onlineMode = True
                UserForm1.Show
            End If
            keepGoing = False
            Exit Sub
        End If
    Loop
    
        
End Sub
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

Public Sub massSandTransfer()
    Dim keepGoing As Boolean
    Dim x As Long
    Dim thisWB As Workbook
    Dim thisSheet As Worksheet
    x = 2
    Set thisWB = ThisWorkbook
    Set thisSheet = thisWB.Sheets("SandTrans")

    keepGoing = True
    With thisSheet
    Do While keepGoing
        If .Cells(x, 6).Value = "" Then
            keepGoing = False
        ElseIf .Cells(x, 6).Value = False Then
            If transferMatFromSS(x) Then
                .Cells(x, "F").Value = True
            End If
        End If
        x = x + 1
    Loop
    End With
    MsgBox "Transfers completed."
    Set thisSheet = Nothing
    Set thisWB = Nothing
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
    
    Print #fileNum, Format(Date, "mm/dd/yyyy") & "," & Time() & "," & Environ("computername") & ",Serialized Draw Application," & Environ$("UserName") & "," & messageType & "," & message
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Sub printLabel(ByVal procOrdNum As String, ByVal itemNum As Long, ByVal batch As String, numTickets As Integer, qtyEach As Variant, ByVal uom As String)
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    statusUpdate "Printing labels"
    
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nzpplabel"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/btn%P010002_1000").Press
        .FindById("wnd[0]/usr/ctxtP_ORDER").Text = procOrdNum
        .FindById("wnd[0]/usr/ctxtP_MATNR").Text = itemNum
        .FindById("wnd[0]/usr/ctxtP_CHARG").Text = batch
        .FindById("wnd[0]/usr/txtP_IGMNG").Text = qtyEach
        .FindById("wnd[0]/usr/ctxtP_GMEIN").Text = uom
        .FindById("wnd[0]/usr/cmbP_ORIGIN").Key = "USA"
        .FindById("wnd[0]/usr/txtP_NUM").Text = numTickets
        .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName
        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        If SAPStatus.IsThereAModalWindow Then .FindById("wnd[1]/tbar[0]/btn[0]").Press
    '    .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisSession = Nothing
    statusUpdate ""

End Sub
Sub printHULabel(ByVal HUNum As String, ByVal numLabels As Integer)
    Dim thisLabel As GuiSession
    Set thisLabel = SAPSession.CurrentSession
    statusUpdate "Printing Handling Unit label"
    With thisLabel
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nzpplabel"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/btn%P150004_1000").Press
        .FindById("wnd[0]/usr/ctxtS_HUNO-LOW").Text = HUNum
        .FindById("wnd[0]/usr/cmbP_ORIGIN").Key = "USA"
        .FindById("wnd[0]/usr/txtP_NUM").Text = numLabels
        .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName

        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisLabel = Nothing
    statusUpdate ""
End Sub

Function migo311Array(ByVal arrayToTransfer As Variant) As String
'returns shipment number

    Dim item As String
    Dim uom As String
    Dim qty As Variant
    Dim location As String
    Dim batch As String
    Dim fromLoc, toLoc As String
    Dim thisGUI As GuiSession
    Set thisGUI = SAPSession.CurrentSession
    Dim x As Long
    
    With thisGUI
    .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmigo"
    .FindById("wnd[0]").SendVKey 0
    .FindById("wnd[0]").SendVKey 0
    .FindById("wnd[0]/tbar[0]/btn[3]").Press
    
    For x = 0 To UBound(arrayToTransfer)
        item = arrayToTransfer(x, 0)
        batch = arrayToTransfer(x, 1)
        qty = arrayToTransfer(x, 3)
        fromLoc = arrayToTransfer(x, 2)
        uom = "KG"
        toLoc = "H300"

        .FindById("wnd[0]/tbar[0]/okcd").Text = "migo"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").Text = item
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").Text = "q105"
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").Text = fromLoc

        .FindById("wnd[0]").SendVKey 0

        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").Text = batch
        .FindById("wnd[0]").SendVKey 0

        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMLGOBE").Text = toLoc
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").Text = qty
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFME").Text = uom

        .FindById("wnd[0]").SendVKey 0
            ' go to next item
        If x < UBound(arrayToTransfer) Then
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/btnOK_COPY").Press
            .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").Selected = True
        End If

    Next x
    .FindById("wnd[0]/tbar[1]/btn[23]").Press
    migo311Array = Mid(SAPStatus.CurrentStatus, 10, 10)
    .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisGUI = Nothing
        
End Function

Function migo311(ByVal itemNum As String, ByVal batchNum As String, ByVal qty As Variant, ByVal fromLoc As String) As String
Dim keepGoing As Boolean

    Dim toLoc As String
    Dim thisGUI As GuiSession
    Set thisGUI = SAPSession.CurrentSession
    With thisGUI
    
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nmigo"
        toLoc = "H300"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").Text = itemNum
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").Text = "q105"
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").Text = fromLoc
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").Text = batchNum
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMLGOBE").Text = toLoc
        .FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").Text = qty
        .FindById("wnd[0]").SendVKey 0
            ' go to next item
        
        .FindById("wnd[0]/tbar[1]/btn[23]").Press
        migo311 = Mid(SAPStatus.CurrentStatus, 10, 10)
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        
    End With
        
End Function
Sub vl32N(ByVal weightArray As Variant)
    Dim thisGUI As GuiSession
    Dim x As Integer
On Error GoTo ErrHandler ' Error handling

    Set thisGUI = SAPSession.CurrentSession
    With thisGUI
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nvl32n"
        .FindById("wnd[0]").SendVKey 0

        .FindById("wnd[0]").SendVKey 0
'        .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB").GetAbsoluteRow(0).Selected = True
        .FindById("wnd[0]/tbar[1]/btn[18]").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").Text = "2309602"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT014").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT011").Press
        .FindById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/btn%#AUTOTEXT001").Press
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        .FindById("wnd[0]").SendVKey 0
'        .FindById("wnd[0]").SendVKey 0
'        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        For x = 0 To UBound(weightArray)
        '.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPSD-PIKMG[17,0]").Text = CStr(weight)
            .FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPSD-PIKMG[17," & x & "]").Text = CStr(weightArray(x, 3))
        Next x
        .FindById("wnd[0]/tbar[1]/btn[20]").Press
    End With
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub
Sub statusUpdate(msg As String)
UserForm1.LabelStatus.Caption = msg
End Sub
