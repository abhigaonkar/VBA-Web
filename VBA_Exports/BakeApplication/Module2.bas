Attribute VB_Name = "Module2"
Option Explicit
Sub saveToSAP(ByVal lineNum As Long) 'serializedFeedstock is for knowing how to handle the hours
'Process:
'   COR6 for confirmation
'   COP1 to create HU for each item
'   COWBHUWE for GR
'   Pack HUs using HUMO or HU02

    Dim PONum As String
    
    Dim goodsIssued() As Variant
    Dim numGoodsToIssue As Integer
    Dim priArray As Variant
    Dim isSerialized As Boolean
    ReDim priArray(3, 1)
    Set localSheet = ThisWorkbook.Sheets("Sheet1")
    If localSheet.Cells(lineNum, 5).Value = "" Then
        localSheet.Cells(lineNum, 5).Value = getProcOrdNums(localSheet.Cells(lineNum, 6).Value, localSheet.Cells(lineNum, 15).Value, True)(0)
    End If
    PONum = localSheet.Cells(lineNum, 5).Value
    Dim thisCOR6 As New clsNewCOR6
    If localSheet.Cells(lineNum, 26).Value = "Serialized" Then
        isSerialized = True
        thisCOR6.numHrs = localSheet.Cells(lineNum, 11).Value
        thisCOR6.tubeSerialNumber = localSheet.Cells(lineNum, 8).Value
        thisCOR6.serialized = True
    Else
        isSerialized = False
        thisCOR6.numHrs = localSheet.Cells(lineNum, 9).Value
        thisCOR6.serialized = False
    End If
    thisCOR6.procOrdNum = PONum
    thisCOR6.oper = "20"
    thisCOR6.yield = localSheet.Cells(lineNum, 14).Value
    thisCOR6.equipmentUsed = localSheet.Cells(lineNum, 1).Value
    thisCOR6.comment = localSheet.Cells(lineNum, 16).Value
    priArray(0, 0) = localSheet.Cells(lineNum, 27).Value
    priArray(0, 1) = localSheet.Cells(lineNum, 28).Value
    priArray(1, 0) = localSheet.Cells(lineNum, 29).Value
    priArray(1, 1) = localSheet.Cells(lineNum, 30).Value
    priArray(2, 0) = localSheet.Cells(lineNum, 31).Value
    priArray(2, 1) = localSheet.Cells(lineNum, 32).Value
    thisCOR6.printArray = priArray
    numGoodsToIssue = 1
    If localSheet.Cells(lineNum, 17).Value = 0 Then
        'no hamper
    Else
        numGoodsToIssue = numGoodsToIssue + 2
        If localSheet.Cells(lineNum, 35).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 1
    End If
    
'    If localSheet.Cells(lineNum, 21).Value <> "" Then numGoodsToIssue = numGoodsToIssue + 1
    thisCOR6.applyNumMatl = numGoodsToIssue
    If localSheet.Cells(lineNum, 12).Value = "" Then
        Dim fsItem() As Variant
        ReDim fsItem(2)
        fsItem = findItemDetailsFromBatch(localSheet.Cells(lineNum, 8).Value)
        localSheet.Cells(lineNum, 12).Value = fsItem(0)
        localSheet.Cells(lineNum, 13).Value = fsItem(1)
        localSheet.Cells(lineNum, 25).Value = fsItem(2)
    End If
    ReDim goodsIssued(numGoodsToIssue, 4)
   
    goodsIssued(0, 0) = localSheet.Cells(lineNum, 12).Value
    goodsIssued(0, 1) = localSheet.Cells(lineNum, 15).Value
    goodsIssued(0, 2) = localSheet.Cells(lineNum, 25).Value
    goodsIssued(0, 3) = localSheet.Cells(lineNum, 8).Value
    goodsIssued(0, 4) = 261
    
    If localSheet.Cells(lineNum, 17).Value > 0 Then
        goodsIssued(2, 0) = localSheet.Cells(lineNum, 17).Value
        goodsIssued(2, 1) = localSheet.Cells(lineNum, 18).Value
        goodsIssued(2, 2) = localSheet.Cells(lineNum, 19).Value
        goodsIssued(2, 3) = "" 'hamper batch
        goodsIssued(2, 4) = localSheet.Cells(lineNum, 20).Value 'hamper movement type
        If localSheet.Cells(lineNum, 21).Value <> "" Then
            goodsIssued(1, 0) = localSheet.Cells(lineNum, 21).Value 'strapping item num
            goodsIssued(1, 1) = localSheet.Cells(lineNum, 22).Value 'strapping qty
            goodsIssued(1, 2) = localSheet.Cells(lineNum, 23).Value 'strapping location
            goodsIssued(1, 3) = "" 'strapping batch
            goodsIssued(1, 4) = localSheet.Cells(lineNum, 24).Value 'strapping movement type
            If localSheet.Cells(lineNum, 35).Value <> "" Then
                goodsIssued(3, 0) = localSheet.Cells(lineNum, 35).Value
                goodsIssued(3, 1) = localSheet.Cells(lineNum, 36).Value
                goodsIssued(3, 2) = localSheet.Cells(lineNum, 37).Value
                goodsIssued(3, 3) = "" 'hamper batch
                goodsIssued(3, 4) = localSheet.Cells(lineNum, 38).Value 'hamper movement type
            End If
        End If
    End If
    thisCOR6.matlArray = goodsIssued
    thisCOR6.applyMatl = True
    localSheet.Cells(lineNum, 10).Value = thisCOR6.saveIt
    localSheet.Cells(lineNum, 7).Value = thisCOR6.FinishedItemDesc
    Set thisCOR6 = Nothing
    
    If isSerialized Then
        Dim thisHU As New clsHU
        thisHU.procOrdNum = PONum
        thisHU.batchNum = localSheet.Cells(lineNum, 8).Value
        thisHU.cop1
        thisHU.cowbhuwe
        Set thisHU = Nothing
    End If
            
End Sub
Sub doHumo(ByVal startRow As Long, t As Long)
    Dim r As Integer 'array position
        Set localSheet = ThisWorkbook.Sheets("Sheet1")

    Dim hamp As Integer 'what hamper are we working in
    For hamp = 1 To 2
        Dim p As Long 'spreadsheet row we're in
        Dim hampCheck As Boolean 'does this hamper have anything in it
        hampCheck = False
        Dim thisArray() As Variant
        ReDim thisArray(t - startRow)
        r = 0
        For p = startRow To t
            If localSheet.Cells(p, 33).Value = hamp Then 'does the saved hamper number match the hamper we're working in
                thisArray(r) = localSheet.Cells(p, 8).Value
                hampCheck = True 'this hamper has something in it
                r = r + 1
            End If
            
        Next p
        If hampCheck Then
            Dim thisHU As New clsHU
            thisHU.itemNum = localSheet.Cells(startRow, 6).Value
            thisHU.batchArray = removeEmptyElementsFromArray(thisArray)
            thisHU.humo4Gr
            Set thisHU = Nothing
        End If
    Next hamp
    
End Sub
Sub printLabel(ByVal procOrdNum As String, ByVal itemNum As Long, ByVal batch As String, numTickets As Integer, qtyEach As Variant, ByVal UOM As String)
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    
    
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzpplabel"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/btn%P010002_1000").Press
    thisSession.FindById("wnd[0]/usr/ctxtP_ORDER").Text = procOrdNum
    thisSession.FindById("wnd[0]/usr/ctxtP_MATNR").Text = itemNum
    thisSession.FindById("wnd[0]/usr/ctxtP_CHARG").Text = batch
    thisSession.FindById("wnd[0]/usr/txtP_IGMNG").Text = qtyEach
    thisSession.FindById("wnd[0]/usr/ctxtP_GMEIN").Text = UOM
    thisSession.FindById("wnd[0]/usr/cmbP_ORIGIN").key = "USA"
    thisSession.FindById("wnd[0]/usr/txtP_NUM").Text = numTickets
    thisSession.FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    If SAPStatus.IsThereAModalWindow Then thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
'    thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing

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
Sub CreateArraysByPalletHamper()
    
    ' Define your worksheet and data range.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim dataRange As Range
    Set dataRange = ws.Range("C62:AL" & lastRow) ' Adjust the columns as needed
    dataRange.Select
    
    ' Dictionary to store unique combinations and their associated data.
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cell As Range
    Dim key As String
    Dim dataRow As Variant
    
    ' Loop through each row to identify unique combinations and store data.
    For Each cell In dataRange.Columns(1).Cells
        
        ' Create a unique key from the pallet and hamper numbers.
        key = CStr(cell.Value) & "-" & CStr(cell.Offset(0, 30).Value)
        
        ' Store the entire row's data in the dictionary.
        dataRow = ws.Range(cell.Row & ":" & cell.Row).Value
        
        If Not dict.Exists(key) Then
            ' Create a new entry if the key doesn't exist.
            ' The item is a 2D array, which we'll later redimension.
            dict.Add key, Array(dataRow)
        Else
            ' If the key exists, add the new data to the existing array.
            Dim tempArray As Variant
            tempArray = dict.Item(key)
            
            Dim newArraySize As Long
            newArraySize = UBound(tempArray) + 1
            
            ReDim Preserve tempArray(newArraySize)
            tempArray(newArraySize) = dataRow
            dict.Item(key) = tempArray
        End If
    Next cell

    ' Now you can access the created arrays.
    ' This part of the code demonstrates how to loop through the dictionary
    ' and see the contents of each array.
    Dim arrayKey As Variant
    Dim currentArray As Variant
    Dim i As Long
    
    For Each arrayKey In dict.Keys
        currentArray = dict.Item(arrayKey)
        Debug.Print "--- Data for Combination: " & arrayKey & " ---"
        
        ' Loop through the rows in the current array
        For i = LBound(currentArray) To UBound(currentArray)
            Dim j As Long
            Dim rowData As Variant
            rowData = currentArray(i)
            
            ' Loop through the columns in the current row
            For j = LBound(rowData, 2) To UBound(rowData, 2)
                Debug.Print rowData(1, j),
            Next j
            Debug.Print
        Next i
        Debug.Print
    Next arrayKey

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

Function getProcOrdNums(ByVal itemNum As String, ByVal numTubes As Integer, ByVal unused As Boolean) As Variant
    Dim thisSession As GuiSession
    Dim x As Integer
    Dim numOpenProcessOrders As Integer
    Dim result() As Variant
    ReDim result(numTubes - 1)
    
    Set thisSession = SAPSession.CurrentSession
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncooispi"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = itemNum
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").Text = "SAPCEM1"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco"
    If unused Then thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "gmps"
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
'    thisSession.FindById("wnd[0]/mbar/menu[3]/menu[8]").Select
    numOpenProcessOrders = thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
'    thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    If numOpenProcessOrders < numTubes Then
        MsgBox "There are not enough open process orders to complete this.", vbCritical
        For x = 0 To numTubes - 1
            result(x) = "0"
        Next x
    Else
        For x = 0 To numTubes - 1
            thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellRow = x
            thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellColumn = "AUFNR"
            thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
            result(x) = thisSession.FindById("wnd[0]/usr/txtCAUFVD-AUFNR").Text
            thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
        Next x
    End If
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing
    getProcOrdNums = result
End Function
Function getCtnProcOrdNums(ByVal itemNum As String) As Variant
    Dim thisSession As GuiSession
    Dim x As Integer
    Dim numOpenProcessOrders As Integer
    Dim result() As Variant
    ReDim result(1)
    
    Set thisSession = SAPSession.CurrentSession
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncooispi"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = ""
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = itemNum
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").Text = "q105"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").Text = "SAPCEM1"
    thisSession.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco"
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
'    thisSession.FindById("wnd[0]/mbar/menu[3]/menu[8]").Select
    numOpenProcessOrders = thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
'    thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    If numOpenProcessOrders < 1 Then
        MsgBox "There are no open process orders to complete building this hamper. This transaction will continue.", vbCritical
    Else
        thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellRow = 0
        thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").CurrentCellColumn = "AUFNR"
        thisSession.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").DoubleClickCurrentCell
        result(0) = thisSession.FindById("wnd[0]/usr/txtCAUFVD-AUFNR").Text
        thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    End If
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    thisSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
    Set thisSession = Nothing
    getCtnProcOrdNums = result
End Function
Function createHUBatch(oldBatch1 As String, oldBatch2 As String) As String
    'The new batch number would be formatted as 5MT001-004, where 5 is the year, M is the month, T is the furnace, and 001-004 are the last three digits of the tubes in the hamper. So, these four tubes:
    '25MNT0001
    '25MNT0002
    '25MNT0003
    '25MNT0004
    'would be batch 5MT001-004
    Dim result As String
    Dim yearStr As String
    Dim monthStr As String
    Dim furnStr As String
    Dim startingNum As String
    Dim endingNum As String
    yearStr = right(left(oldBatch1, 2), 1)
    monthStr = right(left(oldBatch1, 3), 1)
    furnStr = right(left(oldBatch1, 5), 1)
    startingNum = right(oldBatch1, 3)
    endingNum = right(oldBatch2, 3)
    result = yearStr & monthStr & furnStr & startingNum & "-" & endingNum
    createHUBatch = result
End Function
Function reverseHUBatch(batchNum As String) As Variant
    Dim batches() As Variant
    Dim firstNum As Integer
    Dim lastNum As Integer
    Dim t As Integer
    Dim x As Integer
    x = 0
    firstNum = CInt(left(right(batchNum, 7), 3))
    lastNum = CInt(right(batchNum, 3))
    ReDim batches(lastNum - firstNum)
    For t = firstNum To lastNum
        batches(x) = "2" & left(batchNum, 2) & "N" & right(left(batchNum, 3), 1) & addZeros(t)
        x = x + 1
    Next t
    reverseHUBatch = batches()
End Function
Sub do309(ByVal itemNum As Long, ByVal oldBatch As String, ByVal newBatch As String)
    setStartScreen
    Dim thisSession As GuiSession
    Set thisSession = SAPSession.CurrentSession
    thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "migo"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION").key = "A08"
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART").Text = "309"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").Text = itemNum
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").Text = "q105"
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").Text = "0035"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").Text = oldBatch
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMMAKTX").Text = itemNum
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").Text = newBatch
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").Text = "1"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").Select
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").Text = "309"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT").Select
    thisSession.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1001/ctxtCOBL-KOSTL").Text = "nw4449"
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/tbar[1]/btn[23]").Press
    Set thisSession = Nothing
End Sub
Function findItemDetailsFromBatch(ByVal batch As String) As Variant
    Dim thisSession As GuiSession
    Dim result() As Variant
    ReDim result(2)
    Set thisSession = SAPSession.CurrentSession
    With thisSession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "mb51"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = batch
        .FindById("wnd[0]/usr/ctxtBWART-LOW").Text = "101"
        .FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
        .FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = ""
        .FindById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
        .FindById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
        .FindById("wnd[0]/usr/radRFLAT_L").Select
        .FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/S4P_DEFAULT"
     '   xxxx

        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        If .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount > 0 Then
        result(0) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MATNR")
        result(1) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "MAKTX")
        result(2) = .FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "LGORT")
        End If
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press

    End With
    Set thisSession = Nothing
    findItemDetailsFromBatch = result
End Function

' This function records sand transferred to furnaces.
' It logs the transfer locally, to a backup spreadsheet, and if SAP is online,
' it also performs the transfer in SAP.
'
' Args:
'   itemNum: The item number of the sand.
'   batch: The batch number of the sand.
'   qty: The quantity of sand transferred.
'   fromLoc: The location the sand is transferred from.
'   toLoc: The location the sand is transferred to.
'
' Returns:
'   True if the transfer was successful in SAP (if online), False otherwise.
Function transferMaterial(ByVal itemNum As Long, ByVal batch As String, QTY As Variant, fromLoc As String, toLoc As String, toSS As Boolean) As Boolean

  ' Declare variables
  Dim x As Long
  Dim onl As Boolean
  Dim thisMove As GuiSession

  On Error GoTo ErrHandler ' Error handling

  ' Check if SAP is online
  ' (Assumes "Online" status is in cell AA1 of "Sheet2")
  If onlineMode Then
    onl = True
  Else
    onl = False
  End If
    If toSS Then
        ' Record the transfer in the local "SandTrans" sheet
        x = nextEmptyRow(ThisWorkbook.Sheets("SandTrans"), 1, 2) ' Get next empty row
        With ThisWorkbook.Sheets("SandTrans")
            .Cells(x, 1).Value = itemNum
            .Cells(x, 2).Value = UCase(batch)     ' Store batch in uppercase
            .Cells(x, 3).Value = QTY
            .Cells(x, 4).Value = fromLoc
            .Cells(x, 5).Value = toLoc
            .Cells(x, 6).Value = onl             ' Record online/offline status
        End With
    End If
  ' Save the transfer data to the backup spreadsheet
    If productionMode Then
        saveRemoteSandTransfers x
    End If
  ' If SAP is online, perform the transfer in SAP
  If onlineMode Then
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

      .FindById("wnd[0]/usr/txtLAGP-RKAPV").Text = CDec(QTY) ' Quantity as decimal
      .FindById("wnd[0]/usr/txtT001L-LGORT").Text = toLoc
      .FindById("wnd[0]").SendVKey 5 ' Enter

      ' Check for success message in SAP
      If .FindById("wnd[0]/usr/txtV_MESSAGE3").Text = "Created." Then
        transferMaterial = True
      Else
        transferMaterial = False
      End If

      ' Close the transaction in SAP
      .FindById("wnd[0]").SendVKey 8 ' F8 (possibly to save)
      .FindById("wnd[0]/tbar[0]/btn[3]").Press ' Back button
    End With

    Set thisMove = Nothing ' Release the SAP session object
  End If

  Exit Function ' Normal exit

ErrHandler:
  StdErrorHandler "Error " & Err & ": " & Error(Err) & " in transferMaterial Function of Module1"
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
    On Error GoTo ErrHandler ' Error handling
    
    ' Open the remote workbook if it's not already open
    If Not IsWorkBookOpen(filePath) Then
      Workbooks.Open filename:=filePath, ignorereadonlyrecommended:=True
    End If
    
    ' Set workbook and worksheet objects
    Set thisWB = ThisWorkbook
    Set thisSheet = thisWB.Sheets("SandTrans")
    Set remoteWB = Workbooks("OvenLog.xlsx")
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

Sub SaveAndCloseWorkbook()
    LogEvent "Closing workbook", "Application"
    ' Save the active workbook
    ActiveWorkbook.Save

    ' Close the active workbook
    ActiveWorkbook.Close

End Sub
Function addZeros(sourceNum As Integer) As String
'pads a an integer with 0
Dim thisLen As Long
thisLen = Len(CStr(sourceNum))
Select Case thisLen
    Case 1
        addZeros = "000" & CStr(sourceNum)
    Case 2
        addZeros = "00" & CStr(sourceNum)
    Case 3
        addZeros = "0" & CStr(sourceNum)
    Case 4
        addZeros = CStr(sourceNum)
End Select
End Function
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
Function nextEmptyRow(ByRef shet As Worksheet, col As Long, Optional rowStart As Long = 6) As Long

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

Sub printHULabel(ByVal huNum As String, ByVal numLabels As Integer)
    Dim thisLabel As GuiSession
    Set thisLabel = SAPSession.CurrentSession
    With thisLabel
        .FindById("wnd[0]/tbar[0]/okcd").Text = "/nzpplabel"
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/usr/btn%P150004_1000").Press
        .FindById("wnd[0]/usr/ctxtS_HUNO-LOW").Text = huNum
        .FindById("wnd[0]/usr/cmbP_ORIGIN").key = "USA"
        .FindById("wnd[0]/usr/txtP_NUM").Text = numLabels
        .FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName

        .FindById("wnd[0]/tbar[1]/btn[8]").Press
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
        .FindById("wnd[0]/tbar[0]/btn[3]").Press
    End With
    Set thisLabel = Nothing
End Sub
