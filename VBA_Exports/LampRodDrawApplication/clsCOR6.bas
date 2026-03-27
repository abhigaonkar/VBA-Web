VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCOR6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cScrap As Boolean
Dim cScrapQty As Variant
Dim cProcOrdNum As String, cOper As String, cComment As String, cScrapCode As String, cFinItem As String, cFinItemDesc As String, cTubeSerial As String
Dim cSource1Loc As String, cSource2Loc As String, cBatch1Num As String, cBatch2Num As String, cQTY As String
Dim cResNum As String, cEquipmentUsed As String, cLocation As String, cTransType As String
Dim cYield As Integer
Dim cActualTubeWeightDelta As Variant
Dim cByProductFactor As Variant
Dim cNumHrs As Single, cNumSetupHrs As Single
Dim cApplyNumMatl As Long
Dim cSandQtyPerTube As Single
Dim cApplyMatl As Boolean
Dim cExpectedSandType As String
Dim cor6GUI As GuiSession
Dim cMatlArray As Variant
Dim cAvailableSand As Variant
'below variables for byProduct processing
Dim cAddResNum As String ' only need to set this, the boolean will set automagically
Dim cByProduct As Boolean
Dim cThisBatch As String
Dim cByProductBatch As String

Private Function stepOne() As Boolean
    ' This function performs the first step of the COR6 transaction in SAP.
    ' It enters the process order and operation details, handles yield and scrap,
    ' applies materials, and saves the confirmation.
    ' Returns True if the step is completed successfully, False otherwise.

    Dim numHrs As Single
    Dim mainWindow As GuiMainWindow
    Dim keepGoing As Boolean
    On Error GoTo ErrHandler ' Error handling

    stepOne = True ' Initialize return value to True
    keepGoing = True ' Flag to control the flow of execution

    ' --- Enter process order and operation details ---
    cor6GUI.FindById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = cProcOrdNum ' Set the process order number (from a variable cProcOrdNum)
    cor6GUI.FindById("wnd[0]/usr/txtCORUF-VORNR").Text = cOper ' Set the operation number (from a variable cOper)
    cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
    PauseMe 0.25, False ' Pause for 0.25 seconds (using a function PauseMe not defined here)

    ' --- Check if the operation exists ---
    If SAPStatus.CurrentStatus Like "Phase/operation " & cOper & " in process order " & cProcOrdNum & " does not exist (please check entry)" Then
        stepOne = False ' Set return value to False if the operation does not exist
        cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press ' Press the "Back" button
        keepGoing = False ' Set the flag to False to skip further processing
        Exit Function ' Exit the function
    End If

    ' --- Get finished item details ---
    cFinItem = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/ctxtCAUFVD-MATNR").Text ' Get the finished item number
    cFinItemDesc = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/txtCAUFVD-MATXT").Text ' Get the finished item description

    ' --- Enter yield or scrap quantity ---
    If keepGoing Then
'        If Not cYield = 0 Then
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-LMNGA").Text = cYield ' Set the yield quantity (from a variable cYield)
'        ElseIf cScrap Then
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-XMNGA").Text = cScrapQty ' Set the scrap quantity to 1
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/ctxtAFRUD-GRUND").Text = cScrapCode ' Set the scrap code (from a variable cScrapCode)
'        End If

        ' --- Enter hours ---
        If Not cNumHrs = 0 Then
            ' Code for handling setup hours (commented out)
            ' If cNumSetupHrs < 0 Or cNumSetupHrs > 0 Then
            '     cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-ISM01").Text = cNumSetupHrs
            ' End If

            If cNumHrs > 100 Then ' Check if the number of hours is greater than 100
                cNumHrs = cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtCORUF-SLM02").Text ' Get the default hours from SAP
                ActiveCell.Select ' Select the active cell in Excel (presumably to store the default hours)
                ActiveCell.Value = cNumHrs ' Set the cell value to the default hours
                ActiveCell.Offset(1, 0).Select ' Move to the next cell
            End If
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-ISM02").Text = cNumHrs ' Set the number of hours
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/ctxtAFRUD-ILE02").Text = "hr" ' Set the unit to "hr"
            cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
        End If

        cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter
        ' --- Handle quantity confirmation errors ---
        If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
            cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter again if there is a quantity confirmation error
            If SAPStatus.CurrentStatus Like "Total confirmation quantity not equal to planned confirmation quantity" Then
                cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter again if the error persists
            End If
        End If

        ' --- Enter personnel and equipment details ---
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS").Select ' Select the "Personnel" tab
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/ctxtAFRUD-ARBPL").Text = cEquipmentUsed ' Set the equipment used (from a variable cEquipmentUsed)
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/txtAFRUD-LTXA1").Text = cComment ' Set the comment (from a variable cComment)
        cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter

        ' --- Handle quantity confirmation errors again ---
        If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
            cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter again if there is a quantity confirmation error
            If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
                cor6GUI.FindById("wnd[0]").SendVKey 0 ' Press Enter again if the error persists
            End If
        End If

        ' --- Apply materials ---
        If cApplyMatl Then ' Check if materials should be applied (based on a variable cApplyMatl)
            If Not applyArrayMaterial() Then ' Call a function to apply materials
                stepOne = False ' Set return value to False
                Exit Function ' Exit the function
            Else
                stepOne = True
            End If
        End If

        ' --- Save the confirmation ---
        Dim b As Integer
        b = 1 ' Initialize a counter ()
        If cor6GUI.ActiveWindow.Text = "Process Order Confirmation Enter : Details" And SAPStatus.CurrentStatus = vbNullString Then
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[11]").Press
            cor6GUI.FindById("wnd[0]").SendVKey 0
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
        ElseIf Not SAPStatus.CurrentStatus = "Confirmation saved (Goods movements: " & b & ", failed: 0)" Then
            stepOne = False
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[11]").Press
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
            Exit Function
        End If
        If SAPStatus.CurrentStatus Like "Total confirmation quantity not equal to planned confirmation quantity" Then
            cor6GUI.FindById("wnd[0]").SendVKey 0
            If SAPStatus.CurrentStatus Like "Total confirmation quantity not equal to planned confirmation quantity" Then
                cor6GUI.FindById("wnd[0]").SendVKey 0
            End If
        End If
        PauseMe 0.5, True
        If SAPStatus.IsThereAModalWindow Then SAPStatus.CheckModalWindow
    End If
    stepOne = True
    Exit Function
ErrHandler:
    If Err.Number = 619 Then
        Err.Clear
        Resume Next
    End If
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in stepOne of clsCOR6"
    
End Function
Private Function applyArrayMaterial() As Boolean
    ' This function applies materials (primarily sand) to the current process order in SAP during a COR6 transaction.
    ' It handles custom batch numbers, applies materials based on available quantities, and manages byproducts.
    ' Returns True if the materials are applied successfully, False otherwise.

    Dim result As Boolean ' Result of the material application (True = success, False = failure)
    Dim SAPRowNum As Long ' Row number in the SAP "Goods Movements" table
On Error GoTo ErrHandler ' Error handling

    result = True ' Initialize result to True

    With cor6GUI ' Using With statement for better readability
        ' --- Open the "Goods Movements" screen ---
        .FindById("wnd[0]/tbar[1]/btn[18]").Press ' Press the button to open "Goods Movements"

        ' Handle any modal windows that might appear
        If SAPStatus.IsThereAModalWindow Then
            .FindById("wnd[1]").SendVKey 4 ' Press "Enter" in the modal window
            .FindById("wnd[2]/tbar[0]/btn[0]").Press ' Press "Continue" in the next modal window
            .FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" again
            .FindById("wnd[0]").SendVKey 0 ' Press "Enter" in the main window
        End If
         ' if this is a startup or scrap order, receive 0
       cResNum = .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0,0]").Text

        Select Case cResNum
            Case "2107046", "111215", "2107047", "141316"
                .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/txtCOWB_COMP-ERFMG[2,0]").Text = 0
        End Select

        'handle a blank location for GR
        If .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4,0]").Text = "" Then
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4,0]").Text = "0020"
        End If
        ' --- Prepare for material application ---
        Dim rowNum As Integer ' Row number of the material in the cAvailableSand array
        Dim totSandNeed As Single, remainingSandNeed As Single ' Variables to store the total and remaining sand needed
        ' (Note: The comment indicates that these variable names might be misleading)
        cExpectedSandType = .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0,1]").Text
        cThisBatch = .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5,0]").Text
        totSandNeed = .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/txtCOWB_COMP-ERFMG[2,1]").Text  ' Get the total sand needed from SAP
        
        remainingSandNeed = totSandNeed ' Initialize remainingSandNeed with the total sand needed
        Dim needSand As Boolean ' Flag to indicate if sand is still needed

        rowNum = 1 ' Set the starting row number for material application


        ' --- Apply materials (sand) ---
        needSand = True ' Set the flag to True, indicating that sand is needed
        SAPRowNum = rowNum ' Initialize the SAP row number with the starting row number
        rowNum = 0 ' Reset rowNum to 0 for iterating through the cAvailableSand array

        ' Select and delete the first material entry in the SAP table
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(1).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(2).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(3).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(4).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(5).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0,1]").SetFocus
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0,1]").CaretPosition = 0
        .FindById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnDELE").Press

        ' Loop through available sand and apply it to the process order
        Do While needSand ' Loop until all the required sand is applied
            cResNum = cAvailableSand(rowNum, 0) ' Get the resource number from the cAvailableSand array
            If cResNum <> cExpectedSandType Then
                'MsgBox "FYI - This order is expecting sand item " & cExpectedSandType & ". But you are issuing " & cResNum & " to it. This is just for informational purposes. The transaction is being saved and marked for later review.", vbCritical
                rowAlert (nextEmptyRow(localWS, 1, 2) - 1)
            End If
            ' Check if the available sand is enough to cover the remaining need
            If CSng(cAvailableSand(rowNum, 2)) >= CSng(totSandNeed) Then
                cQTY = totSandNeed ' If yes, set the quantity to the total sand needed
                needSand = False ' Set the flag to False, indicating that no more sand is needed
            Else
                cQTY = cAvailableSand(rowNum, 2) ' If not, use the available sand quantity
                totSandNeed = totSandNeed - cQTY ' Update the remaining sand needed
            End If
            If cLocation <> "" Then
                cLocation = furnName ' Set the storage location from the furnName variable
            End If
                
            cBatch1Num = cAvailableSand(rowNum, 1) ' Set the batch number from the cAvailableSand array
            cTransType = "261" ' Set the movement type (hardcoded to "261")

            ' Call the nextGMRow function to enter the material details in the SAP table
            nextGMRow SAPRowNum, cResNum, cLocation, cBatch1Num, Round(cQTY, 3), cTransType

            rowNum = rowNum + 1 ' Increment the row number in the cAvailableSand array
            SAPRowNum = SAPRowNum + 1 ' Increment the row number in the SAP table
        Loop

        ' --- Handle byproduct ---
        If cByProduct Then ' Check if there is a byproduct
            ' Call the nextGMRow function to enter the byproduct details in the SAP table
            nextGMRow SAPRowNum, cAddResNum, "0300", cByProductBatch, Round(totSandNeed * cByProductFactor, 0), 531
        End If

        ' --- Save the goods movements ---
        .FindById("wnd[0]").SendVKey 0 ' Press "Enter"
        .FindById("wnd[0]/tbar[0]/btn[11]").Press ' Press the "Save" button

        ' --- Handle potential "Error in goods movements" ---
        If SAPStatus.CurrentStatus = "Error in goods movements" Then
            .FindById("wnd[0]/tbar[0]/btn[3]").Press ' Press the "Back" button
            .FindById("wnd[1]/usr/btnSPOP-OPTION2").Press ' Press the "Cancel" button in the modal window
        End If
    
        If SAPStatus.IsThereAModalWindow And SAPStatus.SAPWindowName = "Error in actual cost calculation" Then
            .FindById("wnd[1]/usr/btnSPOP-OPTION2").Press
            .FindById("wnd[1]/usr/btnBUTTON_1").Press
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
        End If
        If SAPStatus.CurrentStatus = "Error in goods movements" Then 'saves a COGI error if present
            .FindById("wnd[0]/tbar[0]/btn[3]").Press
            .FindById("wnd[1]/usr/btnSPOP-OPTION2").Press
        End If
    End With
    PauseMe 0.25, True ' Pause for 0.25 seconds

    applyArrayMaterial = result ' Return the result of the material application

    ' --- Print a ticket (if enabled) ---
    If SAPStatus.IsThereAModalWindow Then
        If SAPStatus.SAPWindowName = "Print Label" Then
            If printMe Then ' Check if printing is enabled (using a variable printMe)
                cor6GUI.FindById("wnd[1]/usr/btnDY_VAROPTION1").Press ' Press the "Print with changed parameters" button
                cor6GUI.FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName ' Set the printer name
                cor6GUI.FindById("wnd[0]/usr/txtP_IGMNG").Text = cYield
                'cor6GUI.FindById("wnd[0]/usr/ctxtP_GMEIN").Text =
                cor6GUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Press "Enter" (likely F8)
                cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" in the modal window
    '            cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" again
            Else
                cor6GUI.FindById("wnd[1]/usr/btnCANCEL").Press ' Press the "Cancel" button if printing is not enabled
            End If
        End If
    End If

    cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press ' Press the "Back" button to exit "Goods Movements"
    If cByProduct Then
        If printMe Then printLabels cProcOrdNum, cAddResNum, cByProductBatch, Round(totSandNeed * cByProductFactor, 0), "KG", 1, printerName
    End If
Exit Function ' Exit the function

ErrHandler:
    ' --- Error handling ---
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function applyArrayMaterial of clsCOR6" ' Call the standard error handler
End Function
Private Sub nextGMRow(ByVal rowNum As Integer, itemNum As String, locN As String, Batch As String, ByVal QTY As Single, ByVal mvmtType As Integer)

    With cor6GUI
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0," & rowNum & "]").Text = itemNum
        .FindById("wnd[0]").SendVKey 0
        PauseMe 0.71, False
        'make sure the item number is valid
        If Not SAPStatus.statusCheck("Confirmation saved (Goods movements: 1, failed: 0)", False) Then
            If MsgBox("SAP Error Message: " & SAPStatus.CurrentStatus & " Would you like to save step 40 and fix the feedstock issue later?", vbYesNo) = vbNo Then
                MsgBox "Nothing has been saved."
            Else
            End If
            .FindById("wnd[0]/tbar[0]/btn[12]").Press
            .FindById("wnd[0]/tbar[0]/btn[15]").Press
            .FindById("wnd[1]/usr/btnSPOP-OPTION2").Press
            Exit Sub
        End If
        
        
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/txtCOWB_COMP-ERFMG[2," & rowNum & "]").Text = QTY
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4," & rowNum & "]").Text = locN
        If Not Batch = "" Then
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5," & rowNum & "]").Text = Batch
        End If
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-BWART[6," & rowNum & "]").Text = mvmtType
        .FindById("wnd[0]").SendVKey 0
        If SAPStatus.CurrentStatus = "Batch " & itemNum & " Q105 " & locN & " " & Batch & " does not exist" Then
            .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[1]").SendVKey 4
            .FindById("wnd[2]/tbar[0]/btn[0]").Press
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
        End If

    End With
        
End Sub

Public Function saveIt() As Boolean
    'returns true if the transaction is finished
    'The process goes as follows:
    'stepOne to enter the time and user for the FS (step 20)
    'step20FS to enter the details about the FS
    'stepOne to enter the details of the finished tube step (40)

    
On Error GoTo ErrHandler
    setStartScreen
    saveIt = True
    cor6GUI.FindById("wnd[0]/tbar[0]/okcd").Text = "cor6"
    cor6GUI.FindById("wnd[0]").SendVKey 0
    If Not stepOne() Then
        saveIt = False
        Exit Function
    End If
    
    Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function SaveIt of clsCOR6"
    
End Function
Private Sub Class_Initialize()
On Error GoTo ErrHandler
    Set cor6GUI = SAPSession.CurrentSession
    cScrap = False
    cByProduct = False

    Exit Sub
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " while initializing clsCOR6"
End Sub
Property Let yield(ByVal yield As String) 'yield
On Error GoTo ErrHandler

    cYield = yield

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let yield of clsCOR6."
    
    
End Property
Property Let availableSand(ByVal newValue As Variant)  'array of sand available
On Error GoTo ErrHandler

    cAvailableSand = newValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let availableSand of clsCOR6."
    
    
End Property
Property Let scrapQty(ByVal newValue As Variant)  'scrap quantity
On Error GoTo ErrHandler

    cScrapQty = newValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let scrapQty of clsCOR6."
    
    
End Property
Property Let actualTubeWeightDelta(ByVal newValue As Variant)  'receive actual tube weight from UI
On Error GoTo ErrHandler

    cActualTubeWeightDelta = CDec(newValue)

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let availableSand of clsCOR6."
    
    
End Property
Property Let byProductFactor(ByVal newValue As Variant)  'receive actual tube weight from UI
On Error GoTo ErrHandler

    cByProductFactor = CDec(newValue)

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let byProductFactor of clsCOR6."
    
    
End Property

Property Let matlArray(ByVal matlArray As Variant)  'array of materials to be applied
On Error GoTo ErrHandler

    cMatlArray = matlArray

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let matlArray of clsCOR6."
    
    
End Property

Property Let transType(ByVal transType As String) 'transaction type
On Error GoTo ErrHandler

    cTransType = transType
    
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let transType of clsCOR6."
    
End Property
Property Let location(ByVal location As String) 'location of the item
On Error GoTo ErrHandler

    cLocation = location

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let location of clsCOR6."
    
    
End Property

Property Let applyNumMatl(ByVal applyNumMatl As Long)
On Error GoTo ErrHandler

    cApplyNumMatl = applyNumMatl

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let applyNumMatl of clsCOR6."
    
End Property

Property Let applyMatl(ByVal applyMatl As Boolean)
On Error GoTo ErrHandler

    cApplyMatl = applyMatl

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let applyMatl of clsCOR6."
    
End Property
Property Let byProduct(ByVal newData As Boolean)
On Error GoTo ErrHandler

    cByProduct = newData

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let byProduct of clsCOR6."
    
End Property


Property Let scrapCode(ByVal scrapCode As String)
On Error GoTo ErrHandler

    cScrapCode = scrapCode
    cScrap = True

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let scrapCode of clsCOR6."
    
End Property
Property Let equipmentUsed(ByVal equipmentUsed As String)
On Error GoTo ErrHandler

    cEquipmentUsed = equipmentUsed

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let equipmentUsed of clsCOR6."
    
End Property
Property Let Source1Loc(ByVal Source1Loc As String)
On Error GoTo ErrHandler

    cSource1Loc = Source1Loc

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let Source1Loc of clsCOR6."
    
End Property
Property Let Source2Loc(ByVal Source2Loc As String)
On Error GoTo ErrHandler

    cSource2Loc = Source2Loc

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let Source2Loc of clsCOR6."
    
End Property
Property Let comment(ByVal comment As String)
On Error GoTo ErrHandler

    cComment = comment

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let comment of clsCOR6."
    
End Property
Property Let procOrdNum(ByVal procOrdNum As String)
On Error GoTo ErrHandler

    cProcOrdNum = procOrdNum

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let ProcOrdNum of clsCOR6."
    
End Property
Property Let tubeSerialNumber(ByVal vNewValue As String)
On Error GoTo ErrHandler

    cTubeSerial = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let tubeSerialNumber of clsCOR6."
    
End Property

Property Let oper(ByVal vNewValue As String)
On Error GoTo ErrHandler

    cOper = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let Oper of clsCOR6."
    
End Property

Property Let numHrs(ByVal vNewValue As Single)
On Error GoTo ErrHandler

    cNumHrs = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let fsHrs of clsCOR6."
    
End Property
Property Let numSetupHrs(ByVal vNewValue As Single)
On Error GoTo ErrHandler

    cNumSetupHrs = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let cNumSetupHrs of clsCOR6."
    
End Property
Private Sub Class_Terminate()
On Error GoTo ErrHandler
    
    Set cor6GUI = Nothing
    Exit Sub

ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Class_Terminate of clsCOR6"
    
End Sub
Property Let resNum(ByVal resNum As String) 'resource number of first item of source stock
On Error GoTo ErrHandler

    cResNum = resNum

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let resNum of clsCOR6."
    
End Property
Property Let Batch1Num(ByVal batchNum As String) 'batch number of first item of source stock
On Error GoTo ErrHandler

    cBatch1Num = batchNum

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let batch1Num of clsCOR6."
    
End Property

Property Let sandInTube(ByVal newVal As Single) 'sand in this tube
On Error GoTo ErrHandler
    cSandQtyPerTube = newVal

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let sandInTube of clsCOR6."
    
    
End Property

Property Let additionalSourceStock(ByVal additionalSourceStock As Boolean) 'is there more source stock?
On Error GoTo ErrHandler

'    cAdditionalSourceStock = additionalSourceStock

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let additionalSourceStock of clsCOR6."
    
    
End Property
Property Let addResNum(ByVal addResNum As String) 'resource number of second item of source stock
On Error GoTo ErrHandler

    cAddResNum = addResNum
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let addResNum of clsCOR6."
    
    
End Property
Property Let byProductBatch(ByVal newData As String) 'batch number of a by product
On Error GoTo ErrHandler

    cByProductBatch = newData

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let byProductBatch of clsCOR6."
    
End Property
Property Let Batch2Num(ByVal Batch2Num As String) 'batch number of second item of source stock
On Error GoTo ErrHandler

    cBatch2Num = Batch2Num

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let Batch2Num of clsCOR6."
    
End Property

Public Property Get FinishedItem() As String
On Err GoTo ErrHandler
    FinishedItem = cFinItem
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get FinishedTube of clsCOR6."
End Property
Public Property Get FinishedItemDesc() As String
On Err GoTo ErrHandler
    FinishedItemDesc = cFinItemDesc
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get FinishedTubeDesc of clsCOR6."
End Property
Public Property Get thisBatch() As String 'for storing the batch saved in SAP
On Err GoTo ErrHandler
    thisBatch = cThisBatch
    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Get thisBatch of clsCOR6."
End Property

Private Function LogText() As String
On Err GoTo ErrHandler
    Dim cLogText As String
    LogText = cLogText
    Exit Function
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in logtext of clsCOR6"
End Function



