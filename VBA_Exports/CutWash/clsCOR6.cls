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
Dim cProcOrdNum As String
Dim cOper As String, cComment As String, cScrapCode As String, cFinItem As String, cFinItemDesc As String, cTubeSerial As String
Dim cSource1Loc As String, cSource2Loc As String, cBatch1Num As String, cBatch2Num As String, cQTY As String
Dim cResNum As String, cAddResNum As String, cEquipmentUsed As String, cLocation As String, cTransType As String
Dim cYield As Integer
Dim cFinalConfirmation As Boolean
Dim cNumHrs As Single, cNumSetupHrs As Single
Dim cApplyNumMatl As Long
Dim cApplyMatl As Boolean
Dim cSerialized As Boolean
Dim cor6GUI As GuiSession
Dim cMatlArray As Variant
Dim cAvailableSand As Variant


Private Function stepOne() As Boolean
    Dim numHrs As Single
    Dim mainWindow As GuiMainWindow
    Dim keepGoing As Boolean
On Error GoTo ErrHandler
    
    stepOne = True
    keepGoing = True
    cor6GUI.FindById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = cProcOrdNum
    cor6GUI.FindById("wnd[0]/usr/txtCORUF-VORNR").Text = "20"
    cor6GUI.FindById("wnd[0]").SendVKey 0
    PauseMe 0.25, False
    
    If SAPStatus.CurrentStatus Like "Phase/operation " & cOper & " in process order " & cProcOrdNum & " does not exist (please check entry)" Then
        stepOne = False
        cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
        keepGoing = False
        Exit Function
    End If
    cFinItem = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/ctxtCAUFVD-MATNR").Text
    cFinItemDesc = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/txtCAUFVD-MATXT").Text
    If cFinalConfirmation Then
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/radCORUF-ENDRU").Select
    End If
    If keepGoing Then
        If Not cYield = 0 Then
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-LMNGA").Text = cYield
        ElseIf cScrap Then
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-XMNGA").Text = "1"
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/ctxtAFRUD-GRUND").Text = cScrapCode

        End If
        
        If Not cNumHrs = 0 Then
'            If cNumSetupHrs < 0 Or cNumSetupHrs > 0 Then
'                cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-ISM01").Text = cNumSetupHrs
'            End If
            If cNumHrs > 100 Then
                cNumHrs = cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtCORUF-SLM02").Text
                ActiveCell.Select
                ActiveCell.Value = cNumHrs
                ActiveCell.Offset(1, 0).Select
                
            End If
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-ISM02").Text = cNumHrs
            cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/ctxtAFRUD-ILE02").Text = "hr"
            cor6GUI.FindById("wnd[0]").SendVKey 0
        End If

        cor6GUI.FindById("wnd[0]").SendVKey 0
        If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
            cor6GUI.FindById("wnd[0]").SendVKey 0
            If SAPStatus.CurrentStatus Like "Total confirmation quantity not equal to planned confirmation quantity" Then
                cor6GUI.FindById("wnd[0]").SendVKey 0
            End If
        End If

        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS").Select
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/txtAFRUD-LTXA1").Text = Left(cComment, 40)
        cor6GUI.FindById("wnd[0]").SendVKey 0
        If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
            cor6GUI.FindById("wnd[0]").SendVKey 0
            If SAPStatus.CurrentStatus Like "Total quantity confirmed not equal to planned quantity to be confirmed" Then
                cor6GUI.FindById("wnd[0]").SendVKey 0
            End If
        End If
        
        ' if cscrap delete all lines and issue sand
        ' else keep first line (changing the batch)
        ' delete other lines and issue sand
        If cApplyMatl Then
            If Not applyArrayMaterial(cTubeSerial) Then
                stepOne = False
                Exit Function
            End If
        End If
        
        Dim b As Integer
            b = 1
        If cor6GUI.ActiveWindow.Text = "Process Order Confirmation Enter : Details" And SAPStatus.CurrentStatus = vbNullString Then
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[11]").Press
            cor6GUI.FindById("wnd[0]").SendVKey 0
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
        ElseIf Not SAPStatus.CurrentStatus = "Confirmation saved (Goods movements: " & b & ", failed: 0)" Then
            'stepOne = False
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
Private Function applyArrayMaterial(Optional cTubeSerial As String) As Boolean
    Dim result As Boolean
    Dim materialCount, x As Integer
    Dim SAPRowNum As Long
    'array structure
    'item    batch   quantity    location    movetype
On Error GoTo ErrHandler
    
    result = True
    With cor6GUI
        .FindById("wnd[0]/tbar[1]/btn[18]").Press
        If SAPStatus.IsThereAModalWindow Then
            .FindById("wnd[1]").SendVKey 4
            .FindById("wnd[2]/tbar[0]/btn[0]").Press
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
            .FindById("wnd[0]").SendVKey 0
        End If
        SAPStatus.statusCheck
        If Not SAPStatus.statusCheck("Confirmation saved (Goods movements: 1, failed: 0)", True) Then
            result = False
            Exit Function
        End If
        
        Dim rowNum As Integer 'to hold the row of the material
        'to take care of custom batch
        If cTubeSerial <> "" Then
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5,0]").Text = cTubeSerial
            .FindById("wnd[0]").SendVKey 0
            If SAPStatus.IsThereAModalWindow Then
                '.FindById("wnd[0]").SendVKey 0
                .FindById("wnd[1]").SendVKey 4
                .FindById("wnd[2]/tbar[0]/btn[0]").Press
                .FindById("wnd[1]/tbar[0]/btn[0]").Press
                .FindById("wnd[0]").SendVKey 0
                rowNum = 1
            End If
        Else
            rowNum = 0
        End If
        If SAPStatus.CurrentStatus = "Enter Storage Location" Then
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4,0]").Text = "hvws"
            .FindById("wnd[0]").SendVKey 0
        End If
        If SAPStatus.IsThereAModalWindow Then
       '     .FindById("wnd[0]").SendVKey 0
            .FindById("wnd[1]").SendVKey 4
            .FindById("wnd[2]/tbar[0]/btn[0]").Press
            .FindById("wnd[1]/tbar[0]/btn[0]").Press
            .FindById("wnd[0]").SendVKey 0
            rowNum = 1
        End If
        'This will apply all materials listed to each order
        SAPRowNum = 1
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(1).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(2).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(3).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(4).Selected = True
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500").GetAbsoluteRow(5).Selected = True
        .FindById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnDELE").Press
        x = 0
        If Not cSerialized Then
            Do While x < UBound(cMatlArray) + 1
                nextGMRow SAPRowNum, cMatlArray(x, 0), cMatlArray(x, 3), cMatlArray(x, 1), cMatlArray(x, 2), cMatlArray(x, 4)
                x = x + 1
                SAPRowNum = SAPRowNum + 1
            Loop
        End If
        'save it
        .FindById("wnd[0]").SendVKey 0
        .FindById("wnd[0]/tbar[0]/btn[11]").Press
        
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

    PauseMe 0.25, True
    applyArrayMaterial = result
    'print a ticket
    If SAPStatus.IsThereAModalWindow Then
        If SAPStatus.SAPWindowName = "Print Label" Then
            If printMe Then
                cor6GUI.FindById("wnd[1]/usr/btnDY_VAROPTION1").Press
                cor6GUI.FindById("wnd[0]/usr/txtP_IGMNG").Text = "1"    ' pieces per ticket
                cor6GUI.FindById("wnd[0]/usr/txtP_NUM").Text = cYield  ' number of tickets
                cor6GUI.FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName
                cor6GUI.FindById("wnd[0]/usr/ctxtP_PRINT").SetFocus
                cor6GUI.FindById("wnd[0]/usr/ctxtP_PRINT").CaretPosition = 10
                cor6GUI.FindById("wnd[0]/tbar[1]/btn[8]").Press
                If SAPStatus.IsThereAModalWindow Then cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press
                If SAPStatus.IsThereAModalWindow Then cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press
            Else
                cor6GUI.FindById("wnd[1]/usr/btnDY_VAROPTION2").Press
            End If

        End If
    End If
    cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
    
Exit Function
    
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function applyArrayMaterial of clsCOR6"
    
End Function

Private Sub nextGMRow(ByVal rowNum As Integer, ByVal itemNum As String, ByVal locN As String, ByVal batch As String, ByVal qty As Single, ByVal mvmtType As Integer)

    With cor6GUI
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0," & rowNum & "]").Text = itemNum
        .FindById("wnd[0]").SendVKey 0
        PauseMe 0.31, False
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
        
        
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/txtCOWB_COMP-ERFMG[2," & rowNum & "]").Text = qty
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4," & rowNum & "]").Text = locN
        If Not batch = "" Then
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5," & rowNum & "]").Text = batch
        End If
        .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-BWART[6," & rowNum & "]").Text = mvmtType
        .FindById("wnd[0]").SendVKey 0
        If SAPStatus.CurrentStatus = "Batch " & itemNum & " Q105 " & locN & " " & batch & " does not exist" Then
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
    If Not setStartScreen Then
        MsgBox "Please return the SAP window to the start screen.", vbOKOnly
        saveIt = False
        Exit Function
    End If
    saveIt = True
    cor6GUI.FindById("wnd[0]/tbar[0]/okcd").Text = "cor6"
    cor6GUI.FindById("wnd[0]").SendVKey 0
    If Not stepOne() Then
        saveIt = False
        
        Exit Function
    End If
    If Not SAPStatus.CheckModalWindow Then
       ' MsgBox "Please check the SAP window and clear any errors. Come back here and click OK when complete.", vbOKOnly
    End If
    
    Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function SaveIt of clsCOR6"
    
End Function
Private Sub Class_Initialize()
On Error GoTo ErrHandler
    Set cor6GUI = SAPSession.CurrentSession
    cScrap = False
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


Property Let qty(ByVal qty As String) 'quantity
On Error GoTo ErrHandler

    cQTY = qty

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let qty of clsCOR6."
    
    
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
Property Let finalConfirmation(ByVal finalConfirmation As Boolean)
On Error GoTo ErrHandler

    cFinalConfirmation = finalConfirmation

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let finalConfirmation of clsCOR6."
    
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
Property Let serialized(ByVal vNewValue As Boolean)
On Error GoTo ErrHandler

    cSerialized = vNewValue

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let serialized of clsCOR6."
    

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
Property Let addResNum(ByVal addResNum As String) 'resource number of second item of source stock
On Error GoTo ErrHandler

    cAddResNum = addResNum

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let addResNum of clsCOR6."
    
    
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


Private Function LogText() As String
On Err GoTo ErrHandler
    Dim cLogText As String
    LogText = cLogText
    Exit Function
    
ErrHandler:
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in logtext of clsCOR6"
End Function




