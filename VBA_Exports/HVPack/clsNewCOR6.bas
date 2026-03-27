VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNewCOR6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cScrap As Boolean
Dim cProcOrdNum As String, cOper As String, cComment As String, cScrapCode As String, cFinItem As String, cFinItemDesc As String, cTubeSerial As String
Dim cSource1Loc As String, cSource2Loc As String, cBatch1Num As String, cBatch2Num As String, cQTY As String
Dim cResNum As String, cAddResNum As String, cEquipmentUsed As String, cLocation As String, cTransType As String
Dim cYield As Integer
Dim cNumHrs As Single, cNumSetupHrs As Single
Dim cApplyNumMatl As Long
Dim cApplyMatl As Boolean
Dim cCloseOrder As Boolean
Dim cor6GUI As GuiSession
Dim cMatlArray As Variant
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
    cor6GUI.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncor6"
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


Private Function stepOne() As Boolean
    Dim keepGoing As Boolean
On Error GoTo ErrHandler
    
    stepOne = True
    keepGoing = True
    cor6GUI.FindById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = cProcOrdNum
    cor6GUI.FindById("wnd[0]/usr/txtCORUF-VORNR").Text = cOper
    cor6GUI.FindById("wnd[0]").SendVKey 0
    PauseMe 0.25, False
    
    If SAPStatus.CurrentStatus Like "Phase/operation " & cOper & " in process order " & cProcOrdNum & " does not exist (please check entry)" Then
        stepOne = False
        cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
        keepGoing = False
        Exit Function
    ElseIf SAPStatus.IsThereAModalWindow And SAPStatus.SAPWindowName = "Status management: Confirm order" Then
        If MsgBox("Process order " & cProcOrdNum & " shows that it is complete. Do you want to add to it?", vbYesNo) = vbNo Then
            cor6GUI.FindById("wnd[1]/usr/btnOPTION1").Press
            stepOne = False
            keepGoing = False
            Exit Function
        Else
            cor6GUI.FindById("wnd[1]/usr/btnOPTION2").Press
        End If

'    ElseIf SAPStatus.CurrentStatus Like "Order " & cProcOrdNum & " not found, check entry" Then
'        stepOne = False
'        cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
'        KeepGoing = False
'        Exit Function
    End If
    cFinItem = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/ctxtCAUFVD-MATNR").Text
    cFinItemDesc = cor6GUI.FindById("wnd[0]/usr/subVAR_OPR_10:SAPLCORU:5830/txtCAUFVD-MATXT").Text
    If cCloseOrder Then
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
'        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/ctxtAFRUD-PERNR").Text = "97747"
'        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/ctxtAFRUD-ARBPL").Text = cEquipmentUsed
        cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/txtAFRUD-LTXA1").Text = cComment
 '       cor6GUI.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpRZUS/ssubVAR_CNF_10:SAPLCORU:5870/radCORUF-ENDRU").Select
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
  '              MsgBox "Say something about the material application failure and continuing step 40."
                stepOne = False
                Exit Function
            End If
        Else
            
        End If
        
        Dim b As Integer
            b = 1
        If cor6GUI.ActiveWindow.Text = "Process Order Confirmation Enter : Details" And SAPStatus.CurrentStatus = vbNullString Then
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[11]").Press
            cor6GUI.FindById("wnd[0]").SendVKey 0
            cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press
        ElseIf Not SAPStatus.CurrentStatus = "Confirmation saved (Goods movements: " & b & ", failed: 0)" Then
            stepOne = True
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
    Dim x As Integer 'to hold the array position
    Dim SAPRowNum As Long
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
        Dim rowNum As Integer 'to hold the row of the material
        .FindById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnMALL").Press
        .FindById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnDELE").Press
        
        x = 0
        SAPRowNum = 0
        Do While x < cApplyNumMatl
        
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0," & SAPRowNum & "]").Text = cMatlArray(x, 0)
            .FindById("wnd[0]").SendVKey 0
            PauseMe 0.21, False
            
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/txtCOWB_COMP-ERFMG[2," & SAPRowNum & "]").Text = cMatlArray(x, 1)
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-LGORT[4," & SAPRowNum & "]").Text = cMatlArray(x, 2)
            If cMatlArray(x, 3) <> "" Then .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5," & SAPRowNum & "]").Text = cMatlArray(x, 3)
            .FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-BWART[6," & SAPRowNum & "]").Text = cMatlArray(x, 4)
                
            .FindById("wnd[0]").SendVKey 0
            x = x + 1
            SAPRowNum = SAPRowNum + 1
        Loop
            
            'save it
        .FindById("wnd[0]").SendVKey 0
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
        ' --- Print a ticket (if enabled) ---
    If SAPStatus.IsThereAModalWindow Then
        If SAPStatus.SAPWindowName = "Print Label" Then
            If printYes Then ' Check if printing is enabled (using a variable printMe)
                cor6GUI.FindById("wnd[1]/usr/btnDY_VAROPTION1").Press ' Press the "Print with changed parameters" button
                cor6GUI.FindById("wnd[0]/usr/ctxtP_PRINT").Text = printerName ' Set the printer name

                cor6GUI.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Press "Enter" (likely F8)
                cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" in the modal window
                cor6GUI.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" again
            Else
                cor6GUI.FindById("wnd[1]/usr/btnCANCEL").Press ' Press the "Cancel" button if printing is not enabled
            End If
        End If
    End If

    cor6GUI.FindById("wnd[0]/tbar[0]/btn[3]").Press ' Press the "Back" button to exit "Goods Movements"

    
Exit Function
    
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function applyArrayMaterial of clsCOR6"
    
End Function



Private Sub Class_Initialize()
On Error GoTo ErrHandler
    Set cor6GUI = SAPSession.CurrentSession
    cScrap = False
    cCloseOrder = False
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
Property Let closeOrder(ByVal closeOrder As Boolean)
On Error GoTo ErrHandler

    cCloseOrder = closeOrder

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let closeOrder of clsCOR6."
    
End Property

Property Let applyMatl(ByVal applyMatl As Boolean)
On Error GoTo ErrHandler

    cApplyMatl = applyMatl

    Exit Property
ErrHandler:
    
    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Let applyMatl of clsCOR6."
    
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




