VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQualityEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrItemNum As String
Private mstrBatchNum As String
Private cBowValue As Variant
Private cLength As Variant
Private cBowCalc As Variant
Private mAcceptable As String
Function characteristicEntry() As Boolean
    ' This subroutine enters beta test results and usage decision into an inspection lot in SAP.
    ' It uses the QA32 transaction to access the inspection lot and updates the relevant fields with data from variables.
    Dim result As Boolean
    
    Dim thisSession As GuiSession ' Declare a variable to hold the SAP GUI session object
On Error GoTo ErrHandler
    Set thisSession = SAPSession.CurrentSession   ' Initialize the SAP GUI session

    ' --- Navigate to the inspection lot screen ---
    If Not SAPStatus.SAPWindowName = "Inspection Lot Selection" Then ' Check if the current screen is already the "Inspection Lot Selection" screen
'        setStartScreen ' If not, call the setStartScreen function to navigate to the main screen
        thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nqa33" ' Enter the transaction code QA32 in the command field
        thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter to execute the transaction
    End If
    result = False
    ' --- Enter selection criteria for the inspection lot ---
    thisSession.FindById("wnd[0]/usr/ctxtQL_ENSTD-LOW").Text = Format(DateAdd("m", -1, Now()), "mm/dd/yyyy") ' Enter the inspection start date
    thisSession.FindById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").Text = Format(Now(), "mm/dd/yyyy")
    thisSession.FindById("wnd[0]/usr/ctxtQL_WERKS-LOW").Text = "Q105" ' Enter the plant code
    thisSession.FindById("wnd[0]/usr/ctxtQL_CHARG-LOW").Text = mstrBatchNum ' Enter the batch number (from a variable mstrBatchNum)
    thisSession.FindById("wnd[0]/usr/ctxtVARIANT").Text = "/NQp"
    thisSession.FindById("wnd[0]/usr/radP_NO_UD").Select
    ' --- Execute the search and navigate to the results ---
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press ' Press the "Execute" button (F8)
    
    If thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").CurrentCellColumn = "" Then 'Batch/Item combo doesn't work
        MsgBox "That item and batch combination doesn't have any QM data waiting to be recorded in SAP", vbOKOnly
        characteristicEntry = result
        Exit Function
    End If
    Dim numBatches, t, numberToProcess As Integer
    Dim qties() As Variant
    Dim processAll As Boolean
    processAll = False
    numBatches = thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
    If numBatches > 1 Then 'if more than one entry for this batch
        Dim totQty As Variant
        totQty = 0
        ReDim qties(numBatches - 1)
        For t = 0 To numBatches - 1
            qties(t) = thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(t, "LOSMENGE")
            totQty = totQty + qties(t)
        Next t
        If MsgBox("There are " & numBatches & " entries for this batch/item number combination. The total quantity awaiting quality data is " & totQty & ". Would you like to apply this to all of them?", vbYesNo) = vbYes Then
            'enter this data for all displayed
            processAll = True
            thisSession.FindById("wnd[0]/tbar[1]/btn[5]").Press
            numberToProcess = numBatches
        Else
            'select the one you want
            Dim stopHere As Boolean
            stopHere = False
            t = 0
            Do While Not stopHere
                If MsgBox("Do you want to use the entry with quantity " & qties(t) & "?", vbYesNo) = vbYes Then
                    'use this entry
                    thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = t
                    stopHere = True
                    numberToProcess = 1
                End If
                t = t + 1
                If t = numBatches Then t = 0 'start the list again
            Loop
        End If
    Else
        numberToProcess = 1
        thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
    End If
    thisSession.FindById("wnd[0]/tbar[1]/btn[44]").Press ' Press the "Record results" button
    For t = 1 To numberToProcess
        If SAPStatus.IsThereAModalWindow Then
            ' --- Enter the inspection data ---
'            thisSession.FindById("wnd[0]/usr/subPRUEFPUNKT:SAPLQAPP:0100/ctxtQAPPD-USERC1").Text = "1" ' Enter a value in the "User field 1"
'            thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
        End If
        thisSession.FindById("wnd[0]/usr/subPRUEFPUNKT:SAPLQAPP:0100/ctxtQAPPD-USERC1").Text = "1" ' Enter a value in the "User field 1"
        thisSession.FindById("wnd[0]").SendVKey 0
        ' --- Enter the beta test results ---
        thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,0]").Text = cBowValue ' Enter the bow value (from a variable msngBetaVal)
        thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,1]").Text = cLength ' Enter the measured length (from a variable msngBakeHours)
    
        ' --- Calculate and retrieve the ppm value ---
        thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
        thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnSELECT_ALL").Press ' Press the "Select all" button
        thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnVALUATION").Press ' Press the "Valuation" button
        Do While SAPStatus.IsThereAModalWindow
            thisSession.FindById("wnd[0]/tbar[0]/btn[0]").Press
        Loop

        cBowCalc = CDec(thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/lblQAQEE-SUMPLUS[8,2]").Text) ' Retrieve the bow limit value from the label
        thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnCLOSING").Press ' Press the "Closing" button
    
        ' --- Enter the usage decision ---
 '       thisSession.FindById("wnd[0]/tbar[0]/btn[0]").Press ' Press the "Back" button
        thisSession.FindById("wnd[0]/tbar[0]/btn[11]").Press ' Press the "Save" button
        If SAPStatus.IsThereAModalWindow Then thisSession.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Press "Continue" in the modal window (if it appears)
    Next t
    For t = 1 To numberToProcess
        If processAll Then thisSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = t - 1
        thisSession.FindById("wnd[0]/tbar[1]/btn[41]").Press ' Press the "Usage decision" button
        thisSession.FindById("wnd[0]/usr/tabsUD_DATA/tabpPLMK/ssubSUB_UD_DATA:SAPMQEVA:0101/ssubUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE").Text = mAcceptable ' Enter the usage decision code (from a variable mAcceptable)
        ' --- Save the usage decision ---
        thisSession.FindById("wnd[0]").SendVKey 0 ' Press Enter
        thisSession.FindById("wnd[0]/tbar[0]/btn[11]").Press ' Press the "Save" button
    Next t
    result = True
    characteristicEntry = result
    ' --- Clean up ---
    'thisSession.findById("wnd[0]/tbar[1]/btn[14]").press ' (Optional) Press the "Back" button to return to the previous screen
    Set thisSession = Nothing ' Release the SAP GUI session object
Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in Function characteristicEntry of clsQualityEntry"
    
End Function

Public Property Let itemNum(rData As String)
    mstrItemNum = rData
End Property

Public Property Let batchNum(rData As String)
    mstrBatchNum = rData
End Property

Public Property Let bowValue(rData As Variant)
    cBowValue = rData
End Property
Public Property Let length(rData As Variant)
    cLength = rData
End Property
Public Property Get calculatedBow() As Variant
    calculatedBow = cBowCalc
End Property
Public Property Let acceptable(rData As Boolean)
    If rData Then
        mAcceptable = "ACC"
    Else
        mAcceptable = "REJ"
    End If
End Property

