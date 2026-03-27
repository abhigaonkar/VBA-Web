VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDimEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstritemNum As String
Private mstrbatch As String
Private msngOD1 As Variant
Private msngOD2 As Variant
Private msngOD3 As Variant
Private msngOD4 As Variant
Private msngWALL1 As Variant
Private msngWALL2 As Variant
Private msngWALL3 As Variant
Private msngWALL4 As Variant
Private msngWALL5 As Variant
Private msngWALL6 As Variant
Private msngWALL7 As Variant
Private msngWALL8 As Variant
Private msngBOW As Variant
Private cBowValue As Variant
Private mbooaccOverride As Boolean
Private cODMax As Variant
Private cODMin As Variant
Private cODMean As Variant
Private cWallMax As Variant
Private cWallMin As Variant
Private cWallMean As Variant
Private cSidingMax As Variant
Private cSiding1 As Variant
Private cSiding2 As Variant
Private cActualLength As Variant
Private cWeight As Variant


Private cOvality1 As Variant
Private cOvality2 As Variant
Private cOvalityMax As Variant
Private cGradient As Variant
Private localWS As Worksheet
Private localWB As Workbook


Sub saveDimensions(p As Long)
    Dim thisSession As GuiSession
    Dim x As Long
On Error GoTo ErrHandler
    Set thisSession = SAPSession.CurrentSession
    If Not SAPStatus.SAPWindowName = "Inspection Lot Selection" Then
        
        thisSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nqa33"
        thisSession.FindById("wnd[0]").SendVKey 0
        PauseMe 1, False
    End If
    'doCalcs
    thisSession.FindById("wnd[0]/usr/ctxtQL_ENSTD-LOW").Text = Format(DateAdd("m", -1, Now()), "mm/dd/yyyy")
    thisSession.FindById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").Text = Format(Now(), "mm/dd/yyyy")
    thisSession.FindById("wnd[0]/usr/ctxtQL_WERKS-LOW").Text = "Q105"
    thisSession.FindById("wnd[0]/usr/ctxtQL_MATNR-LOW").Text = localWS.Cells(p, 6)
    thisSession.FindById("wnd[0]/usr/ctxtQL_CHARG-LOW").Text = localWS.Cells(p, 8)
    
    thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press
    thisSession.FindById("wnd[0]/tbar[1]/btn[5]").Press
    thisSession.FindById("wnd[0]/tbar[1]/btn[44]").Press
    
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,0]").Text = localWS.Cells(p, 10) 'OD Top NS
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,1]").Text = localWS.Cells(p, 9) 'OD Top EW
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,2]").Text = localWS.Cells(p, 12) 'OD Bottom NS
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,3]").Text = localWS.Cells(p, 11) 'OD Bottom EW
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,4]").Text = localWS.Cells(p, 14) 'WALL N TOP
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,5]").Text = localWS.Cells(p, 15) 'WALL S TOP
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,6]").Text = localWS.Cells(p, 16) 'WALL E TOP
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,7]").Text = localWS.Cells(p, 13) 'WALL W TOP
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,8]").Text = localWS.Cells(p, 18) 'WALL N BOTTOM
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,9]").Text = localWS.Cells(p, 20) 'WALL S BOTTOM
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,10]").Text = localWS.Cells(p, 19) 'WALL E BOTTOM
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,11]").Text = localWS.Cells(p, 17) 'WALL W BOTTOM
    'thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,12]").Text = localWS.Cells(p, 27) 'OD MEAN  CALC
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,13]").Text = localWS.Cells(p, 28) 'OD MAX
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,14]").Text = localWS.Cells(p, 29) 'OD MIN
    'thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,15]").Text = localWS.Cells(p, 30) 'WALL MEAN  CALC
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,16]").Text = localWS.Cells(p, 31) 'WALL MAX
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,17]").Text = localWS.Cells(p, 32) 'WALL MIN
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,18]").Text = localWS.Cells(p, 33) 'SIDING MAX
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,19]").Text = localWS.Cells(p, 34) 'OVALITY MAX
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,20]").Text = localWS.Cells(p, 35) 'GRADIENT
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/txtQAQEE-SUMPLUS[8,21]").Text = localWS.Cells(p, 21) 'BOW
    
    
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnSELECT_ALL").Press
    Do While right(SAPStatus.CurrentStatus, 16) = "Check your entry"
        For x = 1 To 22
            thisSession.FindById("wnd[0]").SendVKey 0
        Next x
    Loop
    
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnVALUATION").Press
    For x = 1 To 22
        thisSession.FindById("wnd[0]").SendVKey 0
    Next x
    thisSession.FindById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/subSUB_EE_FCODE:SAPLQEEM:5300/btnCLOSING").Press
    Do While SAPStatus.IsThereAModalWindow
        thisSession.FindById("wnd[0]/tbar[0]/btn[0]").Press
    Loop
    thisSession.FindById("wnd[0]/tbar[0]/btn[0]").Press
    'Enter the usage decision
    thisSession.FindById("wnd[0]/tbar[0]/btn[11]").Press
    thisSession.FindById("wnd[0]/tbar[1]/btn[41]").Press
    thisSession.FindById("wnd[0]/usr/tabsUD_DATA/tabpPLMK/ssubSUB_UD_DATA:SAPMQEVA:0101/ssubUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE").Text = "ACC"
    
    thisSession.FindById("wnd[0]").SendVKey 0
    thisSession.FindById("wnd[0]/tbar[0]/btn[11]").Press
    thisSession.FindById("wnd[0]").SendVKey 3
    setStartScreen
    Set thisSession = Nothing
Exit Sub
    
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in saveDimensions of Class clsDimEntry"
    
End Sub
Private Sub doCalcs()
    ' This subroutine calculates various dimensional properties of a tube based on 4 OD (Outside Diameter)
    ' and 8 wall thickness measurements. It determines the maximum, minimum, and mean values for OD and wall thickness,
    ' as well as ovality and siding.

    Dim x As Integer ' Loop counter
    Dim msngWall(1 To 8) As Variant ' Array to store the 8 wall thickness measurements
    
On Error GoTo ErrHandler

    ' --- Calculate OD Max, Min, and Mean ---
    cODMax = CDec(msngOD1) ' Initialize cODMax with the first OD measurement
    cODMin = CDec(msngOD1) ' Initialize cODMin with the first OD measurement

    ' Find the maximum OD value
    If msngOD2 > cODMax Then cODMax = CDec(msngOD2)
    If msngOD3 > cODMax Then cODMax = CDec(msngOD3)
    If msngOD4 > cODMax Then cODMax = CDec(msngOD4)

    ' Find the minimum OD value
    If msngOD2 < cODMin Then cODMin = CDec(msngOD2)
    If msngOD3 < cODMin Then cODMin = CDec(msngOD3)
    If msngOD4 < cODMin Then cODMin = CDec(msngOD4)

    cODMean = (CDec(msngOD1) + CDec(msngOD2) + CDec(msngOD3) + CDec(msngOD4)) / 4 ' Calculate the mean OD value

    

    ' --- Calculate Ovality ---
    cOvality1 = Abs(msngOD1 - msngOD2) ' Calculate the ovality between the first two OD measurements
    cOvality2 = Abs(msngOD3 - msngOD4) ' Calculate the ovality between the last two OD measurements

    ' Determine the maximum ovality
    If cOvality1 > cOvality2 Then
        cOvalityMax = Round(CDec(cOvality1), 2)
    Else
        cOvalityMax = Round(CDec(cOvality2), 2)
    End If

    ' --- Calculate Wall Thickness Max, Min, and Mean ---
    ' Store the wall thickness measurements in an array
    msngWall(1) = CDec(msngWALL1)
    msngWall(2) = CDec(msngWALL2)
    msngWall(3) = CDec(msngWALL3)
    msngWall(4) = CDec(msngWALL4)
    msngWall(5) = CDec(msngWALL5)
    msngWall(6) = CDec(msngWALL6)
    msngWall(7) = CDec(msngWALL7)
    msngWall(8) = CDec(msngWALL8)

    cWallMax = CDec(msngWall(1)) ' Initialize cWallMax with the first wall thickness measurement
    cWallMin = CDec(msngWall(1)) ' Initialize cWallMin with the first wall thickness measurement
    cWallMean = CDec(msngWall(1)) ' Initialize cWallMean with the first wall thickness measurement

    ' Find the maximum and minimum wall thickness values and calculate the sum for the mean
    For x = 2 To 8
        If msngWall(x) > cWallMax Then cWallMax = CDec(msngWall(x))
        If msngWall(x) < cWallMin Then cWallMin = CDec(msngWall(x))
        cWallMean = CDec(cWallMean) + CDec(msngWall(x))
    Next x

    cWallMean = CDec(cWallMean) / 8 ' Calculate the mean wall thickness value

    ' --- Calculate Siding ---
    ' Find the maximum and minimum wall thickness at each end of the tube to calculate siding
    Dim aMax, bMax, aMin, bMin As Variant

    aMax = CDec(msngWall(1)) ' Initialize aMax with the first wall thickness measurement
    aMin = CDec(msngWall(1)) ' Initialize aMin with the first wall thickness measurement
    bMax = CDec(msngWall(5)) ' Initialize bMax with the fifth wall thickness measurement
    bMin = CDec(msngWall(5)) ' Initialize bMin with the fifth wall thickness measurement

    ' Find the maximum and minimum wall thickness values at each end
    For x = 2 To 4
        If msngWall(x) > aMax Then aMax = CDec(msngWall(x))
        If msngWall(x) < aMin Then aMin = CDec(msngWall(x))
        If msngWall(x + 4) > bMax Then bMax = CDec(msngWall(x + 4))
        If msngWall(x + 4) < bMin Then bMin = CDec(msngWall(x + 4))
    Next x

    cSiding1 = CDec(Abs(aMax - aMin)) ' Calculate the siding at one end of the tube
    cSiding2 = CDec(Abs(bMax - bMin)) ' Calculate the siding at the other end of the tube

    ' --- Calculate Gradient (Difference in average walls) ---
    Dim avWall1, avWall2 As Variant
    avWall1 = CDec((aMax + aMin) / 2)
    avWall2 = CDec((bMax + bMin) / 2)
    cGradient = CDec(Abs(avWall1 - avWall2)) ' Calculate the gradient (difference in average walls) as the difference between wall measurement 1 and 2
    
    ' Determine the maximum siding
    If cSiding2 > cSiding1 Then
        cSidingMax = cSiding2
    Else
        cSidingMax = cSiding1
    End If
    
    'set the bow value. Later this can be the real value of bow
    If msngBOW = True Then
        cBowValue = 0
    Else
        cBowValue = 1
    End If
    cWeight = (3.14159 / 4) * (cODMean ^ 2 - (cODMean - (2 * cWallMean)) ^ 2) * cActualLength * 25.4 * 0.0000022

Exit Sub
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in doCalcs Sub of Class clsDimEntry"

End Sub
Sub updateSHwithDims(ByVal rowNum As Long)
    Dim localWB As Workbook
    Dim localWS As Worksheet
On Error GoTo ErrHandler
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")
    doCalcs
    With localWS
        .Cells(rowNum, 27).Value = CDec(cODMean)
        .Cells(rowNum, 28).Value = CDec(cODMax)
        .Cells(rowNum, 29).Value = CDec(cODMin)
        .Cells(rowNum, 30).Value = CDec(cWallMean)
        .Cells(rowNum, 31).Value = CDec(cWallMax)
        .Cells(rowNum, 32).Value = CDec(cWallMin)
        .Cells(rowNum, 33).Value = CDec(cSidingMax)
        .Cells(rowNum, 34).Value = CDec(cOvalityMax)
        .Cells(rowNum, 35).Value = CDec(cGradient)
        .Cells(rowNum, 36).Value = CDec(cActualLength)
        .Cells(rowNum, 38).Value = CDec(cWeight)
    End With
    

    Set localWS = Nothing
    Set localWB = Nothing
Exit Sub
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in updateSHwithDims Sub of Class clsDimEntry"

End Sub

Public Property Get itemNum() As String
    itemNum = mstritemNum
End Property
Public Property Let itemNum(rData As String)
    mstritemNum = rData
End Property
Public Property Get batch() As String
    batch = mstrbatch
End Property
Public Property Let batch(rData As String)
    mstrbatch = rData
End Property
Public Property Get OD1() As Variant
    OD1 = msngOD1
End Property
Public Property Let OD1(rData As Variant)
    msngOD1 = rData
End Property
Public Property Get OD2() As Variant
    OD2 = msngOD2
End Property
Public Property Let OD2(rData As Variant)
    msngOD2 = rData
End Property
Public Property Get OD3() As Variant
    OD3 = msngOD3
End Property
Public Property Let OD3(rData As Variant)
    msngOD3 = rData
End Property
Public Property Get OD4() As Variant
    OD4 = msngOD4
End Property
Public Property Let OD4(rData As Variant)
    msngOD4 = rData
End Property
Public Property Get WALL1() As Variant
    WALL1 = msngWALL1
End Property
Public Property Let WALL1(rData As Variant)
    msngWALL1 = rData
End Property
Public Property Get WALL2() As Variant
    WALL2 = msngWALL2
End Property
Public Property Let WALL2(rData As Variant)
    msngWALL2 = rData
End Property
Public Property Get WALL3() As Variant
    WALL3 = msngWALL3
End Property
Public Property Let WALL3(rData As Variant)
    msngWALL3 = rData
End Property
Public Property Get WALL4() As Variant
    WALL4 = msngWALL4
End Property
Public Property Let WALL4(rData As Variant)
    msngWALL4 = rData
End Property
Public Property Get WALL5() As Variant
    WALL5 = msngWALL5
End Property
Public Property Let WALL5(rData As Variant)
    msngWALL5 = rData
End Property
Public Property Get WALL6() As Variant
    WALL6 = msngWALL6
End Property
Public Property Let WALL6(rData As Variant)
    msngWALL6 = rData
End Property
Public Property Get WALL7() As Variant
    WALL7 = msngWALL7
End Property
Public Property Let WALL7(rData As Variant)
    msngWALL7 = rData
End Property
Public Property Get WALL8() As Variant
    WALL8 = msngWALL8
End Property
Public Property Let WALL8(rData As Variant)
    msngWALL8 = rData
End Property
Public Property Get actualLength() As Variant
    actualLength = cActualLength
End Property
Public Property Let actualLength(rData As Variant)
    cActualLength = rData
End Property
Public Property Get bow() As Variant
    bow = msngBOW
End Property
Public Property Let bow(rData As Variant)
    msngBOW = rData
End Property
Public Property Get accOverride() As Boolean
    accOverride = mbooaccOverride
End Property
Public Property Let accOverride(rData As Boolean)
    mbooaccOverride = rData
End Property

Private Sub Class_Initialize()
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")

End Sub
