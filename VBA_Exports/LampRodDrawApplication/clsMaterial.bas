VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim mySession As GuiSession ' Declare a variable to hold the SAP GUI session
Dim cOutsideDiameter As Single
Dim cOutsideDiameterPlusTolerance As Single
Dim cOutsideDiameterMinusTolerance As Single
Dim cOutsideDiameterMax As Single
Dim cOutsideDiameterMin As Single
Dim cInsideDiameter As Single
Dim cInsideDiameterPlusTolerance As Single
Dim cInsideDiameterMinusTolerance As Single
Dim cInsideDiameterMax As Single
Dim cInsideDiameterMin As Single
Dim cWall As Single
Dim cWallPlusTolerance As Single
Dim cWallMinusTolerance As Single
Dim cWallMax As Single
Dim cWallMin As Single
Dim cLength As Single
Dim cLengthPlusTolerance As Single
Dim cLengthMinusTolerance As Single
Dim cLengthMax As Single
Dim cLengthMin As Single
Dim cSiding As Variant
Dim cOvality As Variant
Dim cBowLimit As Variant
Dim cMsgString As String
Dim cDescription As String
Dim cMatlNum As String
Dim cSpecArray() As Variant
Dim cMatMemoExists As Boolean
Dim cMatMemo As String

Function getDescription(materialNumber As String) As String
    ' This function retrieves the description of a material (item) from SAP using the MM03 transaction.
    ' It takes the material number as input and returns the description as a string.
    ' If the material number is invalid, it returns "Bad Item Number".

On Error GoTo ErrHandler
    cMatlNum = materialNumber ' Assign the input material number to a global variable (cMatlNum) - consider using a local variable instead

    If Not setStartScreen Then Exit Function ' Call a function to set the initial screen


    With mySession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "mm03" ' Enter the transaction code MM03 in the command field
        .FindById("wnd[0]").SendVKey 0 ' Press Enter to execute the transaction

        .FindById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = materialNumber ' Enter the material number in the input field
        .FindById("wnd[0]").SendVKey 0 ' Press Enter to display the material data
        
        If SAPStatus.CurrentStatus = "The material " & materialNumber & " does not exist or is not activated" Then
            MsgBox SAPStatus.CurrentStatus, vbOKOnly
            Exit Function
        End If
        
        .FindById("wnd[1]/tbar[0]/btn[20]").Press 'select all
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        If SAPStatus.IsThereAModalWindow Then
            If SAPStatus.SAPWindowName = "Organizational Levels" Then
                .FindById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "q105"
                .FindById("wnd[1]/tbar[0]/btn[0]").Press
            End If
        End If
        
        ' --- Check if the material number is valid ---
        If SAPStatus.IsThereAModalWindow Then ' Check if a modal window appears
            MsgBox SAPStatus.CurrentStatus, vbOKOnly ' If a modal window appears, display the error message
            getDescription = "Bad Item Number" ' Set the return value to "Bad Item Number"
            .FindById("wnd[0]").SendVKey 3 ' Press F3 to go back
            Exit Function ' Exit the function
        End If
        cDescription = .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").Text ' Retrieve the material description from the text field

        ' --- Navigate to the material description ---
 '       .FindById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(2).Selected = True ' Select the second row in the table control (which usually contains the description)
 '       .FindById("wnd[1]/tbar[0]/btn[7]").Press ' Press the button to display the long text (description)
        If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15").Text = " MRP 4" Then
            .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15").Select
            If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/chkNOTE_EXIST").Selected Then
                cMatMemoExists = True
            Else
                cMatMemoExists = False
            End If
        Else
            .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP16").Select
            If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/chkNOTE_EXIST").Selected Then
                cMatMemoExists = True
            Else
                cMatMemoExists = False
            End If
        End If

        getSpecs
        .FindById("wnd[0]").SendVKey 3 ' Press F3 to go back
        .FindById("wnd[0]").SendVKey 3 ' Press F3 again to go back to the initial screen
    End With

    getDescription = cDescription ' Return the material description
    Exit Function
ErrHandler:
        StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getDescription of clsMaterial"
End Function
Private Function getSpecs() As String
    ' This function retrieves the specifications (dimensions and other characteristics) of a material from SAP.
    ' It uses the MM03 transaction to access the material data and extracts the relevant information from the "Classification" tab.
    ' The function returns a formatted string containing the material description and its specifications.

    Dim x As Integer ' Declare an integer variable for loop counter
    Dim z As Integer ' Declare an integer variable to track if all specifications have been retrieved
On Error GoTo ErrHandler
    z = 0 ' Initialize z to 0

    With mySession

        ' --- Navigate to the "Classification" tab ---
        .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03").Select
        PauseMe 0.5, True
        
        
        If SAPStatus.IsThereAModalWindow And SAPStatus.SAPWindowName = "Class Type 5 Entries" Then ' Check if a modal window appears
            ' Find the correct screen by searching for the "Material class" label
            For x = 3 To 5
                If .FindById("wnd[1]/usr/lbl[1," & x & "]").Text = "Material Class" Then
                    .FindById("wnd[1]/usr/lbl[42," & x & "]").SetFocus ' Set focus to the corresponding field
                    .FindById("wnd[1]").SendVKey 2 ' Press Enter to continue
                    x = 5 ' Exit the loop
                End If
            Next x
        Else
            .FindById("wnd[0]/usr/btn%#AUTOTEXT004").Press
            .FindById("wnd[1]/usr/lbl[42,3]").SetFocus
            .FindById("wnd[1]/usr/lbl[42,3]").CaretPosition = 2
            .FindById("wnd[1]").SendVKey 2
        End If
'        .FindById("wnd[0]/usr/btn%#AUTOTEXT004").Press
'
        cLength = .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,9]").Text

        ' --- Navigate to the specifications ---
'        .FindById("wnd[0]/usr/btn%#AUTOTEXT004").Press ' Press the button to display the classification view
'        .FindById("wnd[1]/usr/lbl[1,3]").CaretPosition = 25 ' Set the caret position in the label
'        .FindById("wnd[1]").SendVKey 2 ' Press Enter to continue

        ' Find the "SPECIFICATION" entry in the table control
        Dim notFound As Boolean
        notFound = True
        Dim s As Integer
        s = 0
        Do While notFound
            If right(.FindById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0," & s & "]").Text, 13) = "SPECIFICATION" Then
                .FindById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0," & s & "]").SetFocus ' Set focus to the "SPECIFICATION" entry
                .FindById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0," & s & "]").CaretPosition = 17 ' Set the caret position
                .FindById("wnd[0]").SendVKey 2 ' Press Enter to display the specifications
                notFound = False ' Exit the loop
            Else
                s = s + 1 ' Increment the counter to check the next row
            End If
        Loop

        ' --- Extract the material description and specifications ---
        Dim msgString As String ' Declare a string variable to store the formatted output
        cDescription = .FindById("wnd[0]/usr/subOBJEKT:SAPLCBCM:0100/txtRMCBC-OBTXT").Text ' Retrieve the material description
        msgString = cDescription & vbCrLf ' Add the description to the output string

        ' Loop through the specifications and add them to the output string
        For x = 0 To 9
            Dim descLen As Long ' Declare a variable to store the length of the description
            Dim y As Integer ' Declare an integer variable for loop counter
            Dim desc As String ' Declare a string variable to store the description of the specification
            Dim spec As Variant ' Declare a string variable to store the value of the specification
            Dim spcs As String ' Declare a string variable for formatting

            spcs = vbNullString ' Initialize the formatting string
            y = 0 ' Initialize the counter to 0

            ' Retrieve the description and value of the specification
            desc = .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MNAME[0," & x & "]").Text
            spec = .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1," & x & "]").Text
            descLen = Len(desc) ' Get the length of the description

            loadClassValues desc, spec ' Call a function to load the class values

            ' Add dots for formatting
            For y = 0 To 29 - descLen
                spcs = spcs & "."
            Next y


            ' --- Handle scrolling to retrieve more specifications ---
            If x = 9 Then
                z = z + 1 ' Increment the counter to indicate that the end of the current page has been reached
            End If
            If z = 1 And x = 9 Then
                .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S").VerticalScrollbar.Position = 10 ' Scroll down to the next page
                x = -1 ' Reset the counter to continue retrieving specifications from the next page
            End If
        Next x

        cMsgString = msgString ' Assign the formatted output string to a global variable (cMsgString) - consider using a local variable instead
        getSpecs = msgString ' Return the formatted output string
    End With

    calculateMaxAndMins ' Call a function to calculate maximum and minimum values
    Exit Function
ErrHandler:
    If Err.Number = 619 Then
        MsgBox "The item number for this process order doesn't have specifications. Please check the process order number.", vbCritical
        Exit Function
    End If
        StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getSpecs of clsMaterial"
End Function

Sub loadClassValues(Description As String, specification As Variant)

    
Select Case Description
    ' Group 1: Outside Diameter (OD/Outer Diameter)
    Case "LD Outside Diameter", "HV Outer Diameter", "TB Outside Diameter", "RD Outer Diameter"
        cOutsideDiameter = CDec(specification)

    ' Group 2: Outside Diameter Plus Tolerance
    Case "LD OD Tolerance +", "HV OD Tolerance +", "TB OD Tolerance +", "RD OD Tolerance +"
        cOutsideDiameterPlusTolerance = CDec(specification)

    ' Group 3: Outside Diameter Minus Tolerance
    Case "LD OD Tolerance -", "HV OD Tolerance -", "TB OD Tolerance -", "RD OD Tolerance -"
        cOutsideDiameterMinusTolerance = CDec(specification)

    ' Group 4: Inside Diameter (ID)
    Case "LD Inside Diameter", "HV Inside Diameter", "TB Inside Diameter", "RD Inside Diameter"
        cInsideDiameter = CDec(specification)

    ' Group 5: Inside Diameter Plus Tolerance
    Case "LD ID Tolerance +", "HV ID Tolerance +", "TB ID Tolerance +", "RD ID Tolerance +"
        cInsideDiameterPlusTolerance = CDec(specification)

    ' Group 6: Inside Diameter Minus Tolerance
    Case "LD ID Tolerance -", "HV ID Tolerance -", "TB ID Tolerance -", "RD ID Tolerance -"
        cInsideDiameterMinusTolerance = CDec(specification)

    ' Group 7: Wall Thickness
    Case "LD Wall", "HV Wall", "TB Wall", "RD Wall"
        cWall = CDec(specification)

    ' Group 8: Wall Plus Tolerance
    Case "LD Wall Tolerance +", "HV Wall Tolerance +", "TB Wall Tolerance +", "RD Wall Tolerance +"
        cWallPlusTolerance = CDec(specification)

    ' Group 9: Wall Minus Tolerance
    Case "LD Wall Tolerance -", "HV Wall Tolerance -", "TB Wall Tolerance -", "RD Wall Tolerance -"
        cWallMinusTolerance = CDec(specification)

    ' Group 10: Length
    Case "LD Length", "HV Length", "TB Length", "RD Length"
        cLength = CDec(specification)

    ' Group 11: Length Plus Tolerance
    Case "LD Length Tolerance +", "HV Length Tolerance +", "TB Length Tolerance +", "RD Length Tolerance +"
        cLengthPlusTolerance = CDec(specification)

    ' Group 12: Length Minus Tolerance
    Case "LD Length Tolerance -", "HV Length Tolerance -", "TB Length Tolerance -", "RD Length Tolerance -"
        cLengthMinusTolerance = CDec(specification)
        
    ' Group 13: TB Specific String Values (no CDec conversion)
    Case "TB Maximum Siding"
        cSiding = specification
    Case "TB Maximum Ovality"
        cOvality = specification
    Case "TB Maximum Bow Deflection"
        cBowLimit = specification
        
End Select
    
    
    
End Sub
Function getSpecsAsString() As String
    Dim result As String
    Dim arrayResult() As Variant
    ReDim arrayResult(14, 1)
   ' getSpecs
    arrayResult(0, 0) = "OD"
    arrayResult(0, 1) = cOutsideDiameter
    arrayResult(1, 0) = "OD+"
    arrayResult(1, 1) = cOutsideDiameterPlusTolerance
    arrayResult(2, 0) = "OD-"
    arrayResult(2, 1) = cOutsideDiameterMinusTolerance
    arrayResult(3, 0) = "ID"
    arrayResult(3, 1) = cInsideDiameter
    arrayResult(4, 0) = "ID+"
    arrayResult(4, 1) = cInsideDiameterPlusTolerance
    arrayResult(5, 0) = "ID-"
    arrayResult(5, 1) = cInsideDiameterMinusTolerance
    arrayResult(6, 0) = "Wall"
    arrayResult(6, 1) = cWall
    arrayResult(7, 0) = "Wall+"
    arrayResult(7, 1) = cWallPlusTolerance
    arrayResult(8, 0) = "Wall-"
    arrayResult(8, 1) = cWallMinusTolerance
    arrayResult(9, 0) = "Length"
    arrayResult(9, 1) = cLength
    arrayResult(10, 0) = "Length+"
    arrayResult(10, 1) = cLengthPlusTolerance
    arrayResult(11, 0) = "Length-"
    arrayResult(11, 1) = cLengthMinusTolerance
    arrayResult(12, 0) = "Ovality-"
    arrayResult(12, 1) = cOvality
    arrayResult(13, 0) = "Siding-"
    arrayResult(13, 1) = cSiding
    arrayResult(14, 0) = "BowLimit-"
    arrayResult(14, 1) = cBowLimit
    cSpecArray = arrayResult
    result = "OD" & vbTab & vbTab & cOutsideDiameter
    If cOutsideDiameterPlusTolerance > 0 Then result = result & " (" & cOutsideDiameterMin & " - " & cOutsideDiameterMax & ")"
    result = result & vbCrLf
    result = result & "OD+ " & vbTab & vbTab & cOutsideDiameterPlusTolerance
    result = result & vbCrLf
    result = result & "OD- " & vbTab & vbTab & cOutsideDiameterMinusTolerance
    result = result & vbCrLf
    result = result & "ID " & vbTab & vbTab & cInsideDiameter
    result = result & vbCrLf
    result = result & "ID+" & vbTab & vbTab & cInsideDiameterPlusTolerance
    result = result & vbCrLf
    result = result & "ID-" & vbTab & vbTab & cInsideDiameterMinusTolerance
    result = result & vbCrLf
    result = result & "WALL" & vbTab & vbTab & cWall
    If cWallPlusTolerance > 0 Then result = result & " (" & cWallMin & " - " & cWallMax & ")"
    result = result & vbCrLf
    result = result & "Wall+" & vbTab & vbTab & cWallPlusTolerance
    result = result & vbCrLf
    result = result & "Wall-" & vbTab & vbTab & cWallMinusTolerance
    result = result & vbCrLf
    result = result & "Length" & vbTab & vbTab & cLength
    result = result & vbCrLf
    result = result & "Length+" & vbTab & cLengthPlusTolerance
    result = result & vbCrLf
    result = result & "Length-" & vbTab & cLengthMinusTolerance
    result = result & vbCrLf
    result = result & "Siding-" & vbTab & cSiding
    result = result & vbCrLf
    result = result & "Ovality-" & vbTab & cOvality
    result = result & vbCrLf
    result = result & "BowLimit-" & vbTab & cBowLimit
    getSpecsAsString = result
End Function

Sub calculateMaxAndMins()
    ' This subroutine calculates the maximum and minimum values for inside diameter, outside diameter, and wall thickness
    ' based on the nominal values and their corresponding tolerances.
    ' It uses global variables to store the input values and calculated results.

    ' --- Calculate inside diameter max and min ---
    If cInsideDiameterPlusTolerance <> 0 And cInsideDiameterMinusTolerance <> 0 Then ' Check if tolerances are defined
        cInsideDiameterMax = cInsideDiameter + cInsideDiameterPlusTolerance ' Calculate the maximum inside diameter
        cInsideDiameterMin = cInsideDiameter - cInsideDiameterMinusTolerance ' Calculate the minimum inside diameter
    End If

    ' --- Calculate outside diameter max and min ---
    If cOutsideDiameterPlusTolerance <> 0 And cOutsideDiameterMinusTolerance <> 0 Then ' Check if tolerances are defined
        cOutsideDiameterMax = cOutsideDiameter + cOutsideDiameterPlusTolerance ' Calculate the maximum outside diameter
        cOutsideDiameterMin = cOutsideDiameter - cOutsideDiameterMinusTolerance ' Calculate the minimum outside diameter
    End If

    ' --- Calculate wall thickness max and min ---
    If cWallPlusTolerance <> 0 And cWallMinusTolerance <> 0 Then ' Check if tolerances are defined
        cWallMax = cWall + cWallPlusTolerance ' Calculate the maximum wall thickness
        cWallMin = cWall - cWallMinusTolerance ' Calculate the minimum wall thickness
    End If
    ' Good place to save length
End Sub
Sub showMemo(ByVal materialNumber As Long)
    cMatlNum = materialNumber ' Assign the input material number to a global variable (cMatlNum) - consider using a local variable instead

    If Not setStartScreen Then Exit Sub      ' Call a function to set the initial screen


    With mySession
        .FindById("wnd[0]/tbar[0]/okcd").Text = "mm03" ' Enter the transaction code MM03 in the command field
        .FindById("wnd[0]").SendVKey 0 ' Press Enter to execute the transaction

        .FindById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = materialNumber ' Enter the material number in the input field
        .FindById("wnd[0]").SendVKey 0 ' Press Enter to display the material data
        
        If SAPStatus.CurrentStatus = "The material " & materialNumber & " does not exist or is not activated" Then
            MsgBox SAPStatus.CurrentStatus, vbOKOnly
            Exit Sub
        End If
        
        .FindById("wnd[1]/tbar[0]/btn[20]").Press 'select all
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        .FindById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "q105"
        .FindById("wnd[1]/usr/ctxtRMMG1-WERKS").CaretPosition = 4
        .FindById("wnd[1]/tbar[0]/btn[0]").Press
        
        ' --- Check if the material number is valid ---
        If SAPStatus.IsThereAModalWindow Then ' Check if a modal window appears
            MsgBox SAPStatus.CurrentStatus, vbOKOnly ' If a modal window appears, display the error message
            .FindById("wnd[0]").SendVKey 3 ' Press F3 to go back
            Exit Sub      ' Exit the function
        End If
        If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15").Text = " MRP 4" Then
            .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15").Select
            If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/chkNOTE_EXIST").Selected Then
                .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/btnMARC_NOTE").Press
                Application.WindowState = xlMinimized
            End If
        Else
            .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP16").Select
            If .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/chkNOTE_EXIST").Selected Then
                .FindById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/btnMARC_NOTE").Press
                Application.WindowState = xlMinimized
            End If

        End If

    End With
End Sub
Function IDSpecPass(MeasuredODMax As Single, MeasuredODMin As Single, MeasuredWallMax As Single, MeasuredWallMin As Single) As Boolean
Dim IDMax As Single
Dim IDMin As Single
Dim result As Boolean
    If cInsideDiameterPlusTolerance = 0 Or cInsideDiameterMinusTolerance = 0 Then
        MsgBox "This isn't an ID spec tube.", vbOKOnly
        IDSpecPass = False
        Exit Function
    End If
    result = True
    IDMax = MeasuredODMax - 2 * MeasuredWallMin
    IDMin = MeasuredODMin - 2 * MeasuredWallMax
    If IDMax > cInsideDiameterMax Then result = False
    If IDMin < cInsideDiameterMin Then result = False
    IDSpecPass = result
End Function
Public Property Get IDPlusTolerance() As Single
    IDPlusTolerance = cInsideDiameterPlusTolerance
End Property
Public Property Get IDMinusTolerance() As Single
    IDMinusTolerance = cInsideDiameterMinusTolerance
End Property
Public Property Get IDCheck() As Boolean
    IDCheck = False
    If cInsideDiameterPlusTolerance <> 0 And cInsideDiameterMinusTolerance <> 0 Then IDCheck = True
End Property
Public Property Get materialMemoExists() As Boolean
    materialMemoExists = cMatMemoExists
End Property
Public Property Get specArray() As Variant
    specArray = cSpecArray
End Property
Public Property Get Description() As String
    Description = cDescription
End Property
Public Property Get LENGTH() As Single
    LENGTH = cLength
End Property
Public Property Get ODMax() As Variant
    ODMax = cOutsideDiameterMax
End Property
Public Property Get ODMin() As Variant
    ODMin = cOutsideDiameterMin
End Property
Public Property Get wallMax() As Variant
    wallMax = cWallMax
End Property
Public Property Get wallMin() As Variant
    wallMin = cWallMin
End Property
Private Sub Class_Initialize()
    Set mySession = SAPSession.CurrentSession ' Initialize the SAP GUI session
End Sub

Private Sub Class_Terminate()
    Set mySession = Nothing
End Sub
