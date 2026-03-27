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
Dim cMsgString As String
Dim cDescription As String
Dim cIsBaked As Boolean
Dim cMatlNum As String
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
    Dim keepGoing As Boolean
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
        z = 1
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
        keepGoing = True
        x = 0
        Do While keepGoing
        ' Loop through the specifications and add them to the output string
            Dim descLen As Long ' Declare a variable to store the length of the description
            Dim Y As Integer ' Declare an integer variable for loop counter
            Dim desc As String ' Declare a string variable to store the description of the specification
            Dim spec As String ' Declare a string variable to store the value of the specification
            Dim spcs As String ' Declare a string variable for formatting

            spcs = vbNullString ' Initialize the formatting string
            Y = 0 ' Initialize the counter to 0

            ' Retrieve the description and value of the specification
            desc = .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MNAME[0," & x & "]").Text
            spec = .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1," & x & "]").Text
            descLen = Len(desc) ' Get the length of the description
            loadClassValues desc, spec ' Call a function to load the class values
 '           Debug.Print desc
            ' Add dots for formatting
            For Y = 0 To 29 - descLen
                spcs = spcs & "."
            Next Y

            
            If x = 9 Then
                .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S").VerticalScrollbar.Position = z * 10
                z = z + 1
                x = -1
            End If
            x = x + 1
            If .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MNAME[0," & x & "]").Text = "______________________________" Then keepGoing = False


            ' --- Handle scrolling to retrieve more specifications ---
 '           If x = 9 Then
 '               z = z + 1 ' Increment the counter to indicate that the end of the current page has been reached
 '           End If
 '           If z = 1 And x = 9 Then
 '               .FindById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S").VerticalScrollbar.Position = 10 ' Scroll down to the next page
 '               x = -1 ' Reset the counter to continue retrieving specifications from the next page
 '           End If
        Loop

        cMsgString = msgString ' Assign the formatted output string to a global variable (cMsgString) - consider using a local variable instead
        getSpecs = msgString ' Return the formatted output string
    End With

    calculateMaxAndMins ' Call a function to calculate maximum and minimum values
    Exit Function
ErrHandler:
        StdErrorHandler "Error " & Err & ": " & Error(Err) & " in getSpecs of clsMaterial"
End Function

Sub loadClassValues(description As String, specification As String)
    Select Case description
        Case "LD Outside Diameter"
            cOutsideDiameter = CSng(specification)
        Case "LD OD Tolerance +"
            cOutsideDiameterPlusTolerance = CSng(specification)
        Case "LD OD Tolerance -"
            cOutsideDiameterMinusTolerance = CSng(specification)
        Case "LD Inside Diameter"
            cInsideDiameter = CSng(specification)
        Case "LD ID Tolerance +"
            cInsideDiameterPlusTolerance = CSng(specification)
        Case "LD ID Tolerance -"
            cInsideDiameterMinusTolerance = CSng(specification)
        Case "LD Wall"
            cWall = CSng(specification)
        Case "LD Wall Tolerance +"
            cWallPlusTolerance = CSng(specification)
        Case "LD Wall Tolerance -"
            cWallMinusTolerance = CSng(specification)
        Case "LD Length"
            cLength = CSng(specification)
        Case "LD Length Tolerance +"
            cLengthPlusTolerance = CSng(specification)
        Case "LD Length Tolerance -"
            cLengthMinusTolerance = CSng(specification)
        Case "HV Bake Level", "TB Bake Level", "DB Bake Level", "RD Bake Level", "SC Bake Level"
            If specification = "UNBAKED" Then
                cIsBaked = False
            Else
                cIsBaked = True
            End If
    End Select
End Sub
Sub displaySpecs()
'    frmSpecs.lblSpecs.Caption = cMsgString
'    frmSpecs.Caption = cDescription
'    frmSpecs.Show

End Sub
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
Public Property Get description() As String
    description = cDescription
End Property
Public Property Get LENGTH() As Single
    LENGTH = cLength
End Property
Public Property Get isBaked() As Boolean
    isBaked = cIsBaked
End Property
Private Sub Class_Initialize()
    Set mySession = SAPSession.CurrentSession ' Initialize the SAP GUI session
End Sub

Private Sub Class_Terminate()
    Set mySession = Nothing
End Sub
