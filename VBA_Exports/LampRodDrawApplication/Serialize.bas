Attribute VB_Name = "Serialize"
Public Function nextTube(furnace As String) As Long
    ' This function generates the next tube ID number based on the last entry in Sheet1.
    ' It uses the current date, furnace name, and a sequential number to create the ID.
    ' Returns the row number where the new ID was written in Sheet1.

    Dim keepGoing As Boolean
    Dim oldSerNum, newSerNum As String
    Dim x, monthTubeCount As Long
    Dim localWB As Workbook
    Dim localWS As Worksheet
On Error GoTo ErrHandler
    Set localWB = ThisWorkbook
    Set localWS = localWB.Sheets("Sheet1")
    localWS.Activate
    x = nextEmptyRow(localWS, "H", 3)
    '                x = 3 ' Start from row 3 in Sheet1
    '                keepGoing = True
                
                    ' --- Find the last tube ID in column H ---
    '                Do While keepGoing
    oldSerNum = Sheet1.Range("H" & x - 1).Value ' Get the value from column H
    '                    x = x + 1 ' Move to the next row
    '                    If localWS.Range("H" & x) = "" Then
    '                        keepGoing = False ' Stop if an empty cell is found
    '                    Else
                            ' Keep looking for the last entry
    '                    End If
    '                Loop
    ' --- End of finding the last ID ---

    monthTubeCount = CLng(right(oldSerNum, 4)) ' Extract the sequential number from the last ID

    ' --- Generate the new tube ID ---
    newSerNum = right(Year(Now), 2) & monthLetter() & "N" & furnLetter(furnace) & newTubeNum(oldSerNum, monthTubeCount)
    '  - Right(Year(Now), 2): Last two digits of the current year
    '  - monthLetter(): Letter code for the current month (A for January, B for February, etc.)
    '  - "N":  A constant letter (purpose unknown)
    '  - furnLetter(furnace): Letter code for the furnace name
    '  - newTubeNum(oldSerNum, monthTubeCount):  Generates the sequential number

    localWS.Range("H" & x).Value = newSerNum ' Write the new ID to Sheet1
    nextTube = x ' Return the row number
    Set localWS = Nothing
    Set localWB = Nothing
Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in nextTube function of Module Serialize"

End Function

Function monthLetter() As String
    ' This function returns a letter code for the current month.
On Error GoTo ErrHandler
    Select Case Month(Now())
        Case 1: monthLetter = "A"
        Case 2: monthLetter = "B"
        Case 3: monthLetter = "C"
        Case 4: monthLetter = "D"
        Case 5: monthLetter = "E"
        Case 6: monthLetter = "F"
        Case 7: monthLetter = "G"
        Case 8: monthLetter = "H"
        Case 9: monthLetter = "J"
        Case 10: monthLetter = "K"
        Case 11: monthLetter = "M"
        Case 12: monthLetter = "N"
    End Select
Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in monthLetter function of Module Serialize"

End Function

Private Function furnLetter(furnName As String) As String
    ' This function returns a letter code for the given furnace name.
On Error GoTo ErrHandler
    Select Case furnName
        Case "NF21": furnLetter = "A"
        Case "NF22": furnLetter = "B"
        Case "NF23": furnLetter = "C"
        Case "NF24": furnLetter = "D"
        Case "NF25": furnLetter = "E"
        Case "NF26": furnLetter = "F"
        Case "NF27": furnLetter = "G"
        Case "NF28": furnLetter = "H"
        Case "NF29": furnLetter = "J"
        Case "NF30": furnLetter = "K"
        Case "NF31": furnLetter = "L"
        Case "NF32": furnLetter = "M"
        Case "NF33": furnLetter = "N"
        Case "NF34": furnLetter = "P"
        Case "NF35": furnLetter = "R"
        Case "NF36": furnLetter = "S"
        Case "NF37": furnLetter = "T"
        Case "NF38": furnLetter = "U"
    End Select
Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in furnLetter function of Module Serialize"

End Function

Private Function newTubeNum(ByVal oldSerNum As String, ByVal monthTubeCount As Long) As String
    ' This function generates the sequential number for the new tube ID.
    ' It checks if the month has changed and resets the counter if necessary.
On Error GoTo ErrHandler
    If monthLetter <> Mid(oldSerNum, 3, 1) Then
        newTubeNum = "0001" ' Reset the counter if the month has changed
    Else
        ' Increment the counter and format it with leading zeros
        If monthTubeCount < 9 Then
            newTubeNum = "000" & monthTubeCount + 1
        ElseIf monthTubeCount < 99 Then
            newTubeNum = "00" & monthTubeCount + 1
        ElseIf monthTubeCount < 999 Then
            newTubeNum = "0" & monthTubeCount + 1
        Else
            newTubeNum = monthTubeCount + 1
        End If
    End If
Exit Function
ErrHandler:

    StdErrorHandler "Error " & Err & ": " & Error(Err) & " in newTubeNum function of Module Serialize"

End Function

