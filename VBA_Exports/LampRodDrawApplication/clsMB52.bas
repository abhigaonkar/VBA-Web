VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMB52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------
' Purpose:  A class module to automate the SAP transaction MB52.
'           This class handles inputting selection criteria, executing the report,
'           and extracting specified data.
'-----------------------------------------------------------------------------------

' Define the class and its public properties.
' Public properties allow you to set the input values from outside the class.
    Public Material
    Public Plant
    Public StorageLocation
    Public Batch
    Public Results

    '-------------------------------------------------------------------------------
    ' Method: RunReport
    ' Purpose: Executes the MB52 report based on the properties set.
    ' Returns: A dictionary object containing the extracted data.
    '-------------------------------------------------------------------------------
    Public Function RunReport()
        ' Get the SAP GUI session. If it fails, exit the function.
        On Error Resume Next ' Use error handling for SAP GUI interactions.
        Dim thisSession As GuiSession
        Set thisSession = SAPSession.CurrentSession  ' Initialize SAP GUI session
        setStartScreen ' Call a function to set the initial screen
        ' Start the transaction code "MB52"
        thisSession.StartTransaction "MB52"

        ' --- Populate the selection screen ---

        ' Set Material, Plant, Storage Location, and Batch from the class properties.
        ' Using `session.findById` to locate the input fields by their IDs.
        ' The `text` property is used to set the value.
        thisSession.FindById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = Me.Material
        thisSession.FindById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = Me.Plant
        thisSession.FindById("wnd[0]/usr/ctxtS_LGORT-LOW").Text = Me.StorageLocation
        thisSession.FindById("wnd[0]/usr/ctxtS_CHARG-LOW").Text = Me.Batch

        ' --- Set additional options ---

        ' Check the "No zero stock lines" checkbox.
        ' The `selected` property is used for checkboxes.
        thisSession.FindById("wnd[0]/usr/chkP_KZVBR").Selected = True

        ' Select the "Non-Hierarchical Representation" radio button.
        ' The `selected` property is used for radio buttons.
        thisSession.FindById("wnd[0]/usr/radP_S_ALV").Selected = True

        ' Set the layout.
        ' The ID for the layout field is `ctxtP_VARIAN`.
        thisSession.FindById("wnd[0]/usr/ctxtP_VARIAN").Text = "/parkinsonj"

        ' --- Execute the report ---

        ' Click the "Execute" button. The button ID is `btn[8]`.
        thisSession.FindById("wnd[0]/tbar[1]/btn[8]").Press

        ' --- Extract the data from the report results screen ---

        ' Create a dictionary to hold the results.
        Dim resultsDict
        Set resultsDict = CreateObject("Scripting.Dictionary")

        ' Target the main user area of the window.
        ' This is a standard ID for the user-defined area of a screen.
        Dim userArea
        Set userArea = thisSession.FindById("wnd[0]/usr")
        
        ' Get all the components within the user area.
        Dim allComponents
        Set allComponents = userArea.Children

        ' Create collections to store the extracted data.
        Dim records
        Set records = CreateObject("Scripting.Dictionary")
        Dim recordCounter: recordCounter = 0

        ' We'll need to use a temporary dictionary to build each record as we find the fields.
        Dim currentRecord
        Set currentRecord = CreateObject("Scripting.Dictionary")

        ' Loop through each component in the collection.
        Dim i
        For i = 0 To allComponents.Count - 1
            Dim currentComponent As GuiComponent
            Set currentComponent = allComponents.Item(i)

            ' Check the type of the component to find the data.
            ' A GuiTextField is a good candidate for holding data.
            ' The ID or name of the component is the key to identifying it.
            
            ' This is a placeholder logic based on common field IDs in SAP.
            If currentComponent.Type = "GuiLabel" Then
                If InStr(1, currentComponent.ID, "LGORT") > 0 Then
                    currentRecord.Add "StorageLocation", currentComponent.Text
                ElseIf InStr(1, currentComponent.ID, "CHARG") > 0 Then
                    currentRecord.Add "Batch", currentComponent.Text
                ElseIf InStr(1, currentComponent.ID, "LABST") > 0 Then
                    currentRecord.Add "Unrestricted", currentComponent.Text
                End If
            End If

            ' This is the most complex part: identifying when a full record is collected.
            ' You'll have to find a consistent pattern. For example, if there's a
            ' blank line or a specific label that indicates the end of a record, you can
            ' use that. Without a specific pattern, we can't guarantee a clean split.
            ' The simplest approach is to check if we've found all the required fields.
            
            If currentRecord.Count = 3 Then
                ' We've found all three fields for a record.
                records.Add CStr(recordCounter), currentRecord
                recordCounter = recordCounter + 1
                
                ' Reset for the next record.
                Set currentRecord = CreateObject("Scripting.Dictionary")
            End If
            Set currentComponent = Nothing
            
        Next

        ' Add the total number of records and the extracted records to the main dictionary.
        resultsDict.Add "TotalRecords", records.Count
        resultsDict.Add "Records", records

        Set Me.Results = resultsDict
        Set RunReport = resultsDict
    On Error GoTo 0
    End Function

