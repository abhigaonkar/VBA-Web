VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPalletCardData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrCardNum As String
Private mstrItemNumber As String
Private mstrDescription As String
Private msngBakeHours As Variant
Private mvarBatchNumbers As Variant
Private cFirstRow As Long
Private cLastRow As Long

Sub ReadFilteredDataIntoArray()
    ' This subroutine reads filtered data from a specific worksheet ("Sheet1") in an external workbook ("BakeTrial1.xlsm")
    ' and writes it to another worksheet ("FS") in the current workbook ("2024 SAP BETAS.xlsm").
    ' It filters the data based on the pallet card number (mstrCardNum) and extracts relevant information into an array.

    ' --- Declare variables ---
    Dim wb, ovenWB As Workbook ' Workbook objects
    Dim ws, fsws As Worksheet ' Worksheet objects
    Dim arr As Variant ' Array to store the filtered data
    Dim x As Long  ' Loop counters
    Dim rng As Range ' Range object
    Dim pFilePath As String
    If Not productionMode Then
        pFilePath = "C:\Users\PARKINSONJ\OneDrive - Momentive Technologies\Documents\Development Files\Archives\SerializedTubeLog.xlsx"
    Else
        pFilePath = "\\US-FSP01.prod.momentivetech.com\HEB\Common\Serialization\Archives\SerializedTubeLog.xlsx"
    End If
    ' --- Open the external workbook ---
    If IsWorkBookOpen(pFilePath) Then ' Check if the workbook is already open
        'MsgBox "Workbook is already open." ' (Optional) Display a message if the workbook is already open
    Else
        Workbooks.Open filename:=pFilePath, ReadOnly:=True  ' Open the workbook if it's not already open
    End If

    ' --- Set the worksheet to work with ---
    Set wb = Workbooks("SerializedTubeLog.xlsx") ' Set the workbook object
    Set ws = wb.Sheets("Sheet1") ' Set the worksheet object

    ' --- Set the worksheet in the current workbook ---
    Set ovenWB = ThisWorkbook ' Set the current workbook object
    Set fsws = ovenWB.Sheets("Sheet1") ' Set the worksheet object

    ' --- Find the next empty row in the "FS" worksheet ---
'    cFirstRow = nextEmptyRow(fsws, "B") ' Call a function to find the next empty row (not defined here)
    x = 0 ' Assign the next empty row number to x
    Dim dataRange As Range
    Dim area As Range
    Dim cell As Range
    ' Step 1: Set the range.
    ' We use the CurrentRegion property of a known cell within the data.
    ' A1 is a common choice, assuming your data starts there.
    ' This automatically selects the entire contiguous table.
    Set dataRange = Range("A1").CurrentRegion
    dataRange.AutoFilter Field:=24, Criteria1:="FALSE"
    dataRange.AutoFilter Field:=3, Criteria1:=mstrCardNum
    ReDim arr(16)
    ' --- Check if there are any visible rows after filtering ---
    If ws.AutoFilterMode And ws.FilterMode Then
        ' --- Set the range to the filtered data ---
        Set rng = ws.AutoFilter.Range.Offset(1, 0).Resize(ws.AutoFilter.Range.Rows.Count - 1, ws.AutoFilter.Range.Columns.Count).SpecialCells(xlCellTypeVisible)
        ' --- Read the filtered data into an array ---
        For Each area In rng.Areas
            For Each cell In area.Rows ' Or loop through cell in area.Cells
                arr(x) = cell.Value2(1, 8)
                x = x + 1
            Next cell
        Next area
    Else
        MsgBox "No filtered data found." ' Display a message if no filtered data is found
    End If

    mvarBatchNumbers = removeEmptyElementsFromArray(arr)  ' Assign the array to a global variable (mvarBatchNumbers)
    wb.Close SaveChanges:=False

End Sub

Public Property Get cardNum() As String
    cardNum = mstrCardNum
End Property
Public Property Let cardNum(rData As String)
    mstrCardNum = rData
End Property
Public Property Get ItemNumber() As String
    ItemNumber = mstrItemNumber
End Property
Public Property Let ItemNumber(rData As String)
    mstrItemNumber = rData
End Property
Public Property Get description() As String
    description = mstrDescription
End Property
Public Property Let description(rData As String)
    mstrDescription = rData
End Property
Public Property Get bakeHours() As Variant
    bakeHours = msngBakeHours
End Property
Public Property Let bakeHours(rData As Variant)
    msngBakeHours = rData
End Property
Public Property Get batchNumbers() As Variant
    If IsObject(mvarBatchNumbers) Then
        Set batchNumbers = mvarBatchNumbers
    Else
        batchNumbers = mvarBatchNumbers
    End If
End Property
Public Property Let batchNumbers(rData As Variant)
    mvarBatchNumbers = rData
End Property
Public Property Set batchNumbers(rData As Variant)
    Set mvarBatchNumbers = rData
End Property
Public Property Get firstRow() As Long
    firstRow = cFirstRow
End Property
Public Property Get lastRow() As Long
    lastRow = cLastRow
End Property
