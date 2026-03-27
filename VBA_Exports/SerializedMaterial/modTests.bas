Attribute VB_Name = "modTests"
Option Explicit
Sub enterQMdata()
    Dim x As Long
    x = 6
    
    Dim theseDims As New clsDimEntry
    theseDims.saveDimensions x  ' Call a method to save the dimensions in SAP
    Set theseDims = Nothing ' Release the clsDimEntry object
    
End Sub

Sub testBatchNum()
    Dim thisOrder As New clsCOOISPI
    Debug.Print thisOrder.getProcOrdNum("2100094", False)
    Debug.Print thisOrder.batch
    Set thisOrder = Nothing
End Sub
Sub testMB52()
    Dim thisInv As New clsMB52
    thisInv.getInventory "2102336", "0100"
    Debug.Print thisInv.totalAtLocation
    Set thisInv = Nothing
End Sub
Sub testhumo4Gr()
       Dim thisInvCheck As New clsMB52
        Dim y As Integer
        thisInvCheck.getInventory "2307813", "0300"
                     
    'when hamper full do HU4GR to complete it
    Dim thisHU As New clsHU
    Dim p As Integer
    Dim newArray() As Variant
    ReDim newArray(UBound(thisInvCheck.inventoryArray))
    For p = 0 To UBound(thisInvCheck.inventoryArray)
        newArray(p) = thisInvCheck.inventoryArray(p, 1)
    Next p
    thisHU.itemNum = "2307813"
    thisHU.batchArray = newArray
    thisHU.humo4Gr
    Set thisHU = Nothing
End Sub
Sub testVL32N()
    Dim thisInvCheck As New clsMB52
    Dim y As Integer
    thisInvCheck.getInventory "2307813", "0300"
    vl32N thisInvCheck.inventoryArray
    
    Set thisInvCheck = Nothing
End Sub

