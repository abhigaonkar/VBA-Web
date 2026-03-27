Attribute VB_Name = "modTests"
Option Explicit

Sub testHULabelPrint()
    Dim testMe As New clsHU
    Dim batcchArray(3) As Variant
    batcchArray(0) = "25jnqa055"
    batcchArray(1) = "25jnqa059"
    batcchArray(2) = "25jnqa060"
    batcchArray(3) = "25jnqa061"
    testMe.batchArray = batcchArray
    testMe.itemNum = 156292
    testMe.humo4Gr
    Set testMe = Nothing
End Sub

Sub testReadOvenFile()
    Dim thisTest As New clsPalletCardData
    thisTest.cardNum = 66555
    thisTest.ReadFilteredDataIntoArray
    Set thisTest = Nothing
End Sub
Sub testUOM()
    Dim thisTest As New clsCOR3
    thisTest.getDetails "4496414"
    Debug.Print thisTest.UOM
    Set thisTest = Nothing
End Sub
Sub testMat()
    Dim thisMat As New clsMaterial
    thisMat.getDescription "110322"
    Set thisMat = Nothing
End Sub
Sub testMRP()
    Dim thisOrder As New clsCOR3
    thisOrder.getDetails "4495393"
    Debug.Print thisOrder.POType
    Set thisOrder = Nothing
End Sub
Sub testCtnMake()
    Dim makeHampers As New clsNewCOR6
    Dim hamperNum As String
    Dim hamperQty As Long
    hamperNum = "134976"
    hamperQty = 3
    With makeHampers
        .procOrdNum = getCtnProcOrdNums(hamperNum)(0)
        .yield = hamperQty
        .applyMatl = False
        .saveIt
    End With
    Set makeHampers = Nothing

End Sub
