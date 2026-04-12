Attribute VB_Name = "modTests"
Option Explicit

Sub testHU()
Dim thisHU As New clsHU
thisHU.huGI "4484911", "1000000159", "25JNT0018"
Set thisHU = Nothing
End Sub
Sub testNumOrders()
Dim thisTest As New clsCOOISPI
Debug.Print thisTest.numOrders("155446")
Set thisTest = Nothing
End Sub
