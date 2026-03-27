VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tubeIDAlert 
   Caption         =   "Tube ID Display"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20670
   OleObjectBlob   =   "tubeIDAlert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tubeIDAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Me.Hide
    UserForm1.Show
End Sub
