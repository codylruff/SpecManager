VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWarpingMainMenu 
   Caption         =   "Warping Menu"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4128
   OleObjectBlob   =   "formWarpingMainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWarpingMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


























Option Explicit

Private Sub cmdSpecManagement_Click()
    Unload Me
    formSpecConfig.Show
End Sub

Private Sub cmdWarperSetup_Click()
    Unload Me
    formWarpDataSheet.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Cancel = True
    End If

End Sub
