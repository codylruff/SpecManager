VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWarpingSearch 
   Caption         =   "Specification Search"
   ClientHeight    =   10248
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9396
   OleObjectBlob   =   "formWarpingSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWarpingSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







































































Option Explicit

Private Sub PopulateCboSelectRevision()
    Dim rev As Variant
    With cboSelectRevision
        For Each rev In App.specs
            .AddItem rev
        Next rev
    End With
End Sub

Private Sub cmdSubmit_Click()
    'App.warp_order
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

Private Sub cmdOptions_Click()
    Unload Me
    GoToMain
End Sub

Private Sub cmdRefresh_Click()
    Set App.current_spec = App.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
