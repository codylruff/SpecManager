VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formViewSpecs 
   Caption         =   "Specification Control"
   ClientHeight    =   11868
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9816
   OleObjectBlob   =   "formViewSpecs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formViewSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





















Option Explicit

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdMaterialSearch_Click()
    SpecManager.RestartApp
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectRevision
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdExportPdf_Click()
    GuiCommands.ConsoleBoxToPdf
End Sub

Private Sub cmdSearch_Click()
    Set App.current_spec = App.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub

Private Sub ClearThisForm()
    Dim i As Integer
    Do While cboSelectRevision.ListCount > 0
        cboSelectRevision.RemoveItem 0
    Loop
    ClearForm Me
End Sub
Private Sub PopulateCboSelectRevision()
    Dim rev As Variant
    Dim i As Integer
    Do While cboSelectRevision.ListCount > 0
        cboSelectRevision.RemoveItem 0
    Loop
    With cboSelectRevision
        For Each rev In App.specs
            .AddItem rev
            .value = rev
        Next rev
    End With
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearThisForm
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Logger.Log "--------- End " & Me.Name & " ----------"
End Sub

Sub MaterialSearch()
    SpecManager.RestartApp
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectRevision
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub ExportPdf()
    GuiCommands.ConsoleBoxToPdf_Test
End Sub

Sub Search()
    Set App.current_spec = App.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub
