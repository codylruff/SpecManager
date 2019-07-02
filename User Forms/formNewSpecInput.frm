VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewSpecInput 
   Caption         =   "Create New Specification"
   ClientHeight    =   2928
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4476
   OleObjectBlob   =   "formNewSpecInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNewSpecInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Private Sub UserForm_Initialize()
    Logger.Log "--------- " & Me.Name & " ----------"
    PopulateCboSelectSpecType
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub PopulateCboSelectSpecType()
    Dim template_type As Variant
    With cboSelectSpecificationType
        For Each template_type In App.templates
            .AddItem CStr(template_type.SpecType)
        Next template_type
    End With
End Sub

Private Sub cmdContinue_Click()
    If SpecManager.NewSpecificationInput(cboSelectSpecificationType.Value, UCase(Utils.RemoveWhiteSpace(txtSpecName.Value))) <> vbNullString Then
        Unload Me
        formCreateSpec.Show vbModeless
    Else
        MsgBox "Please enter a template type and specification name !"
        Exit Sub
    End If
End Sub

Sub Continue()
    If SpecManager.NewSpecificationInput(cboSelectSpecificationType.Value, UCase(Utils.RemoveWhiteSpace(txtSpecName.Value))) <> vbNullString Then
        Logger.Log "Spec Input Pass"
    Else
        Logger.Log "Spec Input Fail"
    End If
End Sub
