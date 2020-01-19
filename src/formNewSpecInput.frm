VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewSpecInput 
   Caption         =   "Create New Specification"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
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
    Dim spec_type_selection As String
    Dim machine_id_selection As String
    Dim material_id As String
    spec_type_selection = cboSelectSpecificationType.value
    machine_id_selection = txtMachineId.value
'    If selection = "Weaving RBA" Then
'        ' For the weaving rba a base file must be selected to load.
'        Unload Me
'        DocumentParser.LoadNewRBA
'        GuiCommands.GoToMain
'        Exit Sub
'    End If
    material_id = UCase(Utils.RemoveWhiteSpace(txtSpecName.value))
    If SpecManager.NewSpecificationInput(spec_type_selection, material_id, machine_id_selection) <> nullstr Then
        Unload Me
        formCreateSpec.show vbModeless
    Else
        PromptHandler.Error "Please enter a template type, specification name, and machine Id !"
        Exit Sub
    End If
End Sub

Sub Continue()
    If SpecManager.NewSpecificationInput(cboSelectSpecificationType.value, UCase(Utils.RemoveWhiteSpace(txtSpecName.value)), txtMachineId.value) <> nullstr Then
        Logger.Log "Spec Input Pass"
    Else
        Logger.Log "Spec Input Fail"
    End If
End Sub
