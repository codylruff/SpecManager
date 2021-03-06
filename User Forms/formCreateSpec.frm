VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateSpec 
   Caption         =   "Specification Control"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   OleObjectBlob   =   "formCreateSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCreateSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Logger.Log "--------- " & Me.Name & " ----------"
    With App
        'Set .current_template = SpecManager.GetTemplate(cboSelectSpecificationType.value)
        '.current_template.SpecType = cboSelectSpecificationType.value
        ' Set App.current_spec = New Specification
        ' .current_spec.SpecType = .current_template.SpecType
        ' .current_spec.Revision = 0#
'        .current_template.Properties.Item(Utils.ConvertToCamelCase( _
'                cboSelectProperty.value)) = txtPropertyValue
        Set .current_spec.Properties = .current_template.Properties
        Set .current_spec.Tolerances = .current_template.Properties
        lblSpecInfo = "Material ID : " & .current_spec.MaterialId & vbNewLine & _
                      "Material Type : " & .current_template.SpecType
    End With
    PopulateCboSelectProperty
    SpecManager.PrintSpecification Me
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdExportPdf_Click()
    MsgBox "Functionality not implemented!"
End Sub

Private Sub cmdSaveChanges_Click()
' Calls method to save a new specification revision x.0)
    Dim ret_val As Long
    ret_val = SpecManager.SaveNewSpecification(App.current_spec)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Private Sub cmdSetProperty_Click()
' This executes a set property command
    SetProperty
End Sub

Private Sub PopulateCboSelectProperty()
    Dim prop As Variant
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    With cboSelectProperty
        For Each prop In App.current_spec.Properties
          .AddItem prop
        Next prop
    End With
    txtPropertyValue.value = vbNullString
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub ExportPdf()
    MsgBox "Functionality not implemented!"
End Sub

Sub SaveChanges()
' Calls method to save a new specification revision x.0)
    Dim ret_val As Long
    ret_val = SpecManager.SaveNewSpecification(App.current_spec)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "Create Spec Fail"
    Else
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "Create Spec Pass"
    End If
End Sub

Sub SetProperty()
' This executes a set property command
    With App.current_spec
        .Properties.Item(cboSelectProperty.value) = txtPropertyValue
    End With
    SpecManager.PrintSpecification Me
End Sub
