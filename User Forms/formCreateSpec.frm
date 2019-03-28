VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateSpec 
   Caption         =   "Specification Control"
   ClientHeight    =   10548
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9732
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
    With manager
        'Set .current_template = SpecManager.GetTemplate(cboSelectSpecificationType.value)
        '.current_template.SpecType = cboSelectSpecificationType.value
        ' Set manager.current_spec = New Specification
        ' .current_spec.SpecType = .current_template.SpecType
        ' .current_spec.Revision = 0#
'        .current_template.Properties.Item(Utils.ConvertToCamelCase( _
'                cboSelectProperty.value)) = txtPropertyValue
        Set .current_spec.Properties = .current_template.Properties
        Set .current_spec.Tolerances = .current_template.Properties
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
    If SpecManager.SaveSpecification(manager.current_spec) <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & DB_PUSH_FAILURE
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & DB_PUSH_SUCCESS
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Private Sub cmdSetProperty_Click()
' This executes a set property command
    With manager.current_spec
        .Properties.Item(Utils.ConvertToCamelCase( _
                cboSelectProperty.value)) = txtPropertyValue
    End With
    SpecManager.PrintSpecification Me
End Sub

Private Sub PopulateCboSelectProperty()
    Dim prop As Variant
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    With cboSelectProperty
        For Each prop In manager.current_spec.Properties
          .AddItem Utils.SplitCamelCase(CStr(prop))
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

Private Sub UserForm_Terminate()
    Set manager = Nothing
End Sub
