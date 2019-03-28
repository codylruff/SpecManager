VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEditTemplate 
   Caption         =   "Specification Control"
   ClientHeight    =   11868
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9816
   OleObjectBlob   =   "formEditTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formEditTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub cmdAddProperty_Click()
    ' This executes an add property command
    With manager.current_template
        .AddProperty Utils.ConvertToCamelCase(txtPropertyName.value)
    End With
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Private Sub cmdRemoveProperty_Click()
    manager.current_template.RemoveProperty Utils.ConvertToCamelCase(cboSelectProperty.value)
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
    PopulateCboSelectTemplate
    Set manager.console = Factory.CreateConsoleBox(Me)
End Sub

Private Sub cmdSearchTemplates_Click()
    SpecManager.LoadExistingTemplate cboSelectTemplate
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdSaveChanges_Click()
' Calls method to save a new specification incremented the revision by +0.1
    manager.current_template.Revision = CStr(CDbl(manager.current_template.Revision) + 1) & ".0"
    If SpecManager.UpdateSpecTemplate(manager.current_template) <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & DB_PUSH_FAILURE
        MsgBox "Template Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & DB_PUSH_SUCCESS
        MsgBox "Template Saved Succesfully Saved."
    End If
End Sub

Private Sub ClearThisForm()
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    ClearForm Me
End Sub

Private Sub PopulateCboSelectProperty()
    Dim prop As Variant
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    
    With cboSelectProperty
        For Each prop In manager.current_template.Properties
          .AddItem Utils.SplitCamelCase(CStr(prop))
        Next prop
    End With
    txtPropertyName.value = vbNullString
    cboSelectProperty.value = vbNullString
End Sub

Private Sub PopulateCboSelectTemplate()
    Dim coll As Collection
    Dim template_type As Variant
    Set coll = SpecManager.ListAllTemplateTypes
    With cboSelectTemplate
        For Each template_type In coll
            .AddItem CStr(template_type)
        Next template_type
    End With
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
