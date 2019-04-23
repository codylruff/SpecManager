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
    With App.current_template
        .AddProperty Utils.ConvertToCamelCase(txtPropertyName.value)
    End With
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Private Sub cmdClear_Click()
    ClearThisForm
End Sub

Private Sub cmdRemoveProperty_Click()
    App.current_template.RemoveProperty Utils.ConvertToCamelCase(cboSelectProperty.value)
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Sub AddProperty()
    ' This executes an add property command
    With App.current_template
        .AddProperty Utils.ConvertToCamelCase(txtPropertyName.value)
    End With
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Sub RemoveProperty()
    App.current_template.RemoveProperty Utils.ConvertToCamelCase(cboSelectProperty.value)
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
    PopulateCboSelectTemplate
    Set App.console = Factory.CreateConsoleBox(Me)
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
    Dim ret_val As Long
    App.current_template.Revision = CStr(CDbl(App.current_template.Revision) + 1) & ".0"
    ret_val = SpecManager.UpdateSpecificationTemplate(App.current_template)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "Template Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "Template Saved Succesfully Saved."
    End If
End Sub

Private Sub ClearThisForm()
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    Do While cboSelectTemplate.ListCount > 0
        cboSelectTemplate.RemoveItem 0
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
        For Each prop In App.current_template.Properties
          .AddItem Utils.SplitCamelCase(CStr(prop))
        Next prop
    End With
    txtPropertyName.value = vbNullString
    cboSelectProperty.value = vbNullString
End Sub

Private Sub PopulateCboSelectTemplate()
    Dim coll As VBA.Collection
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

Sub SearchTemplates()
    SpecManager.LoadExistingTemplate cboSelectTemplate
    SpecManager.PrintTemplate Me
    PopulateCboSelectProperty
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub SaveChanges()
' Calls method to save a new specification incremented the revision by +0.1
    Dim ret_val As Long
    App.current_template.Revision = CStr(CDbl(App.current_template.Revision) + 1) & ".0"
    ret_val = SpecManager.UpdateSpecificationTemplate(App.current_template)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "Edit Template Fail"
    Else
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "Edit Template Pass"
    End If
End Sub
