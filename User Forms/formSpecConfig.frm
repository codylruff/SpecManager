VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   11868
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9816
   OleObjectBlob   =   "formSpecConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formSpecConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdMaterialSearch_Click()
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectProperty
    PopulateCboSelectRevision
    
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdExportPdf_Click()
    GuiCommands.ConsoleBoxToPdf
End Sub

Private Sub cmdSaveChanges_Click()
' Calls method to save a new specification incremented the revision by +0.1
    Dim ret_val Long
    manager.current_spec.Revision = CStr(CDbl(manager.current_spec.Revision) + 1) & ".0"
    ret_val = SpecManager.SaveSpecification(manager.current_spec)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & ret_val
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Private Sub cmdSubmit_Click()
' This executes a set property command
' TODO: Change the name of this to cmdSetProperty
    With manager.current_spec
        .Properties.Item(Utils.ConvertToCamelCase( _
                cboSelectProperty.value)) = txtPropertyValue
        '.Revision = .Properties.Item("Revision")
    End With
    SpecManager.PrintSpecification Me
End Sub

Private Sub cmdSearch_Click()
    Set manager.current_spec = manager.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub

Private Sub ClearThisForm()
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
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
        For Each rev In manager.specs
            .AddItem rev
            .value = rev
        Next rev
    End With
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
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectProperty
    PopulateCboSelectRevision
    
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub ExportPdf()
    GuiCommands.ConsoleBoxToPdf
End Sub

Sub SaveChanges()
' Calls method to save a new specification incremented the revision by +0.1
    Dim ret_val As Long
    manager.current_spec.Revision = CStr(CDbl(manager.current_spec.Revision) + 1) & ".0"
    ret_val = SpecManager.SaveSpecification(manager.current_spec)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & ret_val
        Logger.Log "New Specification Succesfully Saved."
    End If
End Sub

Sub Submit()
' This executes a set property command
' TODO: Change the name of this to cmdSetProperty
    With manager.current_spec
        .Properties.Item(Utils.ConvertToCamelCase( _
                cboSelectProperty.value)) = txtPropertyValue
        '.Revision = .Properties.Item("Revision")
    End With
    SpecManager.PrintSpecification Me
End Sub

Sub Search()
    Set manager.current_spec = manager.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub
