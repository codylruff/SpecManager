VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   11865
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

Private Sub cmdCopySpec_Click()
' Makes a copy of the current spec, with a new material id
    Dim new_material_id As String
    Dim ret_val As Long
    new_material_id = PromptHandler.UserInput(SingleLineText, "Material Id", "Enter a material id for copy?")
    ret_val = SpecManager.CreateSpecificationFromCopy(App.current_spec, new_material_id)
    If ret_val = DB_PUSH_SUCCESS Then
        PromptHandler.Success "Copied Successfully"
    Else
        PromptHandler.Error "Copy Failed"
    End If
End Sub

Private Sub cmdSelectType_Click()
    SelectType
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdMaterialSearch_Click()
    If txtMaterialId = nullstr Or txtMaterialId = " " Then
      PromptHandler.Error "Please enter a material id."
      Exit Sub
   End If
    MaterialSearch
End Sub

Private Sub cmdBack_Click()
    Back
End Sub

Private Sub cmdSaveChanges_Click()
' Calls method to save a new specification incrementing the revision by 1
    SaveChanges
End Sub

Private Sub cmdSubmit_Click()
' This executes a set property command
    Submit
End Sub

Private Sub ClearThisForm()
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    Do While cboSelectType.ListCount > 0
        cboSelectType.RemoveItem 0
    Loop
    ClearForm Me
End Sub

Private Sub PopulateCboSelectType()
    Dim spec_id As Variant
    Dim i As Integer
    Do While cboSelectType.ListCount > 0
        cboSelectType.RemoveItem 0
    Loop
    With cboSelectType
        For Each spec_id In App.DocumentsByUID
            .AddItem spec_id
            .value = spec_id
        Next spec_id
    End With
End Sub

Private Sub PopulateCboSelectProperty()
    Dim prop As Variant
    Dim i As Integer
    Do While cboSelectProperty.ListCount > 0
        cboSelectProperty.RemoveItem 0
    Loop
    
    With cboSelectProperty
        For Each prop In App.current_spec.Properties
            If App.printer.ShouldPrint((prop)) Then
                .AddItem Utils.SplitCamelCase(CStr(prop))
            End If
        Next prop
    End With
    txtPropertyValue.value = nullstr
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
    PopulateCboSelectType
    cboSelectType.value = App.current_spec.UID
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub ExportPdf()
    GuiCommands.DocumentPrinterToPdf
End Sub

Sub SaveChanges()
' Calls method to save a new specification incremented the revision by +0.1
    Dim ret_val As Long
    Dim old_spec As Specification
    Set old_spec = New Specification
    Set old_spec = Factory.CopySpecification(App.current_spec)
    App.specs.Add "to_archive", old_spec
    App.current_spec.Revision = CStr(CDbl(old_spec.Revision) + 1)
    ret_val = SpecManager.SaveSpecification(App.current_spec, old_spec)
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        Logger.Log "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        Logger.Log "New Specification Succesfully Saved."
    End If
End Sub

Sub Submit()
' This executes a set property command
    ' Check for empty controls
    If cboSelectProperty.value = nullstr Then Exit Sub
    ' Change the property desired
    With App.current_spec
        .ChangeProperty cboSelectProperty.value, txtPropertyValue
    End With
    SpecManager.PrintSpecification Me
End Sub

Sub SelectType()
    ' Check for empty controls
    If cboSelectType.value = nullstr Then Exit Sub
    ' Select the specification desired
    Set App.current_spec = App.specs.item(cboSelectType.value)
    PopulateCboSelectProperty
    SpecManager.PrintSpecification Me
End Sub
