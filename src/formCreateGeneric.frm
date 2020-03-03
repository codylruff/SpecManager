VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateGeneric 
   Caption         =   "Create New Spec Type"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9288
   OleObjectBlob   =   "formCreateGeneric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCreateGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












































Option Explicit

Private template_name As String

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
    lblTemplateName = App.current_template.SpecType
    Set App.printer = Factory.CreateDocumentPrinter(Me)
End Sub

Private Sub cmdAddProperty_Click()
   AddProperty
End Sub

Sub AddProperty()
   App.printer.WriteLine Me.txtPropertyName
   App.current_template.AddProperty Me.txtPropertyName
End Sub

Private Sub cmdSubmitTemplate_Click()
    Dim ret_val As Long
    ret_val = SpecManager.SaveSpecificationTemplate(App.current_template)
   If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        PromptHandler.Error "New Specification Was Not Saved"
    Else
        Logger.Log "Data Access returned: " & ret_val & ", New Template Succesfully Saved.", DebugLog
        PromptHandler.Success "New Template Succesfully Saved."
        Set App.templates = SpecManager.GetAllTemplates
    End If
    
End Sub

Sub SubmitTemplate()
    Dim ret_val As Long
    ret_val = SpecManager.SaveSpecificationTemplate(App.current_template)
   If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        Logger.Log "Create Template Fail"
    Else
        Logger.Log "Data Access returned: " & ret_val & ", New Template Succesfully Saved.", DebugLog
        Logger.Log "Create Template Pass"
        Set App.templates = SpecManager.GetAllTemplates
    End If
End Sub

Private Sub UserForm_Terminate()
    Logger.Log "--------- End " & Me.Name & " ----------"
End Sub
