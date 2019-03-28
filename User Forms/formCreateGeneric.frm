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

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
    lblTemplateName = manager.current_template.SpecType
    Set manager.console = Factory.CreateConsoleBox(Me)
End Sub

Private Sub cmdAddProperty_Click()
   manager.console.PrintLine Me.txtPropertyName
   manager.current_template.AddProperty Utils.ConvertToCamelCase(CStr(Me.txtPropertyName))
End Sub

Private Sub cmdSubmitTemplate_Click()
   If SpecManager.SaveSpecTemplate(manager.current_template) <> DB_PUSH_SUCCESS Then
      Logger.Log "Data Access returned: " & DB_PUSH_FAILURE
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "Data Access returned: " & DB_PUSH_SUCCESS & ", New Template Succesfully Saved."
        MsgBox "New Template Succesfully Saved."
    End If
End Sub

Private Sub UserForm_Terminate()
    Logger.Log "--------- End " & Me.Name & " ----------"
End Sub
