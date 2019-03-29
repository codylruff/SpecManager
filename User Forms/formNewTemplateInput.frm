VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewTemplateInput 
   Caption         =   "Create New Template"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4476
   OleObjectBlob   =   "formNewTemplateInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNewTemplateInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdCancel_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdContinue_Click()
    If SpecManager.TemplateInput(txtTemplateName.value) <> vbNullString Then
        Unload Me
        formCreateGeneric.Show
    Else
        MsgBox "Please enter a template name !"
        Exit Sub
    End If
End Sub

Sub Continue()
    If SpecManager.TemplateInput(txtTemplateName.value) <> vbNullString Then
        Logger.Log "Template Input Pass"
    Else
        Logger.Log "Template Input Fail"
    End If
End Sub
