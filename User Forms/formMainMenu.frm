VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMainMenu 
   Caption         =   "SAATI Spec-Manager"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5520
   OleObjectBlob   =   "formMainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


















Option Explicit

Private Sub cmdConfig_Click()
' Initialze developer configuration mode
    Unload Me
    AccessControl.ConfigControl
End Sub

Private Sub cmdCreateSpecification_Click()
' Form to create a new specification based on an existing template. Admin required
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel <> USER_ADMIN Then
        App.logger.Log App.current_user.Name & " attempted access to a restricted function.", UserLog
        PromptHandler.AccessDenied
        Exit Sub
    End If
    Unload Me
    formNewSpecInput.show vbModeless
End Sub

Private Sub cmdCreateTemplate_Click()
' Form to create a new template specification. Admin required.
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel <> USER_ADMIN Then
        App.logger.Log App.current_user.Name & " attempted access to a restricted function.", UserLog
        PromptHandler.AccessDenied
        Exit Sub
    End If
    Unload Me
    formNewTemplateInput.show vbModeless
End Sub

Private Sub cmdDatabaseQuery_Click()
    App.gDll.ShowDialog "Feature Under Development.", vbOkOnly, "Under Development"
End Sub

Private Sub cmdExit_Click()
    SpecManager.StopApp
    GuiCommands.UnloadAllForms
    'GuiCommands.ExitApp
End Sub

Private Sub cmdEditTemplates_Click()
' Form to edit an existing specification template. Admin required.
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel <> USER_ADMIN Then
        App.logger.Log App.current_user.Name & " attempted access to a restricted function.", UserLog
        PromptHandler.AccessDenied
        Exit Sub
    End If
    Unload Me
    formEditTemplate.show vbModeless
End Sub

Private Sub cmdViewDocuments_Click()
    SpecManager.RestartApp
    Unload Me
    formViewSpecs.show vbModeless
End Sub

Private Sub cmdViewSpecifications_Click()
' Form to view existing specifications. Admin not required.
    SpecManager.RestartApp
    Unload Me
    formPrintSpecifications.show vbModeless
End Sub

Private Sub cmdEditSpecifications_Click()
' Form to edit an existing specification. Admin required
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel <> USER_ADMIN Then
        App.logger.Log App.current_user.Name & " attempted access to a restricted function.", UserLog
        PromptHandler.AccessDenied
        Exit Sub
    End If
    Unload Me
    formSpecConfig.show vbModeless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
