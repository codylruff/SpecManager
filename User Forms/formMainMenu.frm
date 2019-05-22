VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMainMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   3444
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6300
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
    If App.current_user.PrivledgeLevel < USER_MANAGER Then
        Logger.Log App.current_user.Name & " attempted access to a restricted function."
        MsgBox "This function is not availble to you"
        Exit Sub
    End If
    Unload Me
    formNewSpecInput.Show vbModeless
End Sub

Private Sub cmdCreateTemplate_Click()
' Form to create a new template specification. Admin required.
    SpecManager.RestartApp
    Logger.Log App.current_user.Name
    Logger.Log App.current_user.PrivledgeLevel
    If App.current_user.PrivledgeLevel < USER_MANAGER Then
        Logger.Log App.current_user.Name & " attempted access to a restricted function."
        MsgBox "This function is not availble to you"
        Exit Sub
    End If
    Unload Me
    formNewTemplateInput.Show vbModeless
End Sub

Private Sub cmdExit_Click()
    SpecManager.StopApp
    GuiCommands.UnloadAllForms
End Sub

Private Sub cmdEditTemplates_Click()
' Form to edit an existing specification template. Admin required.
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel < USER_MANAGER Then
        Logger.Log App.current_user.Name & " attempted access to a restricted function."
        MsgBox "This function is not availble to you"
        Exit Sub
    End If
    Unload Me
    formEditTemplate.Show vbModeless
End Sub

Private Sub cmdViewSpecifications_Click()
' Form to view existing specifications. Admin not required.
    SpecManager.RestartApp
    Unload Me
    formViewSpecs.Show vbModeless
End Sub

Private Sub cmdEditSpecifications_Click()
' Form to edit an existing specification. Admin required
    SpecManager.RestartApp
    If App.current_user.PrivledgeLevel < USER_MANAGER Then
        Logger.Log App.current_user.Name & " attempted access to a restricted function."
        MsgBox "This function is not availble to you"
        Exit Sub
    End If
    Unload Me
    formSpecConfig.Show vbModeless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
