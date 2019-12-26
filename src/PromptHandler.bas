Attribute VB_Name = "PromptHandler"
Option Explicit

' Prompt Sequences
' A prompt sequence is a series of prompts and conditionals
' used to determine the final outcome of events

Function ProtectionPlanningSequence() As DocumentPackageVariant
' This sequence is shown to the protection planners upon clicking print.
'-----------------------------------------------------------------------
' 1. Is this a finishing order?
'       a. If no then proceed check for tie-back
' 2. Is this the first cut?
'       a. If no then print only setupd documents
' 3. After finishing, will this roll be processed on the Isotex?
' 4. Is this a straigh tie-back?
'       a. If yes then proceed to print tie-back checklist
' Roll is first cut so if no then print all specifications
'------------------------------------------------------------------------
    ' Prompt #1 : Is this a finishing order?
    If App.current_spec.ProcessId = "Finishing" Then
        ' Prompt #2 : Is this the first loom cut?
        
        If question("Is this the first loom cut?") Then
            ' Prompt #3 : After finishing, will this roll be processed on the Isotex?
            If question("After finishing, will this roll be processed on the Isotex?") Then
                ProtectionPlanningSequence = FinishingNoQC
                Exit Function
            Else
                ProtectionPlanningSequence = FinishingWithQC
                Exit Function
            End If
        Else
            ProtectionPlanningSequence = FinishingNoQC
            Exit Function
        End If
    ElseIf App.current_spec.ProcessId = "Weaving" Then
        ' Prompt #4 : Is this a straight tie-back?
        If question("Is this a straight tie-back?") Then
            ProtectionPlanningSequence = WeavingTieBack
        Else
            ProtectionPlanningSequence = WeavingStyleChange
        End If
    ElseIf App.current_spec.ProcessId = "Isotex" Then
        ProtectionPlanningSequence = Isotex
    Else
        ProtectionPlanningSequence = Default
    End If

End Function

Private Function question(question_text As String) As Boolean
    question = iif(App.GUI.ShowDialog(question_text, vbYesNo, "Question", "Yes", "No") = vbYes, True, False)
End Function

Public Sub AccessDenied()
' Shows an access denied prompt
    If Not App.TestingMode Then App.GUI.ShowDialog "Access Denied", vbCritical, "Access Control", ThemeBg:="#f44336"
End Sub

Public Sub Error(message_text As String)
' Shows a handled error message
    If Not App.TestingMode Then App.GUI.ShowDialog message_text, vbCritical, "Error Message", ThemeBg:="#f44336"
End Sub

Public Sub Success(message_text As String)
    If Not App.TestingMode Then App.GUI.ShowDialog message_text, vbOkOnly, "Success!"
End Sub

Public Function UserInput(input_type As InputBoxType, title_text As String, message_text As String) As Variant
    UserInput = App.GUI.CreateInputBox(input_type, title_text, message_text)
End Function

Public Function GetPassword() As String
    GetPassword = CStr(UserInput(Password, "Access Control", "Enter Your Password :"))
End Function

Public Function GetMachineId() As String
    GetMachineId = CStr(PromptHandler.UserInput(SingleLineText, "Machine Id Selection", "Enter a Machine Id for this specification:"))
End Function

Public Function ChangePassword() As String
    Dim new_pass_1 As String
    Dim new_pass_2 As String
    new_pass_1 = "new_password_1"
    new_pass_2 = "new_password_2"
    While new_pass_1 <> new_pass_2
        new_pass_1 = UserInput(Password, "Access Control", "Enter Your New Password :")
        new_pass_2 = UserInput(Password, "Access Control", "Confirm Your New Password :")
        If new_pass_1 <> new_pass_2 Then
            PromptHandler.Error "Passwords don't match!"
        End If
    Wend
    ' Return new password
    ChangePassword = CStr(new_pass_1)
    new_pass_1 = nullstr
    new_pass_2 = nullstr
    
End Function

Public Function EnterTemplateType() As String
    EnterTemplateType = CStr(UserInput(SingleLineText, "Template Selection", "Enter A Template Type :"))
End Function

Public Function SelectSpecifcationFile() As String
' Select an specification file from the file dialog.
    SelectSpecifcationFile = App.GUI.OpenFile("Select Specification Document . . .")
End Function
