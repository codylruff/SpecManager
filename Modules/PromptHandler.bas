Attribute VB_Name = "PromptHandler"
Option Explicit
Public Choice As Boolean

Public Enum DocumentPackageVariant
    Default = 0
    WeavingStyleChange = 1
    WeavingTieBack = 2
    FinishingWithQC = 3
    FinishingNoQC = 4
End Enum

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
    question "Is this a finishing order?"
    formUserPrompt.Show
    If Choice = True Then
        ' Prompt #2 : Is this the first loom cut?
        question "Is this the first loom cut?"
        formUserPrompt.Show
        If Choice = True Then
            ' Prompt #3 : After finishing, will this roll be processed on the Isotex?
            question "After finishing, will this roll be processed on the Isotex?"
            formUserPrompt.Show
            If Choice = True Then
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
    Else
        ' Prompt #4 : Is this a straight tie-back?
        question "Is this a straight tie-back?"
        formUserPrompt.Show
        If Choice = True Then
            ProtectionPlanningSequence = WeavingTieBack
        Else
            ProtectionPlanningSequence = WeavingStyleChange
        End If
    End If

End Function

Private Sub question(question_text As String)
    formUserPrompt.lblQuestion.Caption = question_text
End Sub
