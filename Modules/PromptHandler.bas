Attribute VB_Name = "PromptHandler"
Option Explicit
Public Choice As Boolean

' Prompt Sequences
' A prompt sequence is a series of prompts and conditionals
' used to determine the final outcome of events
Function ProtectionPlannerSequence() As Boolean
' This sequence is shown to the protection planners upon clicking print.
'-----------------------------------------------------------------------
' 1. Is this a finishing order?
'       a. If no then proceed to print all specifications
' 2. Is this the first cut?
'       a. If no then print only setupd documents
' 3. After finishing, will this roll be processed on the Isotex?
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
                ProtectionPlannerSequence = True
                Exit Function
            Else
                ProtectionPlannerSequence = False
                Exit Function
            End If
        Else
            ProtectionPlannerSequence = True
            Exit Function
        End If
        ProtectionPlannerSequence = False
    End If
End Function

Private Sub question(question_text As String)
    formUserPrompt.lblQuestion.Caption = question_text
End Sub
