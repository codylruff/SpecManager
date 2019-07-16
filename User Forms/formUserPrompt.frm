VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formUserPrompt 
   Caption         =   "User Prompt"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6900
   OleObjectBlob   =   "formUserPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formUserPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Option Explicit

'Public YesMethod As String
'Public NoMethod As String

Private Sub cmdNo_Click()
   'CallByName(PromptHandler,NoMethod,vbMethod)
   PromptHandler.Choice = False
   Unload Me
End Sub

Private Sub cmdYes_Click()
   'CallByName(PromptHandler,YesMethod,vbMethod)
   PromptHandler.Choice = True
   Unload Me
End Sub
