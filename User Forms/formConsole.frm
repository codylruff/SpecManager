VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formConsole 
   Caption         =   "Spec-Manager Console v0.0.1"
   ClientHeight    =   6480
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14292
   OleObjectBlob   =   "formConsole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents pLogger As Logger

Public Sub ListenTo(logger_ As Logger)
    If Not pLogger Is Nothing Then
        Exit Sub
    End If 
    Set pLogger = Logger
End Sub

Private Sub pLogger_LogChanged(log_text As String)
    Me.txtConsole.Value = Me.txtConsole.Value & vbNewLine & log_text
End Sub

Private Sub txtInput_Enter()
    ParseConsoleCommand
End Sub

Private Sub UserForm_Initialize()
    Me.ListenTo App.Log
End Sub

Private Sub UserForm_Terminate()
    Me.ListenTo Nothing
End Sub

Private Sub ParseConsoleCommand(command As String)
    If command = "Exit" Then
        Back
    End If
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub