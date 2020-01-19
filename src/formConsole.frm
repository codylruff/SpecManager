VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formConsole 
   Caption         =   "Spec-Manager Console v0.0.1"
   ClientHeight    =   6480
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14625
   OleObjectBlob   =   "formConsole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







































Private WithEvents pLogger As SystemLogger
Attribute pLogger.VB_VarHelpID = -1

Public Sub ListenTo(logger_ As SystemLogger)
    If Not pLogger Is Nothing Then
        Exit Sub
    End If
    Set pLogger = logger_
End Sub

Private Sub pLogger_LogChanged(log_text As String)
    Me.txtConsole.value = Me.txtConsole.value & vbNewLine & log_text
End Sub

Private Sub txtInput_Enter()
    ParseConsoleCommand txtInput.text
End Sub

Private Sub UserForm_Initialize()
    Me.ListenTo App.Logger
End Sub

Private Sub UserForm_Terminate()
    Me.ListenTo Nothing
End Sub

Private Sub ParseConsoleCommand(command As String)
    On Error GoTo ErrorHandler
    If command = "Exit" Then
        Back
    Else
        Application.Run command
    End If
    txtInput.text = nullstr
    Exit Sub
ErrorHandler:
    txtInput.text = nullstr
    Logger.Log "Function Unknown"
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub
