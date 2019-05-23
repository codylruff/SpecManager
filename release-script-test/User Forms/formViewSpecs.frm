VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formViewSpecs 
   Caption         =   "Specification Control"
   ClientHeight    =   11865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   OleObjectBlob   =   "formViewSpecs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formViewSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdPrint_Click()
    PrintConsole
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdMaterialSearch_Click()
    MaterialSearch
End Sub

Private Sub cmdBack_Click()
    Back
End Sub

Private Sub cmdExportPdf_Click()
    MsgBox "Function Disabled."
End Sub

Private Sub ClearThisForm()
    ClearForm Me
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearThisForm
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Logger.Log "--------- End " & Me.Name & " ----------"
End Sub

Sub MaterialSearch()
    SpecManager.RestartApp
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub PrintConsole()
' This subroutine prints the contents of the console box using the default printer assign in user settings.
    If Me.txtConsole.Text = vbNullString Then
        MsgBox "There is nothing to print!"
    Else
        App.console.PrintObjectToSheet App.current_spec, shtSpecificationForm
        Utils.PrintSheet shtSpecificationForm
    End If
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.console.PrintObjectToSheet App.current_spec, shtSpecificationForm
    If isTest Then
        GuiCommands.ConsoleBoxToPdf_Test
    Else
        GuiCommands.ConsoleBoxToPdf
    End If
End Sub
