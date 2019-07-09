VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formViewSpecs 
   Caption         =   "Specification Control"
   ClientHeight    =   11865
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9816
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

Private Sub cmdSelectType_Click()
    SelectType
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
    ExportPdf
    'MsgBox "Function Disabled."
End Sub

Private Sub ClearThisForm()
    Dim i As Integer
    Do While cboSelectType.ListCount > 0
        cboSelectType.RemoveItem 0
    Loop
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

Private Sub PopulateCboSelectType()
    Dim rev As Variant
    Dim i As Integer
    Do While cboSelectType.ListCount > 0
        cboSelectType.RemoveItem 0
    Loop
    With cboSelectType
        For Each rev In App.specs
            .AddItem rev
            .Value = rev
        Next rev
    End With
End Sub

Sub MaterialSearch()
    SpecManager.RestartApp
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectType
    cboSelectType.Value = App.current_spec.SpecType
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub SelectType()
    Set App.current_spec = App.specs.Item(cboSelectType.Value)
    SpecManager.PrintSpecification Me
End Sub

Sub PrintConsole()
' This subroutine prints the contents of the console box using the default printer assign in user settings.
    'Check if there is actually text to print
    Dim spec As Specification
    Dim T As Variant
    Dim new_sht As Worksheet
    If Me.txtConsole.text = vbNullString Then
        MsgBox "There is nothing to print!"
    Else
        ' Print the specs one at a time to the default printer
        For Each T In App.specs
            Set spec = App.specs.Item(T)
            Set new_sht = Utils.CreateNewSheet(spec.SpecType)
            App.printer.PrintObjectToSheet spec, new_sht
            Application.PrintCommunication = False
            With new_sht.PageSetup
                .FitToPagesWide = 1
                .FitToPagesTall = False
            End With
            Application.PrintCommunication = True
            Utils.PrintSheet new_sht
        Next T
    End If
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.printer.PrintObjectToSheet App.current_spec, Sheets("pdf")
    If isTest Then
        GuiCommands.DocumentPrinterToPdf_Test
    Else
        GuiCommands.DocumentPrinterToPdf
    End If
End Sub
