VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formViewSpecs 
   Caption         =   "Specification Control"
   ClientHeight    =   7044
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
    'App.gDll.ShowDialog "Function Disabled.", vbOkOnly, "Under Development"
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
            .value = rev
        Next rev
    End With
End Sub

Sub MaterialSearch()
    SpecManager.RestartApp
    SpecManager.MaterialInput UCase(txtMaterialId)
    SpecManager.PrintSpecification Me
    PopulateCboSelectType
    cboSelectType.value = App.current_spec.SpecType
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub SelectType()
    Set App.current_spec = App.specs.item(cboSelectType.value)
    SpecManager.PrintSpecification Me
End Sub

Sub PrintConsole()
' This subroutine prints the contents of the console box using the default printer assign in user settings.
    'Check if there is actually text to print
    Dim spec As Specification
    Dim T As Variant
    Dim new_sht As Worksheet
    If Me.txtConsole.text = vbNullString Then
        PromptHandler.Error "There is nothing to print!"
    Else
        ' Print the specs one at a time to the default printer
        App.gDll.ShowDialog "Function Disabled", vbOkOnly, "Under Development"
    End If
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    If isTest Then
        App.printer.PrintObjectToSheet App.current_spec, Sheets("pdf")
        GuiCommands.DocumentPrinterToPdf_Test
    Else
        App.printer.PrintObjectToSheet App.current_spec, Sheets(App.current_spec.SpecType)
        App.printer.ToPDF Sheets(App.current_spec.SpecType)
    End If
End Sub
