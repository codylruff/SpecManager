VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrintSpecifications 
   Caption         =   "Spec-Manager "
   ClientHeight    =   6612
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "formPrintSpecifications.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formPrintSpecifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit

Private Sub cmdPrintSpecifications_Click()
    Dim prompt_result As DocumentPackageVariant
    If App.printer.FormId Is Nothing Then
         PromptHandler.Error "There is nothing to print!"
         Exit Sub
    ElseIf App.printer.CurrentText = "No specifications are available for this code." Then
         PromptHandler.Error "There is nothing to print!"
         Exit Sub
    ElseIf Not IsNumeric(Me.txtProductionOrder) Then
         PromptHandler.Error "Please enter a production order."
         Exit Sub
    End If
    prompt_result = PromptHandler.ProtectionPlanningSequence
    App.printer.WriteAllDocuments Me.txtProductionOrder, prompt_result
    If Not App.TestingMode Then PrintSelectedPackage(prompt_result)
End Sub

Private Sub cmdSearch_Click()
    ' Check for any white space and remove it
    If Utils.RemoveWhiteSpace(txtMaterialId) = vbNullString Then
       PromptHandler.Error "Please enter a material id."
       Exit Sub
    End If
    MaterialSearch
End Sub

Private Sub UserForm_Initialize()
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdBack_Click()
    Back
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearThisForm
End Sub

Private Sub ClearThisForm()
    ClearForm Me
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
    SpecManager.ListSpecifications Me
    If Me.txtConsole.text = vbNullString Then
        Me.txtConsole.text = "No specifications are available for this code."
    End If
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub PrintSelectedPackage(selected_package As DocumentPackageVariant)
' Prints the select document package for protection

    ' Select document package
    Select Case selected_package
        Case WeavingStyleChange
            Logger.Log "Printing Weaving Style Change Package"
            App.printer.PrintPackage App.specs, selected_package
        Case WeavingTieBack
            Logger.Log "Print Weaving Tie-Back Package"
            App.printer.PrintPackage App.specs, selected_package
        Case FinishingWithQC
            Logger.Log "Printing Finishing with QC Package"
            App.printer.PrintPackage App.specs, selected_package
        Case FinishingNoQC
            Logger.Log "Printing Finishing without QC Package"
            App.printer.PrintPackage DropKeys(App.specs, _
                        Array("Testing Requirements", "Ballistic Testing Requirements")), selected_package
        Case Default
            Logger.Log "Printing All Available Specs"
            App.printer.PrintPackage App.specs, selected_package
    End Select

    
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.printer.PrintObjectToSheet App.specs.item("Testing Requirements"), Utils.CreateNewSheet("pdf"), txtProductionOrder
    If isTest Then
        GuiCommands.DocumentPrinterToPdf_Test
    Else
        GuiCommands.DocumentPrinterToPdf
    End If
End Sub
