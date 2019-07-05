VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrintSpecifications 
   Caption         =   "Spec-Manager "
   ClientHeight    =   6612
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6684
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
    If App.console.CurrentText = "No specifications are available for this code." Then
         MsgBox "There is nothing to print!"
         Exit Sub
    ElseIf Not IsNumeric(Me.txtProductionOrder) Then
         MsgBox "Please enter a production order."
         Exit Sub
    End If
    WriteAllDocuments Me.txtProductionOrder
    PrintSelectedPackage PromptHandler.ProtectionPlanningSequence
End Sub

Private Sub cmdSearch_Click()
    ' Check for any white space and remove it
    If Utils.RemoveWhiteSpace(txtMaterialId) = vbNullString Then
       MsgBox "Please enter a material id."
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

Sub PrintSelectedPackage(selected_package As ProtectionPackage)
' Prints the select document package for protection

    ' Select document package
    Select Case selected_package
        Case WeavingStyleChange
            Logger.Log "Printing Weaving Style Change Package"
            SpecManager.PrintPackage DropKeys(App.specs, Array("Tie Back Checklist"))
        Case WeavingTieBack
            Logger.Log "Print Weaving Tie-Back Package"
            SpecManager.PrintPackage DropKeys(App.specs, Array("Style Change Checklist"))
        Case FinishingWithQC
            Logger.Log "Printing Finishing with QC Package"
            SpecManager.PrintPackage App.specs
        Case FinishingNoQC
            Logger.Log "Printing Finishing without QC Package"
            SpecManager.PrintPackage DropKeys(App.specs, _
                        Array("Testing Requirements", "Ballistic Testing Requirements"))
        Case Else
            Logger.Log "Printing All Available Specs"
            SpecManager.PrintPackage App.specs
    End Select

    
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.console.PrintObjectToSheet App.specs.Item("Testing Requirements"), Utils.CreateNewSheet("pdf"), txtProductionOrder
    If isTest Then
        GuiCommands.DocumentPrinterToPdf_Test
    Else
        GuiCommands.DocumentPrinterToPdf
    End If
End Sub
