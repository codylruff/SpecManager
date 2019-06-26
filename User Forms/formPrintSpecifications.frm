VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrintSpecifications 
   Caption         =   "Spec-Manager "
   ClientHeight    =   6612
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6540
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
    WriteAllDocuments
    PrintSelectedPackage PromptHandler.ProtectionPlanningSequence
End Sub

Private Sub cmdSearch_Click()
   If txtMaterialId = vbNullString Or txtMaterialId = " " Then
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
End Sub

Sub Back()
    Unload Me
    GuiCommands.GoToMain
End Sub

Sub PrintSelectedPackage(selected_pacakge As ProtectionPackage)
' Prints the select document package for protection
    Dim origCalcMode As xlCalculation

    ' Toggle Gui Functions to speed up printing
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    origCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' Select document package
    Select Case selected_package
        Case WeavingTieIn
            'SpecManager.PrintPackage DropKeys(App.specs, Array("Tie-Back Checklist"))
        Case WeavingTieBack
            'SpecManager.PrintPackage DropKeys(App.specs, Array("Tie-In Checklist"))
        Case FinishingWithQC
            'SpecManager.PrintPackage App.specs
        Case FinishingNoQC
            'SpecManager.PrintPackage DropKeys(App.specs, Array("Testing Requirements", "Ballistic Testing Requirements"))
        Case Else
            'SpecManager.PrintPackage App.specs
    End Select

    ' Toggle Gui Functions to speed up printing
    GuiCommands.ResetExcelGUI
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = origCalcMode
End Sub

Sub WriteAllDocuments()
' Write all specification docs to the correct worksheets / create worksheet if missing
    Dim spec As Specification
    Dim sheetsToPrint
    Dim T As Variant
    Dim origCalcMode As xlCalculation
    Dim new_sht As Worksheet

    If Me.txtConsole.text = "No specifications are available for this code." Then
         MsgBox "There is nothing to print!"
         Exit Sub
    ElseIf Not IsNumeric(Me.txtProductionOrder) Then
         MsgBox "Please enter a production order."
         Exit Sub
    End If

    ' Toggle Gui Functions to speed up printing
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    origCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual    

    For Each T In App.specs
        Set spec = App.specs.Item(T)
            App.console.PrintObjectToSheet spec, _
                        Utils.CreateNewSheet(spec.SpecType), _
                        txtProductionOrder
        End If
    Next T
    ' Toggle Gui Functions to speed up printing
    GuiCommands.ResetExcelGUI
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = origCalcMode
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.console.PrintObjectToSheet App.specs.Item("Testing Requirements"), Utils.CreateNewSheet("pdf"), txtProductionOrder
    If isTest Then
        GuiCommands.ConsoleBoxToPdf_Test
    Else
        GuiCommands.ConsoleBoxToPdf
    End If
End Sub
