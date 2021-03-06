VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrintSpecifications 
   Caption         =   "Spec-Manager "
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
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
    'MsgBox "Function un-available"
    PrintSelectedSpecs PromptHandler.ProtectionPlannerSequence
    'Debug.Print PromptHandler.ProtectionPlannerSequence
    'ExportPdf
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

Sub PrintSelectedSpecs(setup_only As Boolean)
' This subroutine prints the contents of the console box using the default printer assign in user settings.
    'Check if there is actually text to print
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long
    Dim spec As Specification
    Dim T As Variant
    Dim new_sht As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If Me.txtConsole.text = "No specifications are available for this code." Then
         MsgBox "There is nothing to print!"
    ElseIf Not IsNumeric(Me.txtProductionOrder) Then
         MsgBox "Please enter a production order."
    Else
        ' Print the specs one at a time to the default printer
        If setup_only Then
            Dim setup_spec As Specification
            If App.specs.exists("Setup Requirements") Then
                ' Print only the setup spec
                Set new_sht = Utils.CreateNewSheet(spec.SpecType)
                Set spec = App.specs.Item("Setup Requirements")
                App.console.PrintObjectToSheet spec, new_sht, vbNullString
                Utils.PrintSheet new_sht, True
            Else
                MsgBox "No Setup Requirements Exist for this Material."
            End If
            Exit Sub
        End If
            lngNumberOfTasks = App.specs.Count
            lngCounter = 0
        For Each T In App.specs
            lngCounter = lngCounter + 1
            Set spec = App.specs.Item(T)
            modProgress.ShowProgress _
                lngCounter,lngNumberOfTasks,"Printing : " & spec.SpecType, IIf(lngCounter = lngNumberOfTasks,True, False)
            If spec.SpecType = "Testing Requirements" Or spec.SpecType = "Ballistic Testing Requirements" Then
                Set new_sht = Utils.CreateNewSheet(spec.SpecType)
                App.console.PrintObjectToSheet spec, new_sht, txtProductionOrder
                Utils.PrintSheet new_sht, True
            Else 
                App.console.PrintObjectToSheet spec, new_sht, txtProductionOrder
                Utils.PrintSheet shtRBA
            End If
        Next T
    End If
    Set ActiveSheet = shtStart
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub ExportPdf(Optional isTest As Boolean = False)
    App.console.PrintObjectToSheet App.specs.Item("Testing Requirements"), Utils.CreateNewSheet("pdf"), txtProductionOrder
    If isTest Then
        GuiCommands.ConsoleBoxToPdf_Test
    Else
        GuiCommands.ConsoleBoxToPdf
    End If
End Sub
