VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
Implements IForm
'Factory Module Constructor function:
'Function CreatePlanning() As PlanningForm
' DEFAULT PROPERTY VALUES
Private Const FORM_NAME As String = "Planning Form"
Private Const HELP_TEXT As String = "TODO" ' Thinking of loading this from a text file??
Private m_IForm_Console As Console
Private m_IForm_Fields(3) As Variant
'--------------
' I_PROPERTIES
'--------------
Public Property Set IForm_Console(value As Console): Set m_IForm_Console = value: End Property
Public Property Get IForm_Console() As Console: Set IForm_Console = m_IForm_Console: End Property

Public Property Get IForm_Sheet() As Worksheet
    Set IForm_Sheet = shtPlanning
End Property

Public Property Get IForm_Name() As String
    IForm_Name = FORM_NAME
End Property

Public Property Get IForm_HelpText() As String
    IForm_HelpText = HELP_TEXT
End Property

Public Property Get IForm_Fields() As Variant
    IForm_Fields = m_IForm_Fields
End Property

'----------
' I_METHODS
'----------
Public Function IForm_Clear()
    ' Console must be cleared through object
    IForm_Console.Clear
    ' Fields can be cleared manually.
    For i = 0 To UBound(IForm_Fields) - 1
        IForm_Sheet.Range(IForm_Fields(i)).value = nullstr
    Next i
End Function

Public Function IForm_CmdByName(cmd As String) As Boolean
    Dim ret_val As Boolean
    Select Case cmd
        Case "PrintDocuments"
            cmdPrintDocuments
            ret_val = True
        Case "Search"
            cmdSearch
            ret_val = True
        Case "PrintSelectedPackage"
            'cmdPrintSelectedPackage
            ret_val = False
        Case Default
            ret_val = False
    End Select
    IForm_CmdByName = ret_val
End Function

Public Function IForm_ReformatSheet() As Boolean
    ' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
    IForm_ReformatSheet = False
End Function

'----------
' COMMANDS
'----------
Private Sub cmdPrintDocuments()
    Dim prompt_result As DocumentPackageVariant
    If IForm_Sheet.Range("console") = nullstr Then
         Prompt.Error "There is nothing to print!"
         Exit Sub
    ElseIf IForm_Sheet.Range("console") = "No specifications are available for this code." Then
         Prompt.Error "There is nothing to print!"
         Exit Sub
    ElseIf Not IsNumeric(IForm_Sheet.Range("work_order")) Then
         Prompt.Error "Please enter a production order."
         Exit Sub
    End If
    ' Consider process exceptions based on input from planners.
    prompt_result = Prompt.ProtectionPlanningSequence
    If Not App.TestingMode Then
        ' Check for alternate machine ids (currently only for weaving)
        If App.current_doc.ProcessId = "Weaving" Then
                SpecManager.FilterByMachineId IForm_Sheet.Range("machine_id")
        End If
        ' Write the documents to their repsective worksheets
        App.printer.WriteAllDocuments IForm_Sheet.Range("work_order"), prompt_result
        ' Print all of the documents based on the selected doc package and machine
        PrintSelectedPackage prompt_result
    Else
        Logger.Log CStr(prompt_result)
    End If
    App.Shutdown
End Sub

Private Sub cmdSearch()
    ' Check for any white space and remove it
    App.Start
    If Utils.RemoveWhiteSpace(IForm_Sheet.Range("material_id")) = nullstr Then
       Prompt.Error "Please enter a material id."
       Exit Sub
    ElseIf Utils.RemoveWhiteSpace(IForm_Sheet.Range("work_order")) = nullstr Then
        Prompt.Error "Please enter a work order number."
       Exit Sub
    ElseIf Utils.RemoveWhiteSpace(IForm_Sheet.Range("machine_id")) = nullstr Then
        Prompt.Error "Please enter a machine id."
       Exit Sub
    End If
    ' Run search routine
    MaterialSearch

End Sub

Private Sub MaterialSearch()
    SpecManager.MaterialInput UCase(IForm_Sheet.Range("material_id"))
    Logger.Log "Listing Documents . . . "
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.specs
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
    If IForm_Sheet.Range("console") = nullstr Then
        IForm_Sheet.Range("console") = "No specifications are available for this code."
    End If
End Sub

Private Sub cmdPrintSelectedPackage(selected_package As DocumentPackageVariant)
' Prints the select document package for protection

    ' Select document package
    Select Case selected_package
        Case WeavingStyleChange
            Logger.Log "Printing Weaving Style Change Package"
            App.printer.PrintPackage App.specs, selected_package, IForm_Sheet.Range("work_order")
        Case WeavingTieBack
            Logger.Log "Print Weaving Tie-Back Package"
            App.printer.PrintPackage App.specs, selected_package, IForm_Sheet.Range("work_order")
        Case FinishingWithQC
            Logger.Log "Printing Finishing with QC Package"
            App.printer.PrintPackage App.specs, selected_package, IForm_Sheet.Range("work_order")
        Case FinishingNoQC
            Logger.Log "Printing Finishing without QC Package"
            App.printer.PrintPackage DropKeys(App.specs, _
                        Array("Testing Requirements", "Ballistic Testing Requirements")), selected_package, IForm_Sheet.Range("work_order")
        Case Isotex
            Logger.Log "Printing Isotex TSPP"
            App.printer.PrintPackage App.specs, selected_package, IForm_Sheet.Range("work_order")
        Case Default
            Logger.Log "Printing All Available Specs"
            Debug.Print IsEmpty(App.specs)
            App.printer.PrintPackage App.specs, selected_package, IForm_Sheet.Range("work_order")
    End Select

    
End Sub
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    Set IForm_Console = Factory.CreateConsole(Me.IForm_Sheet)
    ' Initialize the fields array to save on Array() calls.
    m_IForm_Fields = Array("material_id", "work_order", "machine_id")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    'REVIEW Is there anything that needs to be destroyed here?
End Sub
