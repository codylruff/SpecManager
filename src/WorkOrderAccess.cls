VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkOrderAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' This forms is used to print spec docs for protection
'====================================
'Factory Module Constructor function:
'Function CreateWorkOrderAccess() As WorkOrderAccess
Implements IForm
Implements ISearchForm
'----------------
'CLASS VARIABLES
'----------------
Private m_Name As String
Private m_Sheet As Worksheet
'------------
' COLLECTIONS
'------------
'------------
' PROPERTIES
'------------
Public Property Get Sheet() As Worksheet: Set Sheet = m_Sheet: End Property
Public Property Get IForm_Name() As String: Name = m_Name: End Property
'----------
' METHODS
'----------
Public Sub cmdPrintDocuments()
    ISearchForm_DisplayResults
End Sub

Public Sub cmdSearch()
    ' Check for any white space and remove it
    ' Run search routine
    ISearchForm_Search (FormatSearchInput)
End Sub

Public Sub cmdClear()
    IForm_Clear
End Sub

Private Function FormatSearchInput() As String
    If Utils.RemoveWhiteSpace(shtPlanning.Range("material_id")) = nullstr Then
       Prompt.Error "Please enter a material id."
       Exit Function
    ElseIf Utils.RemoveWhiteSpace(shtPlanning.Range("work_order")) = nullstr Then
        Prompt.Error "Please enter a work order number."
       Exit Function
    ElseIf Utils.RemoveWhiteSpace(shtPlanning.Range("machine_id")) = nullstr Then
        Prompt.Error "Please enter a machine id."
       Exit Function
    End If
    FormatSearchInput = Me.Sheet.Range("material_id").value
End Function

Public Function IForm_Clear()
    Logger.NotImplementedException
    Prompt.FNI
End Function

Public Function ISearchForm_Search(id_value As String)
    SpecManager.MaterialInput UCase(shtPlanning.Range("material_id"))
    Logger.Log "Listing Documents . . . "
    Set App.printer = Factory.CreateDocumentPrinter
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.specs
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
    If shtPlanning.Range("console") = nullstr Then
        shtPlanning.Range("console") = "No specifications are available for this code."
    End If
End Function

Public Function ISearchForm_DisplayResults()
    Dim prompt_result As DocumentPackageVariant
    If shtPlanning.Range("console") = nullstr Then
         Prompt.Error "There is nothing to print!"
         Exit Function
    ElseIf shtPlanning.Range("console") = "No specifications are available for this code." Then
         Prompt.Error "There is nothing to print!"
         Exit Function
    ElseIf Not IsNumeric(shtPlanning.Range("work_order")) Then
         Prompt.Error "Please enter a production order."
         Exit Function
    End If
    ' Consider process exceptions based on input from planners.
    ' Eventually this will be obsolete
    prompt_result = Prompt.ProtectionPlanningSequence
    If Not App.TestingMode Then
        ' Check for alternate machine ids (currently only for weaving)
        If App.current_doc.ProcessId = "Weaving" Then
                SpecManager.FilterByMachineId shtPlanning.Range("machine_id")
        End If
        ' Write the documents to their repsective worksheets
        App.printer.WriteAllDocuments shtPlanning.Range("work_order"), prompt_result
        ' Print all of the documents based on the selected doc package and machine
        PrintSelectedPackage prompt_result
    Else
        Logger.Log CStr(prompt_result)
    End If
    App.Shutdown
End Function

Sub PrintSelectedPackage(selected_package As DocumentPackageVariant)
' Prints the select document package for protection

    ' Select document package
    Select Case selected_package
        Case WeavingStyleChange
            Logger.Log "Printing Weaving Style Change Package"
            App.printer.PrintPackage App.specs, selected_package, shtPlanning.Range("work_order")
        Case WeavingTieBack
            Logger.Log "Print Weaving Tie-Back Package"
            App.printer.PrintPackage App.specs, selected_package, shtPlanning.Range("work_order")
        Case FinishingWithQC
            Logger.Log "Printing Finishing with QC Package"
            App.printer.PrintPackage App.specs, selected_package, shtPlanning.Range("work_order")
        Case FinishingNoQC
            Logger.Log "Printing Finishing without QC Package"
            App.printer.PrintPackage DropKeys(App.specs, _
                        Array("Testing Requirements", "Ballistic Testing Requirements")), selected_package, shtPlanning.Range("work_order")
        Case Isotex
            Logger.Log "Printing Isotex TSPP"
            App.printer.PrintPackage App.specs, selected_package, shtPlanning.Range("work_order")
        Case Default
            Logger.Log "Printing All Available Specs"
            Debug.Print IsEmpty(App.specs)
            App.printer.PrintPackage App.specs, selected_package, shtPlanning.Range("work_order")
    End Select

    
End Sub
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    m_Name = "Protection Planning"
    Set m_Sheet = shtPlanning
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
