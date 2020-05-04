Attribute VB_Name = "Planning"

Public Sub cmdPrintSpecifications()
    Dim prompt_result As DocumentPackageVariant
    If shtPlanning.Range("console") = nullstr Then
         PromptHandler.Error "There is nothing to print!"
         Exit Sub
    ElseIf shtPlanning.Range("console") = "No specifications are available for this code." Then
         PromptHandler.Error "There is nothing to print!"
         Exit Sub
    ElseIf Not IsNumeric(shtPlanning.Range("work_order")) Then
         PromptHandler.Error "Please enter a production order."
         Exit Sub
    End If
    ' Consider process exceptions based on input from planners.
    prompt_result = PromptHandler.ProtectionPlanningSequence
    If Not App.TestingMode Then
        ' Check for alternate machine ids (currently only for weaving)
        If App.current_spec.ProcessId = "Weaving" Then
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
End Sub

Public Sub cmdSearch()
    ' Check for any white space and remove it
    App.Start
    If Utils.RemoveWhiteSpace(shtPlanning.Range("material_id")) = nullstr Then
       PromptHandler.Error "Please enter a material id."
       Exit Sub
    ElseIf Utils.RemoveWhiteSpace(shtPlanning.Range("work_order")) = nullstr Then
        PromptHandler.Error "Please enter a work order number."
       Exit Sub
    ElseIf Utils.RemoveWhiteSpace(shtPlanning.Range("machine_id")) = nullstr Then
        PromptHandler.Error "Please enter a machine id."
       Exit Sub
    End If
    ' Run search routine
    MaterialSearch

End Sub

Sub MaterialSearch()
    SpecManager.MaterialInput UCase(shtPlanning.Range("material_id"))
    Logger.Log "Listing Specifications . . . "
    Set App.printer = Factory.CreateDocumentPrinter
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.specs
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
    If shtPlanning.Range("console") = nullstr Then
        shtPlanning.Range("console") = "No specifications are available for this code."
    End If
End Sub

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

