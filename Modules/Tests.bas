Attribute VB_Name = "Tests"
Option Explicit

Private pTestResults As New VBA.Collection

Private Function GetTestResults() As String
    Dim result As Variant
    For Each result In pTestResults
        If Not result Then
            GetTestResults = "FAIL"
            Exit Function
        End If
    Next result
    GetTestResults = "PASS"
End Function

Sub AllTests()
' End to end testings for the GUI and other hard to test functionality
    On Error GoTo TestSuiteFailed
    App.Start
    App.logger.SetImmediateLog TestLog
    App.logger.Log "------------ Starting Test Suite ----------", TestLog
    Utils.UnloadAllForms
    App.InitializeTestSuiteCredentials
    App.logger.Log CreateTemplate_Test, TestLog
    App.logger.Log CreateSpecification_Test, TestLog
    App.logger.Log ViewSpecification_AfterCreate_Test, TestLog
    App.logger.Log EditTemplate_Test, TestLog
    App.logger.Log EditSpecification_Test, TestLog
    App.logger.Log ViewSpecification_AfterEdit_Test, TestLog
    ' Account Control
    ' TODO: This feature has not been implemented yet.
    App.DeInitializeTestSuiteCredentials
    Utils.UnloadAllForms
    App.logger.Log "------------- Test Suite Complete ---------", TestLog
    App.logger.Log "Test Suite " & GetTestResults, TestLog
    App.logger.SaveAllLogs
    App.logger.ResetLog TestLog
    App.logger.SetImmediateLog RuntimeLog
    Exit Sub
TestSuiteFailed:
    App.logger.Log "Test Suite " & GetTestResults, TestLog
    App.logger.SaveAllLogs
    App.logger.ResetLog TestLog
    App.logger.SetImmediateLog RuntimeLog
    'PromptHandler.Error "Somethings wrong, please contact the administrator."
End Sub

Function CreateTemplate_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to create template
    'GuiCommands.GoToMain
    ' 2. Enter a template name "test_template"
    formNewTemplateInput.cboProductLine.value = "Test"
    formNewTemplateInput.txtTemplateName.value = "test_template"
    ' 3. Click the Submit button
    formNewTemplateInput.Continue
    ' 4. Change txtPropertyName = "test_property"
    formCreateGeneric.txtPropertyName = "test_property"
    ' 5. Click the Set Property Button
    formCreateGeneric.AddProperty
    ' 6. Click the Save Changes Button
    formCreateGeneric.SubmitTemplate
    ' 7. Report pass / fail
    ' TODO: I am not sure how to report on this other than a crash.
    ' 8. Go to main menu
    CreateTemplate_Test = Utils.FormatTestResult("Create Template Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    CreateTemplate_Test = Utils.FormatTestResult("Create Template Test", "FAIL")
End Function

Function CreateSpecification_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to  create specification
    App.RefreshObjects
    ' 2. Select a template type from the combo box "test_template"
    formNewSpecInput.cboSelectSpecificationType = "test_template"
    ' 3. Enter a material ID "test_specification"
    formNewSpecInput.txtSpecName = "test_specification"
    ' 4. Click the submit button
    formNewSpecInput.Continue
    ' 5. Select the property "test_property" from the combo box
    formCreateSpec.cboSelectProperty = "test_property"
    ' 6. Enter a value in the txtPropertyValue textbox "Create specification test"
    formCreateSpec.txtPropertyValue = "Create specification test"
    ' 7. Click the set property button
    formCreateSpec.SetProperty
    ' 8. Click the save changes button
    formCreateSpec.SaveChanges
    ' 9. Report pass / fail
    ' TODO: No idea how to do this yet
    ' 10. Go to main menu
    CreateSpecification_Test = Utils.FormatTestResult("Create Specification Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    CreateSpecification_Test = Utils.FormatTestResult("Create Specification Test", "FAIL")
End Function

Function ViewSpecification_AfterCreate_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to view specifications
    App.RefreshObjects
    ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
    formViewSpecs.txtMaterialId = "test_specification"
    ' 3. Click the search button
    formViewSpecs.MaterialSearch
    ' 4. Click the save pdf button but make sure it does not display a file input box and the file is save in a predetermined place that subsequently
    '    noted in the tests.log file. This pdf should not display but it must be check after the tests are complete as a final step to validate
    '    the process. (save as test_specification_rev1.pdf)
    'formViewSpecs.ExportPdf
    ' 5. Report pass / fail
    ' TODO: No idea how to do this yet.
    ' 6. Go to main menu
    ViewSpecification_AfterCreate_Test = Utils.FormatTestResult("View Specification Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    ViewSpecification_AfterCreate_Test = Utils.FormatTestResult("View Specification Test", "FAIL")
End Function

Function EditTemplate_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to edit template
    App.RefreshObjects
    ' 2. Select a template name from the combo box
    formEditTemplate.cboSelectTemplate = "test_template"
    ' 3. Click submit
    formEditTemplate.SearchTemplates
    ' 4. Enter a new property into the txtPropertyName box "new_test_property"
    formEditTemplate.txtPropertyName = "new_test_property"
    ' 5. Click the set property button
    formEditTemplate.AddProperty
    ' 6. Save Changes
    formEditTemplate.SaveChanges
    ' 7. Report pass / fail
    ' TODO:
    ' 8. Select a property in the property combo box "test_property"
    formEditTemplate.cboSelectProperty = "test_property"
    ' 9. Click the delete property button
    formEditTemplate.RemoveProperty
    ' 10. Save Changes
    formEditTemplate.SaveChanges
    ' 11. Report pass / fail
    ' TODO:
    ' 12. Go to main menu
    EditTemplate_Test = Utils.FormatTestResult("Edit Template Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    EditTemplate_Test = Utils.FormatTestResult("Edit Template Test", "FAIL")
End Function

Function EditSpecification_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to edit specification
    App.RefreshObjects
    ' 2. Enter a material ID txtSAPcode(?) = "test_specification"
    formSpecConfig.txtMaterialId = "test_specification"
    ' 3. Click the search button
    formSpecConfig.MaterialSearch
    ' 4. Select a property name in the combo box "new_test_property"
    formSpecConfig.cboSelectProperty = "new_test_property"
    ' 5. Enter a value in the property value box "Edit specification test"
    formSpecConfig.txtPropertyValue = "Edit specification test"
    ' 6. Click the set property button
    formSpecConfig.Submit
    ' 7. Save changes
    formSpecConfig.SaveChanges
    ' 8. Remove old specification from the archive
    App.logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(App.specs.item("to_archive"), "archived_specifications"), SqlLog
    ' 9. Report pass / fail
    ' TODO:
    ' 10. Go to main menu
    EditSpecification_Test = Utils.FormatTestResult("Edit Specification Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    EditSpecification_Test = Utils.FormatTestResult("Edit Specification Test", "FAIL")
End Function

Function ViewSpecification_AfterEdit_Test() As String
    On Error GoTo TestFailed
    ' 1. Main menu button to view specifications
    App.RefreshObjects
    ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
    formViewSpecs.txtMaterialId = "test_specification"
    ' 3. Click the search button
    formViewSpecs.MaterialSearch
    ' 4. Click the save pdf button but make sure it does not display a file input box and the file is save in a predetermined place that subsequently
    '    noted in the tests.log file. This pdf should not display but it must be check after the tests are complete as a final step to validate
    '    the process. (save as test_specification_rev1.pdf)
    'formViewSpecs.ExportPdf True
    ' 5. Report pass / fail
    ' TODO: No idea how to do this yet.
    ' 6. Go to main menu
    ' 7. Remove Spec Template
    App.logger.Log "SQLite returned : " & SpecManager.DeleteSpecificationTemplate(App.current_spec.Template), SqlLog
    ' 8. Remove Spec
    App.logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(App.current_spec), SqlLog
    ViewSpecification_AfterEdit_Test = Utils.FormatTestResult("View Specification After Edit Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    ViewSpecification_AfterEdit_Test = Utils.FormatTestResult("Create Specification After Edit Test", "FAIL")
End Function

Sub AccessControl_Test()
    App.logger.Log "------------- Start Access Control Test --------------", TestLog
    App.RefreshObjects
    App.logger.Log "------------- End Access Control Test ----------------", TestLog
End Sub

Public Sub TestKrish()
    Dim ws As Worksheet
    Set ws = shtRBA
    Dim fileName As String
    fileName = PUBLIC_DIR & "\Specifications\" & "Test"
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub

Public Sub ProtectionPlanningSequence_Tests()
    App.Start
    With App.logger
        .SetImmediateLog TestLog
        .Log "Protection Planning Prompt Sequence", TestLog
        PromptHandler.Success "weaving style change"
        .Log IIf(PromptHandler.ProtectionPlanningSequence = WeavingStyleChange, "PASS", "FAIL"), TestLog
        PromptHandler.Success "weaving tie-back"
        .Log IIf(PromptHandler.ProtectionPlanningSequence = WeavingTieBack, "PASS", "FAIL"), TestLog
        PromptHandler.Success "finishing first roll not isotex bound"
        .Log IIf(PromptHandler.ProtectionPlanningSequence = FinishingWithQC, "PASS", "FAIL"), TestLog
        PromptHandler.Success "finishing second roll"
        .Log IIf(PromptHandler.ProtectionPlanningSequence = FinishingNoQC, "PASS", "FAIL"), TestLog
        PromptHandler.Success "finishing first roll + isotex bound"
        .Log IIf(PromptHandler.ProtectionPlanningSequence = FinishingNoQC, "PASS", "FAIL"), TestLog
        .SaveAllLogs
        .SetImmediateLog RuntimeLog
    End With
End Sub
