Attribute VB_Name = "Tests"
Option Explicit

Private pTestResults As VBA.Collection

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

Private Sub ResetTestResults()
    Set pTestResults = New VBA.Collection
End Sub

Sub AllTests()
' End to end testings for the GUI and other hard to test functionality
    On Error GoTo TestSuiteFailed
    ResetTestResults
    App.Start
    With Logger
        .SetLogLevel LOG_TEST
        .SetImmediateLog TestLog
        .Log "------------ Starting Test Suite ----------", TestLog
        Utils.UnloadAllForms
        App.InitializeTestSuiteCredentials
        .Log CreateTemplate_Test, TestLog
        .Log CreateSpecification_Test, TestLog
        .Log ViewSpecification_AfterCreate_Test, TestLog
        .Log EditTemplate_Test, TestLog
        .Log EditSpecification_Test, TestLog
        .Log ViewSpecification_AfterEdit_Test, TestLog
        ' Delete test template
        .Log Utils.FormatTestResult("Delete Template Test", IIf(Tests.DeleteTestTemplate = DB_DELETE_SUCCESS, "PASS", "FAIL")), TestLog
        ' Account Control
        ' TODO: This feature has not been implemented yet.
        App.DeInitializeTestSuiteCredentials
        Utils.UnloadAllForms
        .Log "------------- Test Suite Complete ---------", TestLog
        .Log "Test Suite " & GetTestResults, TestLog
        .SaveAllLogs
        .ResetLog TestLog
        .SetImmediateLog RuntimeLog
        .SetLogLevel LOG_LOW
    End With
    Exit Sub
TestSuiteFailed:
    With Logger
        .Log "Test Suite " & GetTestResults, TestLog
        .SaveAllLogs
        .ResetLog TestLog
        .SetImmediateLog RuntimeLog
        .SetLogLevel LOG_LOW
    End With
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
    ' 4. Enter a machine id for this specification
    formNewSpecInput.txtMachineId = "test_machine"
    ' 5. Click the submit button
    formNewSpecInput.Continue
    ' 6. Select the property "test_property" from the combo box
    formCreateSpec.cboSelectProperty = "test_property"
    ' 7. Enter a value in the txtPropertyValue textbox "Create specification test"
    formCreateSpec.txtPropertyValue = "Create specification test"
    ' 8. Click the set property button
    formCreateSpec.SetProperty
    ' 9. Click the save changes button
    formCreateSpec.SaveChanges
    ' 10. Report pass / fail
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
    ' 4. Select a specification UID
    formViewSpecs.cboSelectType = "test_template(test_machine)"
    ' 5. Report pass / fail
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
    ' 7. Select a property in the property combo box "test_property"
    formEditTemplate.cboSelectProperty = "test_property"
    ' 8. Click the delete property button
    formEditTemplate.RemoveProperty
    ' 9. Save Changes
    formEditTemplate.SaveChanges
    ' 10. Report pass / fail
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
    ' 4. Select a Specification UID
    formSpecConfig.cboSelectType = "test_template(test_machine)"
    ' 5. Select a property name in the combo box "new_test_property"
    formSpecConfig.cboSelectProperty = "new_test_property"
    ' 6. Enter a value in the property value box "Edit specification test"
    formSpecConfig.txtPropertyValue = "Edit specification test"
    ' 7. Click the set property button
    formSpecConfig.Submit
    ' 8. Save changes
    formSpecConfig.SaveChanges
    ' 9. Remove old specification from the archive
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(App.specs.item("to_archive"), "archived_specifications"), SqlLog
    ' 10. Report pass / fail
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
    ' 4. Select a specification UID
    formViewSpecs.cboSelectType = "test_template(test_machine)"
    ' 5. Remove Spec Template
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecificationTemplate(App.current_spec.Template), SqlLog
    ' 6. Remove Spec
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(App.current_spec), SqlLog
    ' 7. Report pass / fail
    ViewSpecification_AfterEdit_Test = Utils.FormatTestResult("View Specification After Edit Test", "PASS")
    pTestResults.Add True
    Exit Function
TestFailed:
    pTestResults.Add False
    ViewSpecification_AfterEdit_Test = Utils.FormatTestResult("Create Specification After Edit Test", "FAIL")
End Function

Sub AccessControl_Test()
    Logger.Log "------------- Start Access Control Test --------------", TestLog
    App.RefreshObjects
    Logger.Log "------------- End Access Control Test ----------------", TestLog
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
    With Logger
        .SetLogLevel LOG_TEST
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
        .SetLogLevel LOG_LOW
    End With
End Sub

Public Sub SqlTransaction_Tests()
    App.Start
    SpecManager.ApplyTemplateChangesToSpecifications "Transaction Test", Array("Change 1", "Change 2")
    Logger.SaveAllLogs
    App.Shutdown
End Sub

Function DeleteTestTemplate() As Long
' Deletes a record
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbDeleteFailException
    Set transaction = Factory.CreateSqlTransaction(DATABASE_PATH)
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM template_specifications " & _
              "WHERE Spec_Type ='test_template' AND Revision ='3.0'"
    transaction.ExecuteSQL (SQLstmt)
    DeleteTestTemplate = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    Logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
    DeleteTestTemplate = DB_DELETE_FAILURE
End Function

Public Sub GetFiles_Test()
    Dim file As Variant
    For Each file In Utils.GetFiles(pfilters:=Array("csv"))
        Debug.Print file
    Next file
End Sub

Function CreateTestSpec() As Specification
    Dim test_spec As Specification
    
    Set test_spec = Factory.CreateSpecification
    With test_spec
        .MaterialId = "TEST_SPECIFICATION"
        .AddProperty "Test Property"
        .SpecType = "TEST"
    End With
    
    Set CreateTestSpec = test_spec
End Function

Public Sub AddNewMaterialDescription_Test()
    Dim test_spec As Specification
    App.Start
    App.InitializeTestSuiteCredentials
    Set test_spec = CreateTestSpec
    Logger.Log SpecManager.SaveNewSpecification(test_spec)
    SpecManager.DeleteSpecification test_spec
    App.DeInitializeTestSuiteCredentials
    App.Shutdown
End Sub

Public Sub SelectAllWhere_Test()
    Dim df As DataFrame
    Dim coll As VBA.Collection
    Dim spec As Specification
    Dim dict As Object
    Set dict = Factory.CreateDictionary
    Set df = DataAccess.SelectAllWhere(Array("Spec_Type"), Array("Testing Requirements"), "standard_specifications")
    Debug.Print df.ToString
End Sub

Public Sub BuildBallisticPackage_Test()
    Dim package As BallisticPackage
    App.Start
    Set package = SpecManager.BuildBallisticTestSpec("test", 15, 63, 173, 0.75)
    App.Shutdown
    With package
        Debug.Print "Target Weight [psf]: " & .TargetPsf
        Debug.Print "Conditioned Weight [gsm] : " & .ConditionedWeight
        Debug.Print "Package Length [in] : " & .PackageLengthInches
        Debug.Print "Actual Weight [psf] : " & .ActualPsf
        Debug.Print "Minimum Sample Length [LY] : " & .MinRequiredSampleLength
        Debug.Print "Package Mass [lbm] : " & .PackageMass
    End With
End Sub

    
