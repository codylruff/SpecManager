Attribute VB_Name = "Tests"
Option Explicit
'TODO Change all of this to Tim Halls test library. vba-test
'NOTE How can this be converted to vba-test?
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
   GUI.Start
   With Logger
       .SetLogLevel LOG_TEST
       .SetImmediateLog TestLog
       .Log "------------ Starting Test Suite ----------", TestLog
       App.InitializeTestSuiteCredentials
       .Log CreateTemplate_Test, TestLog
       .Log CreateDocument_Test, TestLog
       .Log ViewDocument_AfterCreate_Test, TestLog
       .Log EditTemplate_Test, TestLog
       .Log EditDocument_Test, TestLog
       '.Log ViewDocument_AfterEdit_Test, TestLog
       ' Delete test template
       .Log Utils.FormatTestResult("Delete Template Test", IIf(Tests.DeleteTestTemplate = DB_DELETE_SUCCESS, "PASS", "FAIL")), TestLog
       ' Account Control
       ' TODO This feature has not been implemented yet.
   End With
Finally:
   App.DeInitializeTestSuiteCredentials
   Logger.Log "------------- Test Suite Complete ---------", TestLog
   Logger.Log "Test Suite " & GetTestResults, TestLog
   Logger.SaveAllLogs
   Logger.ResetLog TestLog
   Logger.SetImmediateLog RuntimeLog
   Logger.SetLogLevel LOG_LOW
   Exit Sub
TestSuiteFailed:
   Prompt.Error "Somethings wrong, please contact the administrator."
   GoTo Finally
End Sub

Function CreateTemplate_Test() As String
On Error GoTo TestFailed
   With GUI.GetForm("FormCreate")
   ' 1.Select Design Type "Template"
       .ChangeFieldValue "design_class", "Template"
   '  Click the submit button
      .CmdByName "Continue"
   ' 2. Enter a template name "test_template"
      .ChangeFieldValue "product_line", "Test"
      .ChangeFieldValue "design_id", "test_template"
   ' 4. Change txtPropertyName = "test_property"
      .ChangeFieldValue "property_id", "test_property"
   ' 6. Click the Save Changes Button
       .CmdByName "SaveChanges"
   End With
   ' 7. Report pass / fail
   CreateTemplate_Test = Utils.FormatTestResult("Create Template Test", "PASS")
   pTestResults.Add True
Finally:
   Exit Function
TestFailed:
   pTestResults.Add False
   CreateTemplate_Test = Utils.FormatTestResult("Create Template Test", "FAIL")
   GoTo Finally
End Function

Function CreateDocument_Test() As String
' NOTE This should be handled by FormCreate.cls
On Error GoTo TestFailed
   App.RefreshObjects
   With GUI.GetForm("FormCreate")
   ' 1. Select Design Type "Document"
       .ChangeFieldValue "design_class", "Document"
   ' 2. Select a template type from the combo box "test_template"
       .ChangeFieldValue "design_type", "test_template"
   ' 5. Click the submit button
       .CmdByName "Continue"
   ' 3. Enter a material ID "test_specification"
       .ChangeFieldValue "design_id", "test_document"
   ' 4. Enter a machine id for this specification
       .ChangeFieldValue "machine_id", "test_machine"
   ' 6. Select the property "test_property" from the combo box
       .ChangeFieldValue "property_id", "test_property"
   ' 7. Enter a value in the txtPropertyValue textbox "Create specification test"
       .ChangeFieldValue "property_value", "Create ocument Test"
   ' 8. Click the set property button
       .CmdByName "SetProperty"
   ' 9. Click the save changes button
       .CmdByName "SaveChanges"
   End With
   ' 10. Report pass / fail
   CreateDocument_Test = Utils.FormatTestResult("Create Document Test", "PASS")
   pTestResults.Add True
Finally:
   Exit Function
TestFailed:
   pTestResults.Add False
   CreateDocument_Test = Utils.FormatTestResult("Create Document Test", "FAIL")
   GoTo Finally
End Function

Function ViewDocument_AfterCreate_Test() As String
' NOTE This should be handled by FormView.cls
On Error GoTo TestFailed
   App.RefreshObjects
   With GUI.GetForm("FormView")
       ' 1. Select Design Type "Document"
       .ChangeFieldValue "design_class", "Document"
       ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
       .ChangeFieldValue "design_id", "test_document"
       ' 3. Click the search button
       .CmdByName "SearchDesigns"
       ' 4. Select a specification UID
       .ChangeFieldValue "select_design", "test_template(test_machine)"
   End With
   ' 5. Report pass / fail
   ViewDocument_AfterCreate_Test = Utils.FormatTestResult("View Document Test", "PASS")
   pTestResults.Add True
Finally:
   Exit Function
TestFailed:
   pTestResults.Add False
   ViewDocument_AfterCreate_Test = Utils.FormatTestResult("View Document Test", "FAIL")
   GoTo Finally
End Function

Function EditTemplate_Test() As String
' NOTE This should be handled by FormEdit
On Error GoTo TestFailed
   App.RefreshObjects
   With GUI.GetForm("FormEdit")
       ' 1. Select Design Type "Template"
       .ChangeFieldValue "design_class", "Template"
       ' 2. Select a template name from the combo box
       .ChangeFieldValue "design_type", "test_template"
       ' 3. Click submit
       .CmdByName "SearchDesigns"
       ' 4. Enter a new property into the txtPropertyName box "new_test_property"
       .ChangeFieldValue "property_id", "new_test_property"
       ' 5. Click the set property button
       .CmdByName "AddProperty"
       ' 6. Save Changes
       .CmdByName "SaveChanges"
       ' 7. Select a property in the property combo box "test_property"
       .ChangeFieldValue "property_id", "test_property"
       ' 8. Click the delete property button
       .CmdByName "RemoveProperty"
       ' 9. Save Changes
       .CmdByName "SaveChanges"
   End With
   ' 10. Report pass / fail
   EditTemplate_Test = Utils.FormatTestResult("Edit Template Test", "PASS")
   pTestResults.Add True

TestFailed:
   pTestResults.Add False
   EditTemplate_Test = Utils.FormatTestResult("Edit Template Test", "FAIL")
   GoTo Finally
Finally:
   Exit Function
End Function

Function EditDocument_Test() As String
' NOTE This should be handled by FormEdit.cls
On Error GoTo TestFailed
   App.RefreshObjects
   With GUI.GetForm("FormEdit")
       ' 1. Select Design Type "Document"
       .ChangeFieldValue "design_class", "Document"
       ' 2. Enter a material ID txtSAPcode(?) = "test_specification"
       .ChangeFieldValue "design_id", "test_document"
       ' 3. Click the search button
       .CmdByName "SearchDesigns"
       ' 4. Select a Document UID
       .ChangeFieldValue "design_type", "test_template(test_machine)"
       ' 5. Select a property name in the combo box "new_test_property"
       .ChangeFieldValue "property_id", "new_test_property"
       ' 6. Enter a value in the property value box "Edit specification test"
       .ChangeFieldValue "property_value", "Edit Document Test"
       ' 7. Click the set property button
       .CmdByName "Submit"
       ' 8. Save changes
       .CmdByName "SaveChanges"
   End With
   ' 9. Remove old specification from the archive
   Logger.Log "SQLite returned : " & SpecManager.DeleteDocument(App.specs.item("to_archive"), "archived_specifications"), SqlLog
   ' 10. Report pass / fail
   EditDocument_Test = Utils.FormatTestResult("Edit Document Test", "PASS")
   pTestResults.Add True
Finally:
   Exit Function
TestFailed:
   pTestResults.Add False
   EditDocument_Test = Utils.FormatTestResult("Edit Document Test", "FAIL")
   GoTo Finally
End Function

'Function ViewDocument_AfterEdit_Test() As String
'' NOTE This should be handled by FormView.cls
'On Error GoTo TestFailed
'   App.RefreshObjects
'   With GUI.GetForm("FormView")
'       ' 1. Select Design Type "Document"
'       .cboDesignType.value = "Document"
'       ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
'       .txtMaterialId = "test_specification"
'       ' 3. Click the search button
'       .MaterialSearch
'       ' 4. Select a specification UID
'       .cboSelectType = "test_template(test_machine)"
'   End With
'   ' 5. Remove Spec Template
'   Logger.Log "SQLite returned : " & SpecManager.DeleteTemplate(App.current_doc.Template), SqlLog
'   ' 6. Remove Spec
'   Logger.Log "SQLite returned : " & SpecManager.DeleteDocument(App.current_doc), SqlLog
'   ' 7. Report pass / fail
'   ViewDocument_AfterEdit_Test = Utils.FormatTestResult("View Document After Edit Test", "PASS")
'   pTestResults.Add True
'Finally:
'   Exit Function
'TestFailed:
'   pTestResults.Add False
'   ViewDocument_AfterEdit_Test = Utils.FormatTestResult("View Document After Edit Test", "FAIL")
'   GoTo Finally
'End Function

' MISC TESTING ROUTINES
'Sub AccessControl_Test()
'   Logger.Log "------------- Start Access Control Test --------------", TestLog
'   App.RefreshObjects
'   Logger.Log "------------- End Access Control Test ----------------", TestLog
'End Sub
'
'Public Sub TestKrish()
'   Dim ws As Worksheet
'   Set ws = shtRBA
'   Dim fileName As String
'   fileName = PUBLIC_DIR & "\Specifications\" & "Test"
'   ws.ExportAsFixedFormat _
'       Type:=xlTypePDF, _
'       fileName:=fileName, _
'       Quality:=xlQualityStandard, _
'       IncludeDocProperties:=True, _
'       IgnorePrintAreas:=False, _
'       OpenAfterPublish:=True
'End Sub
'
'Public Sub ProtectionPlanningSequence_Tests()
'   App.Start
'   With Logger
'       .SetLogLevel LOG_TEST
'       .SetImmediateLog TestLog
'       .Log "Protection Planning Prompt Sequence", TestLog
'       Prompt.Success "weaving style change"
'       .Log IIf(Prompt.ProtectionPlanningSequence = WeavingStyleChange, "PASS", "FAIL"), TestLog
'       Prompt.Success "weaving tie-back"
'       .Log IIf(Prompt.ProtectionPlanningSequence = WeavingTieBack, "PASS", "FAIL"), TestLog
'       Prompt.Success "finishing first roll not isotex bound"
'       .Log IIf(Prompt.ProtectionPlanningSequence = FinishingWithQC, "PASS", "FAIL"), TestLog
'       Prompt.Success "finishing second roll"
'       .Log IIf(Prompt.ProtectionPlanningSequence = FinishingNoQC, "PASS", "FAIL"), TestLog
'       Prompt.Success "finishing first roll + isotex bound"
'       .Log IIf(Prompt.ProtectionPlanningSequence = FinishingNoQC, "PASS", "FAIL"), TestLog
'       .SaveAllLogs
'       .SetImmediateLog RuntimeLog
'       .SetLogLevel LOG_LOW
'   End With
'End Sub
'
'Public Sub SqlTransaction_Tests()
'   App.Start
'   SpecManager.ApplyTemplateChangesToDocuments "Transaction Test", Array("Change 1", "Change 2")
'   Logger.SaveAllLogs
'   App.Shutdown
'End Sub
'
Function DeleteTestTemplate() As Long
' Deletes a record
   Dim SQLstmt As String
   Dim transaction As SqlTransaction
On Error GoTo Catch
   Set transaction = Factory.CreateSqlTransaction(DATABASE_PATH)
   ' Create SQL statement from objects
   SQLstmt = "DELETE FROM template_specifications " & _
             "WHERE Spec_Type ='test_template' AND Revision ='3.0'"
   transaction.ExecuteSQL (SQLstmt)
   DeleteTestTemplate = DB_DELETE_SUCCESS
Finally:
   Set transaction = Nothing
   Exit Function
Catch:
   Logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
   DeleteTestTemplate = DB_DELETE_ERR
   GoTo Finally
End Function
'
'Public Sub GetFiles_Test()
'   Dim file As Variant
'   For Each file In Utils.GetFiles(pfilters:=Array("csv"))
'       Debug.Print file
'   Next file
'End Sub
'
'Function CreateTestSpec() As Document
'   Dim test_spec As Document
'
'   Set test_spec = Factory.CreateDocument
'   With test_spec
'       .MaterialId = "TEST_SPECIFICATION"
'       .AddProperty "Test Property"
'       .SpecType = "TEST"
'   End With
'
'   Set CreateTestSpec = test_spec
'End Function
'
'Public Sub AddNewMaterialDescription_Test()
'   Dim test_spec As Document
'   App.Start
'   App.InitializeTestSuiteCredentials
'   Set test_spec = CreateTestSpec
'   Logger.Log SpecManager.SaveNewDocument(test_spec)
'   SpecManager.DeleteDocument test_spec
'   App.DeInitializeTestSuiteCredentials
'   App.Shutdown
'End Sub
'
'Public Sub SelectAllWhere_Test()
'   Dim df As DataFrame
'   Dim coll As VBA.Collection
'   Dim doc As Document
'   Dim dict As Object
'   Set dict = Factory.CreateDictionary
'   Set df = DataAccess.SelectAllWhere(Array("Spec_Type"), Array("Testing Requirements"), "standard_specifications")
'   Debug.Print df.ToString
'End Sub

    
