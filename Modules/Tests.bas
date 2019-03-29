Attribute VB_Name = "Tests"

Sub AllTests()
    Logger.Log "----------- Starting Test Suite -----------------"
    Utils.UnloadAllForms
    SpecManager.StopSpecManager
    CreateTemplate_Test
    CreateSpecification_Test
    ViewSpecification_AfterCreate_Test
    EditTemplate_Test
    EditSpecification_Test
    ViewSpecification_AfterEdit_Test
    ' Account Control
    ' TODO: This feature has not been implemented yet.
    Logger.SaveLog "Tests"
    Logger.Log "----------- Test Suite Complete ------------------"
End Sub

Sub CreateTemplate_Test()
    Logger.Log "------------- Start Create Template Test ---------"
    ' 1. Main menu button to create template
    'GuiCommands.GoToMain
    SpecManager.StartSpecManager
    ' 2. Enter a template name "test_template"
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
    SpecManager.StopSpecManager
    Logger.Log "------------- End Create Template Test ------------"
End Sub

Sub CreateSpecification_Test()
    Logger.Log "------------- Start Create Specification Test ---------"
    ' 1. Main menu button to  create specification
    SpecManager.StartSpecManager
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
    SpecManager.StopSpecManager
    Logger.Log "------------- End Create Specification Test ---------"
End Sub

Sub ViewSpecification_AfterCreate_Test()
    Logger.Log "------------- Start View Specification Test ---------"
    ' 1. Main menu button to view specifications
    SpecManager.StartSpecManager
    ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
    formViewSpecs.txtMaterialId = "test_specification"
    ' 3. Click the search button
    formViewSpecs.MaterialSearch
    ' 4. Click the save pdf button but make sure it does not display a file input box and the file is save in a predetermined place that subsequently
    '    noted in the tests.log file. This pdf should not display but it must be check after the tests are complete as a final step to validate
    '    the process. (save as test_specification_rev1.pdf)
    formViewSpecs.ExportPdf
    ' 5. Report pass / fail
    ' TODO: No idea how to do this yet.
    ' 6. Go to main menu
    SpecManager.StopSpecManager
    Logger.Log "------------- End View Specification Test(1) --------"
End Sub

Sub EditTemplate_Test()
    Logger.Log "------------- Start Edit Template Test --------------"
    ' 1. Main menu button to edit template
    SpecManager.StartSpecManager
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
    SpecManager.StopSpecManager
    Logger.Log "------------- End Edit Template Test ----------------"
End Sub

Sub EditSpecification_Test()
    Logger.Log "------------- Start Edit Specification Test ---------"
    ' 1. Main menu button to edit specification
    SpecManager.StartSpecManager
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
    ' 8. Report pass / fail
    ' TODO:
    ' 9. Go to main menu
    SpecManager.StopSpecManager
    Logger.Log "------------- End Edit Specification Test -----------"
End Sub

Sub ViewSpecification_AfterEdit_Test()
    Logger.Log "------------- Start View Specification Test ---------"
    ' 1. Main menu button to view specifications
    SpecManager.StartSpecManager
    ' 2. Enter a material ID txtMaterialId(?) = "test_specification"
    formViewSpecs.txtMaterialId = "test_specification"
    ' 3. Click the search button
    formViewSpecs.MaterialSearch
    ' 4. Click the save pdf button but make sure it does not display a file input box and the file is save in a predetermined place that subsequently
    '    noted in the tests.log file. This pdf should not display but it must be check after the tests are complete as a final step to validate
    '    the process. (save as test_specification_rev1.pdf)
    formViewSpecs.ExportPdf
    ' 5. Report pass / fail
    ' TODO: No idea how to do this yet.
    ' 6. Go to main menu
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(manager.current_spec)
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecification(manager.specs.Item("1.0"))
    Logger.Log "SQLite returned : " & SpecManager.DeleteSpecTemplate(manager.current_template)
    SpecManager.StopSpecManager
    Logger.Log "------------- End View Specification Test(2) ---------"
End Sub
