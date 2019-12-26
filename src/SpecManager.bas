Attribute VB_Name = "SpecManager"

Public Sub StartApp()
    App.Start
    GuiCommands.ResetExcelGUI
    Logger.Trace "Starting Application"
    'App.current_user.ListenTo App.printer
End Sub

Public Sub RestartApp()
    Logger.Trace "Restarting Application"
    GuiCommands.ResetExcelGUI
    App.RefreshObjects
End Sub

Public Sub StopApp()
    On Error GoTo ResumeShutdown
    Logger.Trace "Stopping Application"
    Logger.SaveLog
    App.Shutdown
ResumeShutdown:
    GuiCommands.ResetExcelGUI
End Sub

Public Sub LoadExistingTemplate(template_type As String)
    With App
        Set .current_template = SpecManager.GetTemplate(template_type)
        .current_template.SpecType = template_type
    End With

End Sub

Function NewSpecificationInput(template_type As String, spec_name As String, machine_id As String) As String
    If template_type <> nullstr Then
        LoadExistingTemplate template_type
        With App
        Set .current_spec = New Specification
        .current_spec.SpecType = .current_template.SpecType
        .current_spec.Revision = "1.0"
        .current_spec.MaterialId = spec_name
        .current_spec.MachineId = machine_id
        End With
        NewSpecificationInput = spec_name
    Else
        NewSpecificationInput = nullstr
    End If
End Function

Function TemplateInput(template_type As String) As String
    Set App.current_template = Factory.CreateNewTemplate(template_type)
    TemplateInput = template_type
End Function

Sub MaterialInput(material_id As String)
' Takes user input for material search
    Dim ret_val As Long
    If material_id = nullstr Then
        ' You must enter a material id before clicking search
        PromptHandler.Error "Specification not found!"
        Exit Sub
    End If
    ret_val = SpecManager.SearchForSpecifications(material_id)
    If ret_val = SM_SEARCH_FAILURE Then
        ' Let the user know that the specifcation could not be found.
        PromptHandler.Error "Specification not found!"
        Exit Sub
    ElseIf ret_val = SM_SEARCH_AGAIN Then
        ret_val = SpecManager.SearchForSpecifications(material_id)
        If ret_val = SM_SEARCH_FAILURE Then
            ' Let the user know that the specifcation could not be found.
            PromptHandler.Error "Specification not found!"
            Exit Sub
        End If
    End If

End Sub

Function SearchForSpecifications(material_id As String) As Long
' Manages the search procedure
    Dim specs_dict As Object
    Dim itms
    Set specs_dict = SpecManager.GetSpecifications(material_id)
    If specs_dict Is Nothing Then
        Logger.Log "Could not find a specifaction for : " & material_id
        SearchForSpecifications = SM_SEARCH_FAILURE
    Else
        Set App.specs = specs_dict
        itms = App.specs.Items
        Set App.current_spec = itms(0)
        Logger.Log "Succesfully retrieved specifications for : " & material_id
        ' If SpecManager.UpdateTemplateChanges Then
        '     Logger.Log "Specs updated"
        ' End If
        SearchForSpecifications = SM_SEARCH_SUCCESS
    End If
End Function

Function GetTemplate(template_type As String) As SpecificationTemplate
    Dim df As DataFrame
    Set df = DataAccess.GetTemplateRecord(template_type)
    If Not df Is Nothing Then
        Logger.Log "Succesfully retrieved template for : " & template_type
        Set GetTemplate = Factory.CreateTemplateFromRecord(df)
    Else
        Logger.Log "Could not find a template for : " & template_type
        Set GetTemplate = Nothing
    End If

End Function

Function GetAllTemplates() As VBA.Collection
    Dim df As DataFrame
    Dim dict As Object
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    Set df = DataAccess.GetTemplateTypes
    ' obsoleted
    Logger.Log "Listing all template types (spec Types) . . . "
    For Each dict In df.records
        coll.Add item:=Factory.CreateTemplateFromDict(dict), Key:=dict.item("Spec_Type")
    Next dict
    Set GetAllTemplates = coll
End Function

Private Function UpdateTemplateChanges() As Boolean
    ' Apply any changes to material specs that happened since the previous template was revised.
    Dim Key, T As Variant
    Dim ret_val As Long
    Dim Updated As Boolean
    Dim spec As Specification
    Dim Template As SpecificationTemplate
    Dim old_spec As Specification
    Logger.Log "Checking specifications for any template updates . . ."
    For Each T In App.specs
    Updated = False
        Set spec = App.specs.item(T)
        Set old_spec = Factory.CopySpecification(spec)
        Set App.current_template = GetTemplate(spec.SpecType)
        For Each Key In App.current_template.Properties
            ' Checks for existence current template properites in previous spec
            If Not spec.Properties.Exists(Key) Then
                ' Missing properties are added.
                Logger.Log "Adding : " & Key & " to " & spec.MaterialId & " properties list."
                spec.Properties.Add Key:=Key, item:=nullstr
                Updated = True
            End If
        Next Key
        For Each Key In spec.Properties
            ' Checks for existance of current_spec Properties in current_template.
            If Not App.current_template.Properties.Exists(Key) Then
                ' Old properties are removed
                Logger.Log "Removing : " & Key & " from " & spec.MaterialId & " properties list."
                spec.Properties.Remove Key
                Updated = True
            End If
        Next Key
        If Updated = True Then
            spec.Revision = CStr(CDbl(spec.Revision) + 1#)
            ret_val = SpecManager.SaveSpecification(spec, old_spec)
            If ret_val <> DB_PUSH_SUCCESS Then
                Logger.Log "Data Access returned: " & ret_val, DebugLog
                Logger.Log "New Specification Was Not Saved. Contact Admin."
            Else
                Logger.Log "Data Access returned: " & ret_val, DebugLog
                Logger.Log "New Specification Succesfully Saved."
            End If
        End If
    Next T

    UpdateTemplateChanges = Updated
End Function

Function GetSpecifications(material_id As String) As Object
    Dim json_dict As Object
    Dim specs_dict As Object
    Dim spec As Specification
    Dim rev As String
    Dim Key As Variant
    Dim df As DataFrame

    On Error GoTo NullSpecException

    Set df = DataAccess.GetSpecificationRecords(MaterialInputValidation(material_id))
    Set specs_dict = Factory.CreateDictionary
    
    If df.records.Count = 0 Then
        Set GetSpecifications = Nothing
        Exit Function
    Else
        For Each json_dict In df.records
            Set spec = Factory.CreateSpecFromDict(json_dict)
            specs_dict.Add spec.UID, spec
        Next json_dict
        Set GetSpecifications = specs_dict
    End If
    Exit Function
NullSpecException:
    Logger.Error "SpecManager.GetSpecifications()"
    Set GetSpecifications = Nothing
End Function

Sub ListSpecifications(frm As MSForms.UserForm)
' Lists the specifications currently selected in txtConsole for the given form
    Logger.Log "Listing Specifications . . . "
    Set App.printer = Factory.CreateDocumentPrinter(frm)
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.specs
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
End Sub

Sub PrintSpecification(frm As MSForms.UserForm)
    Logger.Log "Writing Specification to Console. . . "
    Set App.printer.FormId = frm
    If Not App.current_spec Is Nothing Then
        App.printer.PrintObjectToConsole App.current_spec
    End If
End Sub

Sub PrintTemplate(frm As MSForms.UserForm)
    Logger.Log "Writing Template to Console . . . "
    Set App.printer.FormId = frm
    App.printer.PrintObjectToConsole App.current_template
End Sub

Public Sub UpdateSingleProperty(property_name As String, property_value As Variant, material_id As String)
' Updates the value of a single property without the use of the UI. This should make Admin easier.
End Sub

Public Sub ApplyTemplateChangesToSpecifications(spec_type As String, changes As Variant)
' Apply template changes to all existing specs of that type
    Dim specifications As VBA.Collection
    Dim spec As Specification
    Dim old_spec As Specification
    Dim i As Long
    Dim transaction As SqlTransaction
    Set specifications = SelectAllSpecificationsByType(spec_type)
    Set transaction = DataAccess.BeginTransaction
    For Each spec In specifications
        Set old_spec = Factory.CopySpecification(spec)
        For i = LBound(changes) To UBound(changes)
            spec.AddProperty CStr(changes(i))
        Next i
        spec.Revision = CStr(CDbl(old_spec.Revision) + 1)
        Logger.Log SpecManager.SaveSpecification(spec, old_spec, transaction), DebugLog
    Next spec
    'Logger.Log transaction.Commit, DebugLog
End Sub

Private Function SelectAllSpecificationsByType(spec_type As String) As VBA.Collection
    Dim record_coll As VBA.Collection
    Dim record_dict As Object
    Dim specifications As New VBA.Collection
    Set record_dict = Factory.CreateDictionary
    Set record_coll = DataAccess.SelectAllSpecifications(spec_type)
    For Each record_dict In record_coll
        specifications.Add Factory.CreateSpecFromDict(record_dict)
    Next record_dict
    Set SelectAllSpecificationsByType = specifications
End Function

Function CreateSpecificationFromCopy(spec As Specification, material_id As String) As Long
' Takes a material and makes a copy of it under a new material id
    Dim spec_copy As Specification
    Set spec_copy = Factory.CopySpecification(spec)
    spec_copy.MaterialId = material_id
    spec_copy.Revision = 1
    CreateSpecificationFromCopy = SaveNewSpecification(spec_copy)
End Function

Function GetMaterialDescription(material_id As String) As Variant
' Retrieve material description from the database.
    GetMaterialDescription = DataAccess.GetColumn("Material_Id", material_id, "Description", "materials").Data
End Function

Function AddNewMaterialDescription(material_id As String, description As String, process_id As String) As Long
' If a material description does not exist create it.
    Dim ret_val As Long
    If GetMaterialDescription(material_id) = Empty Then
        ret_val = DataAccess.PushValue("Material_Id", material_id, "Description", description, "materials")
        AddNewMaterialDescription = DataAccess.UpdateValue("Material_Id", material_id, "Process_Id", process_id, "materials")
    Else
        AddNewMaterialDescription = SM_MATERIAL_EXISTS
    End If
End Function

Function SaveNewSpecification(spec As Specification, Optional material_description As String) As Long
    Dim ret_val As Long
    If ManagerOrAdmin Then
        If DataAccess.GetSpecification(spec.MaterialId, spec.SpecType, spec.MachineId).records.Count = 0 Then
            ret_val = iif(DataAccess.PushIQueryable(spec, "standard_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
            ActionLog.CrudOnSpecification spec, "Created New Specification"
            If IsEmpty(GetMaterialDescription(spec.MaterialId)) Then
                ' Use material_description param or prompt the user to enter one.
                If material_description = nullstr Then
                    material_description = CStr(PromptHandler.UserInput(SingleLineText, "Material Description: " & spec.MaterialId, _
                                "Enter Material Description :"))
                End If
                ' Add the new Material Description to the materials table.
                SaveNewSpecification = AddNewMaterialDescription(spec.MaterialId, material_description, spec.ProcessId)
                ActionLog.CrudOnSpecification spec, "Created New Material"
                Exit Function
            End If
            SaveNewSpecification = ret_val
        Else
            SaveNewSpecification = SM_MATERIAL_EXISTS
        End If
    Else
        SaveNewSpecification = DB_PUSH_DENIED
    End If

End Function

Function SaveSpecification(spec As Specification, old_spec As Specification, Optional transaction As SqlTransaction) As Long
    If ManagerOrAdmin Then
        If Utils.IsNothing(transaction) Then
            If ArchiveSpecification(old_spec) = DB_DELETE_SUCCESS Then
                SaveSpecification = iif(DataAccess.PushIQueryable(spec, "standard_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
            Else
                SaveSpecification = DB_PUSH_DENIED
            End If
        Else
            If ArchiveSpecification(old_spec, transaction) = DB_DELETE_SUCCESS Then
                SaveSpecification = iif(DataAccess.PushIQueryable(spec, "standard_specifications", transaction) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
            Else
                SaveSpecification = DB_PUSH_DENIED
            End If
        End If
    Else
        SaveSpecification = DB_PUSH_DENIED
    End If
    ActionLog.CrudOnSpecification spec, "Revised Specification"
End Function

Function ArchiveSpecification(old_spec As Specification, Optional transaction As SqlTransaction) As Long
' Archives the last spec in order to make room for the new one.
    Dim ret_val As Long
    If Utils.IsNothing(transaction) Then
        ' 1. Insert old version into archived_specifications
        ret_val = iif(DataAccess.PushIQueryable(old_spec, "archived_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
        ' 2. Delete old version from standard_specifications
        If ret_val = DB_PUSH_SUCCESS Then
            ArchiveSpecification = iif(DeleteSpecification(old_spec) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
        End If
    Else
        ' 1. Insert old version into archived_specifications
        ret_val = iif(DataAccess.PushIQueryable(old_spec, "archived_specifications", transaction) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
        ' 2. Delete old version from standard_specifications
        If ret_val = DB_PUSH_SUCCESS Then
            ArchiveSpecification = iif(DeleteSpecification(old_spec, "standard_specifications", transaction) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
        End If
    End If
    'ActionLog.CrudOnSpecification old_spec, "Archived Specification"
End Function

Function SaveSpecificationTemplate(Template As SpecificationTemplate) As Long
    If ManagerOrAdmin Then
        SaveSpecificationTemplate = iif(DataAccess.PushIQueryable(Template, "template_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveSpecificationTemplate = DB_PUSH_DENIED
    End If
    ActionLog.CrudOnTemplate Template, "Created New Template"
End Function

Function UpdateSpecificationTemplate(Template As SpecificationTemplate) As Long
    If ManagerOrAdmin Then
        UpdateSpecificationTemplate = iif(DataAccess.UpdateTemplate(Template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        UpdateSpecificationTemplate = DB_PUSH_DENIED
    End If
    ActionLog.CrudOnTemplate Template, "Revised Template"
End Function

Function DeleteSpecificationTemplate(Template As SpecificationTemplate) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteSpecificationTemplate = iif(DataAccess.DeleteTemplate(Template) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
    Else
        DeleteSpecificationTemplate = DB_DELETE_DENIED
    End If
    ActionLog.CrudOnTemplate Template, "Deleted Template"
End Function

Function DeleteSpecification(spec As Specification, Optional tbl As String = "standard_specifications", Optional trans As SqlTransaction) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        If IsNothing(trans) Then
            DeleteSpecification = iif(DataAccess.DeleteSpec(spec, spec.MachineId, tbl) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
        Else
            DeleteSpecification = iif(DataAccess.DeleteSpec(spec, spec.MachineId, tbl, trans) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
        End If
    Else
        DeleteSpecification = DB_DELETE_DENIED
    End If
    ActionLog.CrudOnSpecification spec, "Deleted Specification"
End Function

Private Function ManagerOrAdmin() As Boolean
' Test to see if the current account has the manager privledges.
    On Error GoTo ErrorHandler
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        ManagerOrAdmin = True
    Else
        ManagerOrAdmin = False
    End If
    ManagerOrAdmin = True
    Exit Function
ErrorHandler:
    Dim Account As Account
    Set Account = AccessControl.Account_Initialize
    ManagerOrAdmin = iif(Account.ProductLine = "Admin", True, False)
End Function

Private Function MaterialInputValidation(material_id As String) As String
' Ensures that the material id input by the user is parseable.
    ' PASS
    MaterialInputValidation = material_id
    
End Function

Function InitializeNewSpecification()
    With App
        Set App.current_spec = New Specification
        .current_spec.SpecType = .current_template.SpecType
        .current_spec.Revision = "1.0"
        Set .current_spec.Properties = .current_template.Properties
    End With
End Function

Public Sub DumpAllSpecsToWorksheet(spec_type As String)
    Dim ws As Worksheet
    Dim dicts As Collection
    Dim dict As Object
    Dim props As Variant
    'RestartApp
    
    ' Turn on Performance Mode
    App.PerformanceMode True

    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Utils.CreateNewSheet(spec_type & " Dump " & Format(CStr(Now()), "dd-mm-yy"), True)
    Set dicts = DataAccess.SelectAllSpecifications(spec_type)
    i = 2
    For Each dict In dicts
        Set App.current_spec = Factory.CreateSpecFromDict(dict)
        props = App.current_spec.ToArray
        If i = 2 Then ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).value = App.current_spec.header
        ws.Range(Cells(i, 1), Cells(i, ArrayLength(props))).value = props
        i = i + 1
    Next dict
    ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).Columns.AutoFit
    
    ' Turn off Performance Mode
    App.PerformanceMode False

End Sub

Public Sub MassCreateSpecifications(num_rows As Integer, num_cols As Integer, ws As Worksheet, Optional start_row As Integer = 2, Optional start_col As Integer = 1, Optional print_json_column As Boolean = True, Optional write_to_live As Boolean = False)
' Create a column at the end of a table and fill it with a json string represent each row.
    Dim dict As Object
    Dim i, k As Integer
    Dim json_string As String
    Dim new_spec As Specification
    Dim spec_dict As Object
    App.Start
    With ws
        For i = start_row To num_rows - start_row + 1
            Set dict = Factory.CreateDictionary
            Set spec_dict = Factory.CreateDictionary
            For k = start_col To num_cols
                dict.Add .Cells(1, k), .Cells(i, k)
            Next k
            json_string = JsonVBA.ConvertToJson(dict)
            ' If requested, print the json string in a new column.
            If print_json_column Then
                .Cells(i, num_cols + start_col).value = json_string
            End If
            If write_to_live Then
                spec_dict.Add "Properties_Json", json_string
                spec_dict.Add "Material_Id", .Cells(i, 1).value
                spec_dict.Add "Spec_Type", .Cells(i, 2).value
                spec_dict.Add "Revision", 1
                Set new_spec = Factory.CreateSpecFromDict(spec_dict)
                ret_val = SpecManager.SaveNewSpecification(new_spec, .Cells(i, 3))
                If ret_val = DB_PUSH_SUCCESS Then
                    Logger.Log new_spec.GetName & " Created."
                    ActionLog.CrudOnSpecification new_spec, "Created New Specification"
                ElseIf ret_val = SM_MATERIAL_EXISTS Then
                    Logger.Log new_spec.GetName & " Already Exists."
                Else
                    Logger.Log new_spec.GetName & " Was Not Saved."
                End If
            End If
        Next i
        
    End With
    App.Shutdown
End Sub

Public Sub ParseSpecsTable(ws_name As String, table_name As String, Optional print_json_column As Boolean = True, Optional write_to_live As Boolean = False)
' Converts each row in the table to json format, then loads it into the specs db
    Dim tbl As Table

    Set tbl = Factory.CreateTable(ActiveWorkbook.Sheets(ws_name), table_name)
    ' Validate table column headers
    If tbl.HeaderRowRange(1) <> "Material_Id" Then ' The first column must be the material_id
        Logger.Log "The first column must be the 'Material_Id'"
    ElseIf tbl.HeaderRowRange(2) <> "Spec_Type" Then ' The second column must be the spec_type
        Logger.Log "The second column must be the 'Spec_Type'"
    ElseIf tbl.HeaderRowRange(3) <> "Description" Then ' The third column must be the material_description
        Logger.Log "The second column must be the 'Description'"
    Else
        MassCreateSpecifications num_rows:=tbl.Rows.Count, _
                    num_cols:=CInt(tbl.Columns.Count), _
                    ws:=tbl.Worksheet, _
                    start_col:=4, _
                    print_json_column:=print_json_column, _
                    write_to_live:=write_to_live
    End If

End Sub

Public Sub CopyPropertiesFromFile()
    ' Get range of material ids
    Dim ws As Worksheet
    Dim style_number As String
    Dim json_string As String
    Dim json_file_path As String
    Dim r As Long
    Set ws = Sheet4
    For r = 2 To 58
        style_number = Mid(ws.Cells(r, 1), 6, 3)
        json_file_path = ThisWorkbook.path & "\RBAs\" & style_number & ".json"
        json_string = Replace(JsonVBA.ReadJsonFileToString(json_file_path), "NaN", nullstr)
        ws.Cells(r, 2).value = json_string
    Next r
End Sub

'Private Sub UpdateRBAFromSheet()
'' ****THIS IS BROKEN****
'' This routine will update a specificaiton in the database
'' with the parameters entered into the sheet.
'    Dim material_id As String
'    Dim spec_type As String
'    Dim spec As Specification
'    Dim old_spec As Specification
'    Dim prop As Variant
'    Dim T As SpecificationTemplate
'    App.Start
'    ' Start up spec-manager
'    material_id = CStr(Range("article_code").value)
'    spec_type = "Weaving RBA"
'    Set spec = Factory.CreateSpecificationFromRecord(DataAccess.GetSpecification(material_id, spec_type))
'    Set props = CreateObject("Scripting.Dictionary")
'    ' Get spec by material_id and make a copy
'    Set old_spec = Factory.CopySpecification(spec)
'    ' Loop through named ranges and create a dictionary
'    spec.Revision = CStr(Range("revision").value + 1)
'    For Each prop In spec.Properties
'        spec.Properties(CStr(prop)) = Range(prop).value
'        Logger.Log "Set : " & CStr(prop) & " = " & spec.Properties(CStr(prop))
'    Next prop
'    ' Update the specification
'    Logger.Log "Data Access Returned : " & SaveSpecification(spec, old_spec), DebugLog
'    App.Shutdown
'    AccessControl.ConfigControl
'End Sub

Public Function BuildBallisticTestSpec(material_id As String, package_length_inches As Double, fabric_width_inches As Double, conditioned_weight_gsm As Double, target_psf As Double) As BallisticPackage
    Dim spec As Specification
    Dim package As BallisticPackage
    Set package = Factory.CreateBallisticPackage(package_length_inches, fabric_width_inches, conditioned_weight_gsm, target_psf)
    Set spec = Factory.CreateSpecification
    With spec
        .MaterialId = material_id
        .SpecType = "Ballistic Testing Requirements"
        .Revision = 1
        .CreateFromTestingPlan package
    End With
    'SaveNewSpecification spec
    Set BuildBallisticTestSpec = package
End Function
