Attribute VB_Name = "SpecManager"

Public Sub StartApp()
    Logger.Trace "Starting Application"
    GuiCommands.ResetExcelGUI
    App.Start
End Sub

Public Sub RestartApp()
    Logger.Trace "Restarting Application"
    GuiCommands.ResetExcelGUI
    App.RefreshObjects
End Sub

Public Sub StopApp()
    Logger.Trace "Stopping Application"
    Logger.SaveLog
    GuiCommands.ResetExcelGUI
    App.Shutdown
End Sub

Public Sub LoadExistingTemplate(template_type As String)
    With App
        Set .current_template = SpecManager.GetTemplate(template_type)
        .current_template.SpecType = template_type
    End With

End Sub

Function NewSpecificationInput(template_type As String, spec_name As String) As String
    If template_type <> vbNullString Then
        LoadExistingTemplate template_type
        With App
        Set .current_spec = New Specification
        .current_spec.SpecType = .current_template.SpecType
        .current_spec.Revision = "1.0"
        .current_spec.MaterialId = spec_name
        End With
        NewSpecificationInput = spec_name
    Else
        NewSpecificationInput = vbNullString
    End If
End Function

Function TemplateInput(template_type As String) As String
    Set App.current_template = Factory.CreateNewTemplate(template_type)
    TemplateInput = template_type
End Function

Sub MaterialInput(material_id As String)
' Takes user input for material search
    Dim ret_val As Long
    If material_id = vbNullString Then
        ' You must enter a material id before clicking search
        MsgBox "You must enter a material id.", , "Invalid Search Exception"
        Exit Sub
    End If
    ret_val = SpecManager.SearchForSpecifications(material_id)
    If ret_val = SM_SEARCH_FAILURE Then
        ' Let the user know that the specifcation could not be found.
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    ElseIf ret_val = SM_SEARCH_AGAIN Then
        ret_val = SpecManager.SearchForSpecifications(material_id)
        If ret_val = SM_SEARCH_FAILURE Then
            ' Let the user know that the specifcation could not be found.
            MsgBox "Specification not found!", , "Null Spec Exception"
            Exit Sub
        End If
    End If

End Sub

Function SearchForSpecifications(material_id As String) As Long
' Manages the search procedure
    Dim coll As Collection
    Dim specs_dict As Object
    Dim itms
    Set specs_dict = SpecManager.GetSpecifications(material_id)
    If specs_dict Is Nothing Then
        Logger.Log "Could not find a standard for : " & material_id
        SearchForSpecifications = SM_SEARCH_FAILURE
    Else
        Set App.specs = specs_dict
        itms = App.specs.Items
        Set App.current_spec = itms(0)
        Set coll = New Collection
        For Each Key In App.specs
            coll.Add App.specs.Item(Key)
        Next Key
        Logger.Log "Succesfully retrieved specifications for : " & material_id
        ' If SpecManager.UpdateTemplateChanges Then
        '     Logger.Log "Specs updated"
        ' End If
        SearchForSpecifications = SM_SEARCH_SUCCESS
    End If
End Function

Function GetTemplate(template_type As String) As SpecificationTemplate
    Dim record As DatabaseRecord
    Set record = DataAccess.GetTemplateRecord(template_type)
    If Not record Is Nothing Then
        Logger.Log "Succesfully retrieved template for : " & template_type
        Set GetTemplate = Factory.CreateTemplateFromRecord(record)
    Else
        Logger.Log "Could not find a template for : " & template_type
        Set GetTemplate = Nothing
    End If

End Function

Function GetAllTemplates() As VBA.Collection
    Dim record As DatabaseRecord
    Dim dict As Object
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    Set record = DataAccess.GetTemplateTypes
    ' obsoleted
    Logger.Log "Listing all template types (spec Types) . . . "
    For Each dict In record.records
        coll.Add Item:=Factory.CreateTemplateFromDict(dict), Key:=dict.Item("Spec_Type")
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
        Set spec = App.specs.Item(T)
        Set old_spec = Factory.CopySpecification(spec)
        Set App.current_template = GetTemplate(spec.SpecType)
        For Each Key In App.current_template.Properties
            ' Checks for existence current template properites in previous spec
            If Not spec.Properties.Exists(Key) Then
                ' Missing properties are added.
                Logger.Log "Adding : " & Key & " to " & spec.MaterialId & " properties list."
                spec.Properties.Add Key:=Key, Item:=vbNullString
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
                Logger.Log "Data Access returned: " & ret_val
                Logger.Log "New Specification Was Not Saved. Contact Admin."
            Else
                Logger.Log "Data Access returned: " & ret_val
                Logger.Log "New Specification Succesfully Saved."
            End If
        End If
    Next T

    UpdateTemplateChanges = Updated
End Function

Function GetSpecifications(material_id As String) As Object
    Dim json_dict As Object
    Dim specs_dict As Object
    Dim json_coll As VBA.Collection
    Dim spec As Specification
    Dim rev As String
    Dim Key As Variant
    Dim record As DatabaseRecord

    On Error GoTo NullSpecException

    Set record = DataAccess.GetSpecificationRecords(MaterialInputValidation(material_id))
    Set json_coll = record.records
    Set specs_dict = Factory.CreateDictionary
    
    If json_coll.Count = 0 Then
        Set GetSpecifications = Nothing
        Exit Function
    Else
        For Each json_dict In json_coll
            Set spec = Factory.CreateSpecFromDict(json_dict)
            specs_dict.Add json_dict.Item("Spec_Type"), spec
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
    Set App.console = Factory.CreateConsoleBox(frm)
    If Not App.specs Is Nothing Then
        App.console.ListObjects App.specs
    Else
        App.console.PrintLine "No specifications are available for this code."
    End If
End Sub

Sub PrintSpecification(frm As MSForms.UserForm)
    Logger.Log "Printing Specification . . . "
    Set App.console = Factory.CreateConsoleBox(frm)
    If Not App.current_spec Is Nothing Then
        App.console.PrintObject App.current_spec
    End If
End Sub

Sub PrintTemplate(frm As MSForms.UserForm)
    Logger.Log "Printing Template . . . "
    Set App.console = Factory.CreateConsoleBox(frm)
    App.console.PrintObject App.current_template
End Sub

Function SaveNewSpecification(spec As Specification) As Long
    If ManagerOrAdmin Then
        SaveNewSpecification = IIf(DataAccess.PushSpec(spec) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveNewSpecification = DB_PUSH_DENIED
    End If
End Function

Function SaveSpecification(spec As Specification, old_spec As Specification) As Long
    If ManagerOrAdmin Then
        If ArchiveSpecification(old_spec) = DB_DELETE_SUCCESS Then
            SaveSpecification = IIf(DataAccess.PushSpec(spec) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
        Else
            SaveSpecification = DB_PUSH_DENIED
        End If
    Else
        SaveSpecification = DB_PUSH_DENIED
    End If
End Function

Function ArchiveSpecification(old_spec As Specification) As Long
' Archives the last spec in order to make room for the new one.
    Dim ret_val As Long
    ' 1. Insert old version into archived_specifications
    ret_val = IIf(DataAccess.PushSpec(old_spec, "archived_specifications") = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    ' 2. Delete old version from standard_specifications
    If ret_val = DB_PUSH_SUCCESS Then
        ArchiveSpecification = IIf(DeleteSpecification(old_spec) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
    End If
End Function

Function SaveSpecificationTemplate(Template As SpecificationTemplate) As Long
    If ManagerOrAdmin Then
        SaveSpecificationTemplate = IIf(DataAccess.PushTemplate(Template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveSpecificationTemplate = DB_PUSH_DENIED
    End If
End Function

Function UpdateSpecificationTemplate(Template As SpecificationTemplate) As Long
    If ManagerOrAdmin Then
        UpdateSpecificationTemplate = IIf(DataAccess.UpdateTemplate(Template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        UpdateSpecificationTemplate = DB_PUSH_DENIED
    End If
End Function

Function DeleteSpecificationTemplate(Template As SpecificationTemplate) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteSpecificationTemplate = IIf(DataAccess.DeleteTemplate(Template) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
    Else
        DeleteSpecificationTemplate = DB_DELETE_DENIED
    End If
End Function

Function DeleteSpecification(spec As Specification, Optional tbl As String = "standard_specifications") As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteSpecification = IIf(DataAccess.DeleteSpec(spec, tbl) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
    Else
        DeleteSpecification = DB_DELETE_DENIED
    End If
End Function

Private Function ManagerOrAdmin() As Boolean
' Test to see if the current account has the manager privledges.
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        ManagerOrAdmin = True
    Else
        ManagerOrAdmin = False
    End If
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
        Set .current_spec.Tolerances = .current_template.Properties
    End With
End Function

Public Sub DumpAllSpecsToWorksheet(spec_type As String)
    Dim ws As Worksheet
    Dim dicts As Collection
    Dim dict As Object
    Dim props As Variant
    RestartApp
    Application.ScreenUpdating = False
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Utils.CreateNewSheet(spec_type)
    Set dicts = DataAccess.SelectAllSpecifications(spec_type)
    i = 2
    For Each dict In dicts
        Set App.current_spec = Factory.CreateSpecFromDict(dict)
        props = App.current_spec.ToArray
        If i = 2 Then ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).Value = App.current_spec.Header
        ws.Range(Cells(i, 1), Cells(i, ArrayLength(props))).Value = props
        i = i + 1
    Next dict
    ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).Columns.AutoFit
    Application.ScreenUpdating = True
End Sub

Public Sub TableToJson(num_rows As Integer, num_cols As Integer, ws As Worksheet, Optional start_row As Integer = 2, Optional start_col As Integer = 1)
' Create a column at the end of a table and fill it with a json string represent each row.
    Dim dict As Object
    Dim i As Integer
    Dim k As Integer
    Dim json_string As String
    Dim new_spec As Specification
    Dim spec_dict As Object

    With ws
        For i = start_row To num_rows
            Set dict = Factory.CreateDictionary
            Set spec_dict = Factory.CreateDictionary
            For k = start_col To num_cols
                dict.Add .Cells(1, k), .Cells(i, k)
            Next k
            json_string = JsonVBA.ConvertToJson(dict)
            .Cells(i, num_cols + start_col).Value = json_string
            spec_dict.Add "Properties_Json", json_string
            spec_dict.Add "Tolerances_Json", "{}"
            spec_dict.Add "Material_Id", .Cells(i, 1).Value
            spec_dict.Add "Spec_Type", .Cells(i, 2).Value
            spec_dict.Add "Revision", 1
            Set new_spec = Factory.CreateSpecFromDict(spec_dict)
            If DataAccess.PushSpec(new_spec) <> DB_PUSH_SUCCESS Then
                Logger.Log "Error Writing : " & spec_dict.Item("Material_Id")
                Exit Sub
            End If
            
        Next i
        
    End With
End Sub

Public Sub TestTableToJson()
    TableToJson 83, 17, ActiveWorkbook.Sheets("testing"), start_col:=3
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
        json_string = Replace(JsonVBA.ReadJsonFileToString(json_file_path), "NaN", vbNullString)
        ws.Cells(r, 2).Value = json_string
    Next r
End Sub

Sub PrintPackage(doc_package As Object)
' Print specs from the given doc_package (dictionary)
    'Public Sub PrintSheet(ws As Worksheet, Optional FitToPage As Boolean = False)
    Dim spec As Specification
    Dim doc As Variant
    For Each doc In doc_package
        Set spec = doc_package(doc)
        With spec
            Utils.PrintSheet ThisWorkbook.Sheets(.Spec_Type), IIf(.Spec_Type = "Weaving RBA", False, True)
        End With
    Next doc
End Sub