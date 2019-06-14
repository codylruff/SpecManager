Attribute VB_Name = "SpecManager"

Public Sub StartApp()
    Logger.Trace "Starting Application"
    App.Start
End Sub

Public Sub RestartApp()
    Logger.Trace "Restarting Application"
    App.ResetInteractiveObject
End Sub

Public Sub StopApp()
    Logger.Trace "Stopping Application"
    Logger.SaveLog
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
    If material_id = vbNullString Then
        ' You must enter a material id before clicking search
        MsgBox "You must enter a material id.", , "Invalid Search Exception"
        Exit Sub
    End If
    If SpecManager.SearchForSpecifications(material_id) = SM_SEARCH_FAILURE Then
        ' Let the user know that the specifcation could not be found.
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
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
        'SpecManager.UpdateTemplateChanges coll
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

Function ListAllTemplateTypes() As Collection
    Dim record As DatabaseRecord
    Dim dict As Object
    Dim coll As Collection
    Set coll = New Collection
    Set record = DataAccess.GetTemplateTypes
    record.SetDictionary
    Logger.Log "Listing all template types (spec Types) . . . "
    For Each dict In record.records
        coll.Add dict.Item("Spec_Type")
    Next dict
    Set ListAllTemplateTypes = coll
End Function

Sub UpdateTemplateChanges(specifications As Collection)
    ' Apply a specification template to a collection of specifications
    ' this is done in order to apply any changes to a template that
    ' happened since the previous template was changed.
    Dim Key As Variant
    Dim spec As Specification
    Dim template As SpecificationTemplate
    Logger.Log "Applying specifications for any template changes . . ."
    Set App.current_template = GetTemplate(App.current_spec.SpecType)
    For Each Key In App.current_template.Properties
        If Not App.current_spec.Properties.exists(Key) Then
            Logger.Log "Adding : " & Key & " to specification properties list."
            For Each spec In specifications
                spec.Properties.Add Key:=Key, Item:=vbNullString
            Next spec
        End If
    Next Key
    For Each Key In App.current_spec.Properties
        If Not App.current_template.Properties.exists(Key) Then
            For Each spec In specifications
            Logger.Log "Removing : " & Key & " from specification properties list."
                spec.Properties.Remove Key
            Next spec
        End If
    Next Key
    For Each Key In App.specs
        For Each spec In specifications
            If spec.Revision = Key Then
                Set App.specs.Item(Key) = spec
            End If
        Next spec
    Next Key
End Sub

Function GetSpecifications(material_id As String) As Object
    Dim json_dict As Object
    Dim specs_dict As Object
    Dim json_coll As Collection
    Dim spec As Specification
    Dim rev As String
    Dim Key As Variant
    Dim record As DatabaseRecord

    On Error GoTo NullSpecException

    Set record = DataAccess.GetSpecificationRecords(MaterialInputValidation(material_id))
    record.SetDictionary

    Set json_coll = record.records
    Set specs_dict = Factory.CreateDictionary
    
    If json_coll.count = 0 Then
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
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        SaveNewSpecification = IIf(DataAccess.PushSpec(spec) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveNewSpecification = DB_PUSH_DENIED
    End If
End Function

Function SaveSpecification(spec As Specification, old_spec As Specification) As Long
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
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

Function SaveSpecificationTemplate(template As SpecificationTemplate) As Long
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        SaveSpecificationTemplate = IIf(DataAccess.PushTemplate(template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveSpecificationTemplate = DB_PUSH_DENIED
    End If
End Function

Function UpdateSpecificationTemplate(template As SpecificationTemplate) As Long
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        UpdateSpecificationTemplate = IIf(DataAccess.UpdateTemplate(template) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        UpdateSpecificationTemplate = DB_PUSH_DENIED
    End If
End Function

Function DeleteSpecificationTemplate(template As SpecificationTemplate) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteSpecificationTemplate = IIf(DataAccess.DeleteTemplate(template) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
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

Private Function MaterialInputValidation(material_id As String) As String
' Ensures that the material id input by the user is parseable.
    ' PASS
    MaterialInputValidation = material_id
    
End Function

Function SelectLatestSpec() As Specification
    Dim Key As Variant
    For Each Key In App.specs
        If App.specs.Item(Key).IsLatest = True Then
            Set SelectLatestSpec = App.specs.Item(Key)
        End If
    Next Key
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
    ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).columns.AutoFit
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
