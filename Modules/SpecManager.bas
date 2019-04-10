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
    If material_id = vbNullString Then Exit Sub
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
    Set specs_dict = SpecManager.GetSpecifications(material_id)
    If specs_dict Is Nothing Then
        Logger.Log "Could not find a standard for : " & material_id
        SearchForSpecifications = SM_SEARCH_FAILURE
    Else
        Set App.specs = specs_dict
        Set App.current_spec = SelectLatestSpec()
        Set coll = New Collection
        For Each key In App.specs
            coll.Add App.specs.Item(key)
        Next key
        Logger.Log "Succesfully retrieved specifications for : " & material_id
        SpecManager.UpdateTemplateChanges coll
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
    Dim key As Variant
    Dim spec As Specification
    Dim template As SpecificationTemplate
    Logger.Log "Applying specifications for any template changes . . ."
    Set App.current_template = GetTemplate(App.current_spec.SpecType)
    For Each key In App.current_template.Properties
        If Not App.current_spec.Properties.exists(key) Then
            Logger.Log "Adding : " & key & " to specification properties list."
            For Each spec In specifications
                spec.Properties.Add key:=key, Item:=vbNullString
            Next spec
        End If
    Next key
    For Each key In App.current_spec.Properties
        If Not App.current_template.Properties.exists(key) Then
            For Each spec In specifications
            Logger.Log "Removing : " & key & " from specification properties list."
                spec.Properties.Remove key
            Next spec
        End If
    Next key
    For Each key In App.specs
        For Each spec In specifications
            If spec.Revision = key Then
                Set App.specs.Item(key) = spec
            End If
        Next spec
    Next key
End Sub

Function GetSpecifications(material_id As String) As Object
    Dim json_dict As Object
    Dim specs_dict As Object
    Dim json_coll As Collection
    Dim spec As Specification
    Dim rev As String
    Dim key As Variant
    Dim record As DatabaseRecord
    On Error GoTo NullSpecException
    Set record = DataAccess.GetSpecificationRecords(MaterialInputValidation(material_id))
    record.SetDictionary
    Set json_coll = record.records
    Set specs_dict = CreateObject("Scripting.Dictionary")
    
    If json_coll Is Nothing Then
        Set GetSpecifications = Nothing
        Exit Function
    Else
        For Each json_dict In json_coll
            Set spec = Factory.CreateSpecFromDict(json_dict)
            specs_dict.Add json_dict.Item("Revision"), spec
            rev = json_dict.Item("Revision")
        Next json_dict
        specs_dict.Item(rev).IsLatest = True
        Set GetSpecifications = specs_dict
    End If
    Exit Function
NullSpecException:
    Set GetSpecifications = Nothing
End Function

Sub PrintSpecification(frm As MSForms.UserForm)
    Logger.Log "Printing Specification . . . "
    Set App.console = Factory.CreateConsoleBox(frm)
    App.console.PrintObject App.current_spec
End Sub

Sub PrintTemplate(frm As MSForms.UserForm)
    Logger.Log "Printing Template . . . "
    Set App.console = Factory.CreateConsoleBox(frm)
    App.console.PrintObject App.current_template
End Sub

Function SaveSpecification(spec As Specification) As Long
    If App.current_user.ProductLine = App.current_template.ProductLine Or App.current_user.ProductLine = "Admin" Then
        SaveSpecification = IIf(DataAccess.PushSpec(spec) = DB_PUSH_SUCCESS, DB_PUSH_SUCCESS, DB_PUSH_FAILURE)
    Else
        SaveSpecification = DB_PUSH_DENIED
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

Function DeleteSpecification(spec As Specification) As Long
    If App.current_user.PrivledgeLevel = USER_ADMIN Then
        DeleteSpecification = IIf(DataAccess.DeleteSpec(spec) = DB_DELETE_SUCCESS, DB_DELETE_SUCCESS, DB_DELETE_FAILURE)
    Else
        DeleteSpecification = DB_DELETE_DENIED
    End If
End Function

Private Function MaterialInputValidation(material_id As String) As String
' Ensures that the material id input by the user is parseable.
' TODO: This function is awful need to refactor unsure how due to the
'       ridiculous lack of uniqueness in the database.
'       "The style 101 problem"
    If (material_id <> "101") And (Mid(material_id, 5, 3) <> "101") Then
        MaterialInputValidation = material_id
        Exit Function
    End If
    If Len(material_id) >= 5 Then
        MaterialInputValidation = Mid(material_id, 5, 3) & Mid(material_id, 2, 2)
    Else
        Dim question As Integer
        question = MsgBox("Click Yes for Style 101 Kevlar or No for Hyosung.", vbYesNo + vbQuestion, "Style 101 has two version")
        If question = vbYes Then
            MaterialInputValidation = "101" & "KE"
        Else
            MaterialInputValidation = "101" & "HY"
        End If
    End If
End Function

Function SelectLatestSpec() As Specification
    Dim key As Variant
    For Each key In App.specs
        If App.specs.Item(key).IsLatest = True Then
            Set SelectLatestSpec = App.specs.Item(key)
        End If
    Next key
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

Sub WorksheetToDatabase()
    Dim ws As Worksheet
    Dim i, j As Integer
    Dim last_row As Integer
    Dim number_props As Integer
    Dim property As String
    
    With App
        Set ws = ActiveWorkbook.Sheets(.current_template.SpecType & " Upload")
        last_row = ws.Range("A1").End(xlDown).Row
        number_props = .current_template.Properties.count
        For i = 2 To last_row
            InitializeNewSpecification
            Logger.Log CStr(number_props)
            For j = 1 To number_props
            property = Utils.ConvertToCamelCase(CStr(ws.Cells(1, j).value))
            Logger.Log "Column " & j & ": " & property & ", Row " & i & ": " & CStr(ws.Cells(i, j).value)
            .current_spec.Properties.Item(property) = ws.Cells(i, j).value
            If property = "MaterialNumber" Then
                .current_spec.MaterialId = ws.Cells(i, j).value
            End If
            Next j
            Logger.Log "DataAccess returned : " & SaveSpecification(.current_spec)
            Set .current_spec = Nothing
        Next i
    End With
End Sub

Public Sub DumpAllSpecsToWorksheet(spec_type As String)
    Dim ws As Worksheet
    Dim dicts As Collection
    Dim dict As Object
    Dim props As Variant
    RestartSpecManager
    Logger.LogEnabled False
    Application.ScreenUpdating = False
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Utils.CreateNewSheet(spec_type)
    Set dicts = DataAccess.SelectAllSpecifications(spec_type)
    i = 2
    For Each dict In dicts
        Set App.current_spec = Factory.CreateSpecFromDict(dict)
        props = App.current_spec.ToArray
        If i = 2 Then ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).value = App.current_spec.Header
        ws.Range(Cells(i, 1), Cells(i, ArrayLength(props))).value = props
        i = i + 1
    Next dict
    ws.Range(Cells(1, 1), Cells(1, ArrayLength(props))).columns.AutoFit
    Application.ScreenUpdating = True
End Sub
