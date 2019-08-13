Attribute VB_Name = "Factory"

Function CreateDictionary() As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function CreateTable() As Table
    Set CreateTable = New Table
End Function

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateUserAction(user_name As String, action_description As String, spec As Specification, Optional work_order As String) As UserAction
' Creates a user action for logging into the database.
    Dim action As UserAction
    With action
        .User = user_name
        .Time_Stamp = Format(CStr(Now()), "dd-MMM-yyyy HH:nn:ss")
        .Description = action_description
        .work_order = work_order
        Set .spec = spec
    End With
    CreateUserAction = action
End Function

Function CreateSpecificationFromJsonFile(path As String) As Specification
' Generate a specification object from a json file.
    Dim spec As Specification
    Dim FSO As Object
    On Error GoTo ErrorHandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Debug.Print FSO.GetBaseName(path)
    Set spec = CreateSpecification
    spec.JsonToObject JsonVBA.ReadJsonFileToString(path), vbNullString
    spec.MaterialId = FSO.GetBaseName(path)
    spec.SpecType = "Weaving RBA"
    spec.Revision = "1.0"
    spec.Template = App.templates("Weaving RBA")
    Set CreateSpecificationFromJsonFile = spec
    Exit Function
ErrorHandler:
    App.logger.Log "File could not be read.", ErrorLog
    Exit Function
End Function

Function CopySpecification(spec As Specification) As Specification
    Dim spec_copy As Specification
    Set spec_copy = New Specification
    On Error Resume Next
    With spec
        spec_copy.JsonToObject .PropertiesJson, .TolerancesJson
        spec_copy.MaterialId = .MaterialId
        spec_copy.SpecType = .SpecType
        spec_copy.Revision = .Revision
        spec_copy.MaterialDescription = .MaterialDescription
    End With
    Set CopySpecification = spec_copy
End Function

Function CopyTemplate(temp As SpecificationTemplate) As SpecificationTemplate
    Dim temp_copy As SpecificationTemplate
    Set temp_copy = New SpecificationTemplate
    On Error Resume Next
    With temp
        temp_copy.JsonToObject .PropertiesJson, .SpecType, .Revision, .ProductLine
    End With
    Set CopyTemplate = temp_copy
End Function

Function CreateNewTemplate(Optional template_name As String = vbNullString) As SpecificationTemplate
    Dim Template As SpecificationTemplate
    Set Template = New SpecificationTemplate
    Template.SpecType = template_name
    Template.Revision = 1
    Set CreateNewTemplate = Template
End Function

Function CreateSpecificationFromRecord(record As DatabaseRecord) As Specification
    Dim spec_ As Specification
    Set spec_ = New Specification
    On Error Resume Next
    With record.records(1)
        spec_.MaterialId = .item("Material_Id")
        spec_.MaterialDescription = .item("Description")
        spec_.ProcessId = .item("Process_Id")
        spec_.SpecType = .item("Spec_Type")
        spec_.Revision = CStr(.item("Revision"))
        spec_.JsonToObject .item("Properties_Json"), .item("Tolerances_Json")
    End With
    Set CreateSpecificationFromRecord = spec_
End Function

Function CreateTemplateFromRecord(record As DatabaseRecord) As SpecificationTemplate
    Dim Template As SpecificationTemplate
    Set Template = New SpecificationTemplate
    ' obsoleted
    With record.records(1)
        Template.JsonToObject .item("Properties_Json"), .item("Spec_Type"), .item("Revision"), .item("Product_Line")
    End With
    Set CreateTemplateFromRecord = Template
End Function

Function CreateSpecFromDict(dict As Object) As Specification
    Dim spec As Specification
    Set spec = New Specification
    On Error Resume Next
    With dict
        spec.MaterialId = .item("Material_Id")
        spec.MaterialDescription = .item("Description")
        spec_.ProcessId = .item("Process_Id")
        spec.SpecType = .item("Spec_Type")
        spec.Revision = CStr(.item("Revision"))
        spec.JsonToObject .item("Properties_Json"), .item("Tolerances_Json")
    End With
    Set spec.Template = Factory.CopyTemplate(App.templates(spec.SpecType))
    Set CreateSpecFromDict = spec
End Function

Function CreateTemplateFromDict(dict As Object) As SpecificationTemplate
    Dim temp As SpecificationTemplate
    Set temp = New SpecificationTemplate
    With dict
        temp.JsonToObject .item("Properties_Json"), .item("Spec_Type"), .item("Revision"), .item("Product_Line")
    End With
    Set CreateTemplateFromDict = temp
End Function

Function CreateDocumentPrinter(frm As UserForm) As DocumentPrinter
    Dim obj As DocumentPrinter
    Set obj = New DocumentPrinter
    Set obj.FormId = frm
    Set CreateDocumentPrinter = obj
End Function

Function CreateDatabaseRecord() As DatabaseRecord
' Creates a database record object
    Dim record: Set record = New DatabaseRecord
    Set CreateDatabaseRecord = record
End Function

Function CreateSQLiteDatabase(path As String) As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDatabase = sqlite
End Function

Function CreateSqlTransaction(path As String) As SqlTransaction
' Creates a sqlite transaction object
    Dim sql_trans: Set sql_trans = New SqlTransaction
    sql_trans.Connect path
    Set CreateSqlTransaction = sql_trans
End Function
