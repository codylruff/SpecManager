Attribute VB_Name = "Factory"

Function CreateDictionary() As Object
    Set CreateDictionary = New Dictionary
End Function

Public Function CreateTable() As Table
    Set CreateTable = New Table
End Function

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateSpecificationFromJsonFile(path As String) As Specification
' Generate a specification object from a json file.
    Dim spec As Specification
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print fso.GetBaseName(path)
    Set spec = CreateSpecification
    spec.JsonToObject JsonVBA.ReadJsonFileToString(path), vbNullString, fso.GetBaseName(path), "Weaving RBA", "1.0"   
End Function

Function CopySpecification(spec As Specification) As Specification
    Dim spec_copy As Specification
    Set spec_copy = New Specification
    With spec
        spec_copy.JsonToObject .PropertiesJson, .TolerancesJson, .MaterialId, .SpecType, .Revision
    End With
    Set CopySpecification = spec_copy
End Function

Function CreateNewTemplate(Optional template_name As String = vbNullString) As SpecificationTemplate
    Dim template As SpecificationTemplate
    Set template = New SpecificationTemplate
    template.SpecType = template_name
    template.Revision = 1
    Set CreateNewTemplate = template
End Function

Function CreateTemplateFromRecord(record As DatabaseRecord) As SpecificationTemplate
    Dim template As SpecificationTemplate
    Set template = New SpecificationTemplate
    ' obsoleted
    With record.Fields
        template.JsonToObject .Item("Properties_Json"), .Item("Spec_Type"), .Item("Revision"), .Item("Product_Line")
    End With
    Set CreateTemplateFromRecord = template
End Function

Function CreateSpecFromDict(dict As Object) As Specification
    Dim spec As Specification
    Set spec = New Specification
    With dict
        spec.JsonToObject .Item("Properties_Json"), .Item("Tolerances_Json"), .Item("Material_Id"), .Item("Spec_Type"), .Item("Revision")
    End With
    Set CreateSpecFromDict = spec
End Function

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
    Dim obj As ConsoleBox
    Set obj = New ConsoleBox
    Set obj.FormId = frm
    Set CreateConsoleBox = obj
End Function

Function CreateDatabaseRecord() As DatabaseRecord
' Creates a database record object
    Dim record: Set record = New DatabaseRecord
    Set CreateDatabaseRecord = record
End Function

Function CreateSQLiteDatabase() As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDatabase = sqlite
End Function
