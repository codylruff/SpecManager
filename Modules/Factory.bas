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

Function CreateSpecificationFromJsonFile(path As String) As Specification
' Generate a specification object from a json file.
    Dim spec As Specification
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Debug.Print FSO.GetBaseName(path)
    Set spec = CreateSpecification
    spec.JsonToObject JsonVBA.ReadJsonFileToString(path), vbNullString, FSO.GetBaseName(path), "Weaving RBA", "1.0"
    spec.Template = App.templates
    Set CreateSpecificationFromJsonFile = spec
End Function

Function CopySpecification(spec As Specification) As Specification
    Dim spec_copy As Specification
    Set spec_copy = New Specification
    With spec
        spec_copy.JsonToObject .PropertiesJson, .TolerancesJson, .MaterialId, .SpecType, .Revision
    End With
    Set CopySpecification = spec_copy
End Function

Function CopyTemplate(temp As SpecificationTemplate) As SpecificationTemplate
    Dim temp_copy As SpecificationTemplate
    Set temp_copy = New SpecificationTemplate
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

Function CreateTemplateFromRecord(record As DatabaseRecord) As SpecificationTemplate
    Dim Template As SpecificationTemplate
    Set Template = New SpecificationTemplate
    ' obsoleted
    With record.Fields
        Template.JsonToObject .Item("Properties_Json"), .Item("Spec_Type"), .Item("Revision"), .Item("Product_Line")
    End With
    Set CreateTemplateFromRecord = Template
End Function

Function CreateSpecFromDict(dict As Object) As Specification
    Dim spec As Specification
    Set spec = New Specification
    With dict
        spec.JsonToObject .Item("Properties_Json"), .Item("Tolerances_Json"), .Item("Material_Id"), .Item("Spec_Type"), .Item("Revision")
    End With
    Set spec.Template = Factory.CopyTemplate(App.templates(spec.SpecType))
    Set CreateSpecFromDict = spec
End Function

Function CreateTemplateFromDict(dict As Object) As SpecificationTemplate
    Dim temp As SpecificationTemplate
    Set temp = New SpecificationTemplate
    With dict
        temp.JsonToObject .Item("Properties_Json"), .Item("Spec_Type"), .Item("Revision"), .Item("Product_Line")
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

Function CreateSQLiteDatabase() As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDatabase = sqlite
End Function
