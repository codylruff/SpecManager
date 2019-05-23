Attribute VB_Name = "Factory"

Function CreateDictionary() As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
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
    record.SetDictionary
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