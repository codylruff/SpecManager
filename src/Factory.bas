Attribute VB_Name = "Factory"
Function CreateBallisticPackage(package_length_inches As Double, fabric_width_inches As Double, conditioned_weight_gsm As Double, target_psf As Double) As BallisticPackage
    Dim package As BallisticPackage
    Set package = New BallisticPackage
    With package
        .PackageLengthInches = package_length_inches
        .FabricWidthInches = fabric_width_inches
        .ConditionedWeight = conditioned_weight_gsm
        .TargetPsf = target_psf
    End With
    Set CreateBallisticPackage = package
End Function

Function CreateDictionary() As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function CreateTable(ws As Worksheet, table_name As String) As Table
    Dim tbl As Table
    Set tbl = New Table
    Set tbl.ListObject = ws.ListObjects(table_name)
    Set CreateTable = tbl
End Function

Function CreateDocument() As Document
    Set CreateDocument = New Document
End Function

Function CreateDocumentFromJsonFile(path As String) As Document
' Generate a specification object from a json file.
    Dim doc As Document
    Dim FSO As Object
    On Error GoTo ErrorHandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Debug.Print FSO.GetBaseName(path)
    Set spec = CreateDocument
    doc.JsonToObject JsonVBA.ReadJsonFileToString(path)
    doc.MaterialId = FSO.GetBaseName(path)
    doc.SpecType = CStr(Prompt.UserInput(SingleLineText, "Template Type Selection", "Assign a template for this specification:"))
    doc.Revision = "1.0"
    doc.Template = App.templates(doc.SpecType)
    doc.MachineId = Prompt.GetMachineId
    Set CreateDocumentFromJsonFile = spec
    Exit Function
ErrorHandler:
    Logger.Log "File could not be read.", ErrorLog
    Exit Function
End Function

Function CreateTemplateFromJsonFile(path As String, Optional product_line As String = "Protection") As Template
' Generate a template object from a json file.
    Dim Template As Template
    Dim FSO As Object
    On Error GoTo ErrorHandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Debug.Print FSO.GetBaseName(path)
    Set Template = New Template
    Template.SpecType = FSO.GetBaseName(path)
    Template.JsonToObject JsonVBA.ReadJsonFileToString(path), Template.SpecType, "1.0", product_line
    Set CreateTemplateFromJsonFile = Template
    Exit Function
ErrorHandler:
    Logger.Log "File could not be read.", ErrorLog
    Exit Function
End Function

Function CopyDocument(doc As Document) As Document
    Dim spec_copy As Document
    Set spec_copy = New Document
    On Error Resume Next
    With spec
        Set spec_copy.Template = .Template
        spec_copy.JsonToObject .PropertiesJson
        spec_copy.MaterialId = .MaterialId
        spec_copy.SpecType = .SpecType
        spec_copy.Revision = .Revision
        spec_copy.MaterialDescription = .MaterialDescription
        spec_copy.MachineId = .MachineId
    End With
    Set CopyDocument = spec_copy
End Function

Function CopyTemplate(Temp As Template) As Template
    Dim temp_copy As Template
    Set temp_copy = New Template
    On Error Resume Next
    With Temp
        temp_copy.JsonToObject .PropertiesJson, .SpecType, .Revision, .ProductLine
    End With
    Set CopyTemplate = temp_copy
End Function

Function CreateNewTemplate(Optional template_name As String = nullstr) As Template
    Dim Template As Template
    Set Template = New Template
    Template.SpecType = template_name
    Template.Revision = 1
    Set CreateNewTemplate = Template
End Function

Function CreateDocumentFromRecord(df As DataFrame) As Document
    Dim doc As Document
    Set doc = New Document
    On Error Resume Next
    With df.records(1)
        doc.MaterialId = .item("Material_Id")
        doc.MaterialDescription = .item("Description")
        doc.ProcessId = .item("Process_Id")
        doc.SpecType = .item("Spec_Type")
        doc.Revision = CStr(.item("Revision"))
    Set doc.Template = Factory.CopyTemplate(App.templates(doc.SpecType))
        doc.JsonToObject .item("Properties_Json")
        doc.MachineId = .item("Machine_Id")
    End With
    Set CreateDocumentFromRecord = doc
End Function

Function CreateTemplateFromRecord(df As DataFrame) As Template
    Dim Template As Template
    Set Template = New Template
    ' obsoleted
    With df.records(1)
        Template.JsonToObject .item("Properties_Json"), .item("Spec_Type"), .item("Revision"), .item("Product_Line")
    End With
    Set CreateTemplateFromRecord = Template
End Function

Function CreateSpecFromDict(dict As Object) As Document
    Dim doc As Document
    Set doc = New Document
    'On Error Resume Next
    With dict
        doc.MaterialId = .item("Material_Id")
        doc.MaterialDescription = .item("Description")
        doc.ProcessId = .item("Process_Id")
        doc.SpecType = .item("Spec_Type")
        doc.Revision = CStr(.item("Revision"))
    Set doc.Template = Factory.CopyTemplate(App.templates(doc.SpecType))
        doc.JsonToObject .item("Properties_Json")
        doc.MachineId = .item("Machine_Id")
    End With
    Set CreateSpecFromDict = doc
End Function

Function CreateTemplateFromDict(dict As Object) As Template
    Dim Temp As Template
    Set Temp = New Template
    With dict
        Temp.JsonToObject .item("Properties_Json"), .item("Spec_Type"), .item("Revision"), .item("Product_Line")
    End With
    Set CreateTemplateFromDict = Temp
End Function

Function CreateDocumentPrinter() As DocumentPrinter
    Dim obj As DocumentPrinter
    Set obj = New DocumentPrinter
    Set CreateDocumentPrinter = obj
End Function

Function CreateDatabaseRecord() As DataFrame
' Creates a database record object
    Dim df: Set df = New DataFrame
    Set CreateDatabaseRecord = df
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

Function CreateConsole(ws As Worksheet) As Console
' Creates a Console object to handle a IForm_Console output.
    Dim obj As Console
    Set obj = New Console
    Set obj.Sheet = ws
    Set CreateConsole = obj
End Function

Function CreateFormPlanning() As FormPlanning

    Dim frm As FormPlanning
    Set frm = New FormPlanning
    Set CreateFormPlanning = frm
End Function

Function CreateFormCreate() As FormCreate
' Creates a FormCreate Object for the GUI to handle.
    Dim frm As FormCreate
    Set frm = New FormCreate
    Set CreateFormCreate = frm
End Function

Function CreateFormPortal() As FormPortal
' Creates a FormPortal Object for the GUI to handle.
    Dim frm As FormPortal
    Set frm = New FormPortal
    Set CreateFormPortal = frm
End Function

Function CreateFormNavigation() As FormNavigation
' Creates a FormNavigation Object for the GUI to handle.
    Dim frm As FormNavigation
    Set frm = New FormNavigation
    Set CreateFormNavigation = frm
End Function

Function CreateFormEdit() As FormEdit
' Creates a FormEdit Object for the GUI to handle.
    Dim frm As FormEdit
    Set frm = New FormEdit
    Set CreateFormEdit = frm
End Function

Function CreateFormView() As FormView
' Creates a FormView Object for the GUI to handle.
    Dim frm As FormView
    Set frm = New FormView
    Set CreateFormView = frm
End Function

'Function CreateFiltrationPlanningForm() As FiltrationPlanningForm
'
'    Dim frm As FiltrationPlanningForm
'    Set frm = New FiltrationPlanningForm
'    With frm
'        Set .Sheet = Nothing
'    End With
'    Set CreateFiltrationPlanningForm = frm
'End Function
'
'Function CreateAdminForm() As AdminForm
'
'    Dim frm As AdminForm
'    Set frm = New AdminForm
'    With frm
'        Set .Sheet = shtCreate
'    End With
'    Set CreateAdminForm = frm
'End Function
'
'Function DocumentConfigForm() As DocumentConfigForm
'
'    Dim frm As DocumentConfigForm
'    Set frm = New DocumentConfigForm
'    With frm
'        Set .Sheet = shtSpecConfig
'    End With
'    Set DocumentConfigForm = frm
'End Function
