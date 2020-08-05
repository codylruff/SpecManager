VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
Implements IForm
' DEFAULT PROPERTY VALUES
Private Const FORM_NAME As String = "FormCreate"
Private Const HELP_TEXT As String = "TODO" ' Thinking of loading this from a text file??
Private m_IForm_Console As Console
Private m_IForm_Fields() As Variant
'--------------
' I_PROPERTIES
'--------------
Public Property Set IForm_Console(value As Console): Set m_IForm_Console = value: End Property
Public Property Get IForm_Console() As Console: Set IForm_Console = m_IForm_Console: End Property

Public Property Get IForm_Sheet() As Worksheet
    Set IForm_Sheet = shtCreate
End Property

Public Property Get IForm_Name() As String
    IForm_Name = FORM_NAME
End Property

Public Property Get IForm_HelpText() As String
    IForm_HelpText = HELP_TEXT
End Property

Public Property Get IForm_Fields() As Variant
    IForm_Fields = m_IForm_Fields
End Property

'----------
' I_METHODS
'----------
Public Function IForm_Clear()
    ' Console must be cleared through object
    IForm_Console.Clear
    ' Fields can be cleared manually.
    For i = 0 To UBound(IForm_Fields) - 1
        IForm_Sheet.Range(IForm_Fields(i)).value = nullstr
    Next i
End Function

Public Function IForm_CmdByName(cmd As String) As Boolean
    Dim ret_val As Boolean
    Select Case cmd
        Case "LoadExcelDocument"
            Cmd_LoadExcelDocument
            ret_val = True
        Case "LoadJsonDocument"
            Cmd_LoadJsonDocument
            ret_val = True
        Case "CreateTemplateFromFile"
            Cmd_CreateTemplateFromFile
            ret_val = True
        Case "CreateBallisticsDocument"
            Cmd_CreateBallisticsDocument
            ret_val = True
        Case Default
            err.Raise INTERNAL_ERR, "Command Not Found"
            ret_val = False
    End Select
    IForm_CmdByName = ret_val
End Function

Public Function IForm_ReformatSheet() As Boolean
    ' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
    IForm_ReformatSheet = False
End Function

Public Function IForm_FieldValue(field As String) As Variant
' Reads the value of a field in the form
    Dim i As Long
    Dim val As Variant
    On Error GoTo Catch
    IForm_FieldValue = IForm_Sheet.Range(field).value
    GoTo Finally
Catch:
    err.Raise INTERNAL_ERR, "Field not found"
    IForm_FieldValue = nullstr
Finally:
    ' Pass
End Function

Public Function IForm_ChangeFieldValue(field As String, val As Variant)
' Changes the value of a field in the form
    IForm_Sheet.Range(field).value = val
    'err.Raise INTERNAL_ERR, "Field not found"
End Function
'----------
' COMMANDS
'----------
Private Sub Cmd_LoadExcelDocument()
    DocumentParser.LoadNewDocument "excel"
End Sub

Private Sub Cmd_LoadJsonDocument()
    DocumentParser.LoadNewDocument "json"
End Sub

Private Sub Cmd_CreateTemplateFromFile()
    DocumentParser.CreateTemplateFromFile
End Sub

Private Sub Cmd_CreateBallisticsDocument()
' Creates a ballistic testing document
    Dim material_id As String
    Dim package_length_inches As Double
    Dim fabric_width_inches As Double
    Dim conditioned_weight_gsm As Double
    Dim target_psf As Double
    Dim ret_val As Long
    Dim machine_id As String

    App.Start
    material_id = shtCreate.Range("material_id").value ' this is the material id (SAP Code)
    package_length_inches = shtCreate.Range("package_length_inches")
    fabric_width_inches = shtCreate.Range("fabric_width_inches")
    conditioned_weight_gsm = shtCreate.Range("conditioned_weight_gsm")
    target_psf = shtCreate.Range("target_psf")
    machine_id = CStr(shtCreate.Range("machine_id").value)   ' This is the machine id (ie. loom number, warper, etc...)

    ret_val = SpecManager.BuildBallisticTestSpec(material_id, package_length_inches, fabric_width_inches, conditioned_weight_gsm, target_psf, machine_id, False)
    
    ' Parse return value.
    If ret_val = DB_PUSH_SUCCESS Then
        Prompt.Success "New Document Saved."
    ElseIf ret_val = MATERIAL_EXISTS_ERR Then
        Prompt.Error "Material Already Exists."
    Else
        Prompt.Error "Error Saving Document."
    End If
    
    App.Shutdown
End Sub

Private Sub Cmd_SaveChanges()
' Writes the pending design to the database.
' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
End Sub

Private Function InitializeFieldsArray() As Variant
' Creates an array of ranges representing the fields in this form.
    Dim i, num_fields As Long
    Dim arr() As Variant
    Dim field As String
    Fields = Array("design_class", "design_type", "product_line", "template_id", "property_id", "machine_id")
    num_fields = UBound(Fields) - LBound(Fields) + 1
    ReDim arr(num_fields)
    For i = 0 To num_fields - 1
        arr(i) = IForm_Sheet.Range(field)
    Next i

    InitializeFieldsArray = arr
End Function
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    Set IForm_Console = Nothing
    ' Initialize the fields array to save on Array() calls.
    ' We do not want to include 'Console' because it should only be accessed through the Console object.
    'm_IForm_Fields = InitializeFieldsArray
End Sub
' Deconstructor
Private Sub Class_Terminate()
    'REVIEW Is there anything that needs to be destroyed here?
End Sub
