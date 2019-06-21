VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Specification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event PropertyAdded
' Properties dictionary
Public Properties   As Object
Public Tolerances   As Object
' FIELDS and PROPERTIES---
Private SpecType_   As String
Private Revision_   As String
Private MaterialId_ As String
Private IsStandard_ As Boolean
Private IsLatest_   As Boolean
Private Template_   As SpecificationTemplate

' Template
Public Property Set Template(obj As SpecificationTemplate)
    Set Template_ = obj
End Property
Public Property Get Template() As SpecificationTemplate
    Set Template = Template_
End Property
Public Property Let IsStandard(Value As Boolean): IsStandard_ = Value: End Property
Public Property Get IsStandard() As Boolean: IsStandard = IIf(IsStandard_ = Null, False, IsStandard_): End Property
Public Property Let IsLatest(Value As Boolean): IsLatest_ = Value: End Property
Public Property Get IsLatest() As Boolean: IsLatest = IIf(IsLatest_ = Null, False, IsLatest_): End Property
Public Property Let Revision(Value As String): Revision_ = Value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(Value As String): SpecType_ = Value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
Public Property Get MaterialId() As String: MaterialId = MaterialId_: End Property
Public Property Let MaterialId(Value As String): MaterialId_ = UCase(Value): End Property


' METHODS---
Public Sub AddProperty(property As String, tolerance As Variant)
' Adds a property to the list of properties
    Me.Properties.Add property, vbNullString
    Me.Tolerances.Add property, tolerance
    RaiseEvent PropertyAdded
End Sub

Public Sub JsonToObject(properties_json As String, tolerances_json As String, material_id As String, spec_type As String, Revision_ As String)
' Map json to the spec
    Set Properties = JsonVBA.ParseJson(properties_json)
    Set Tolerances = JsonVBA.ParseJson(tolerances_json)
    MaterialId = material_id
    SpecType = spec_type
    Revision = Revision_
    ' Properties.Add key:="Material_Id", Item:=Me.MaterialId
    ' Properties.Add key:="Spec_Type", Item:=Me.SpecType
    ' Properties.Add key:="Revision", Item:=Me.Revision
End Sub

Public Function PropertiesJson() As String
' Store spec Properties dictionary as json
    PropertiesJson = JsonVBA.ConvertToJson(Properties)
End Function

Public Function TolerancesJson() As String
    TolerancesJson = JsonVBA.ConvertToJson(Me.Tolerances)
End Function

Public Function Header() As Variant
    Dim arr() As String
    Dim i As Long
    Dim prop As Variant
    Dim number_of_props As Long
    number_of_props = Me.Properties.Count + 2
    ReDim arr(number_of_props)
    arr(0) = "Material Id"
    arr(2) = "Revision"
    i = 2
    For Each prop In Me.Properties
        arr(i) = CStr(prop)
        i = i + 1
    Next prop
    Header = arr
End Function

Public Function ToArray() As Variant
    Dim arr() As String
    Dim i As Long
    Dim prop As Variant
    Dim number_of_props As Long
    number_of_props = Me.Properties.Count + 2
    ReDim arr(number_of_props)
    arr(0) = Me.MaterialId
    arr(2) = Me.Revision
    i = 2
    For Each prop In Me.Properties
        arr(i) = Properties.Item(prop)
        i = i + 1
    Next prop
    ToArray = arr
End Function

' Constructor
Private Sub Class_Initialize()
    Set Properties = New Dictionary
    Set Tolerances = New Dictionary
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Properties = Nothing
    Set Tolerances = Nothing
End Sub
