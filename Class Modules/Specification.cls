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
' Properties dictionary
Public Properties   As Dictionary
Public Tolerances   As Dictionary
' FIELDS and PROPERTIES---
Private SpecType_   As String
Private Revision_   As String
Private MaterialId_ As String
Private IsStandard_ As Boolean
Private IsLatest_   As Boolean
Public Property Let IsStandard(value As Boolean): IsStandard_ = value: End Property
Public Property Get IsStandard() As Boolean: IsStandard = IIf(IsStandard_ = Null, False, IsStandard_): End Property
Public Property Let IsLatest(value As Boolean): IsLatest_ = value: End Property
Public Property Get IsLatest() As Boolean: IsLatest = IIf(IsLatest_ = Null, False, IsLatest_): End Property
Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
Public Property Get MaterialId() As String: MaterialId = MaterialId_: End Property
Public Property Let MaterialId(value As String): MaterialId_ = value: End Property

' METHODS---
Public Sub AddProperty(property As String, tolerance As Variant)
' Adds a property to the list of properties
    Me.Properties.Add property, vbNullString
    Me.Tolerances.Add property, tolerance
End Sub

Public Sub JsonToObject(properties_json As String, tolerances_json As String, material_id As String, spec_type As String, Revision_ As String)
' Map json to the spec
    Set Properties = JsonConverter.ParseJson(properties_json)
    Set Tolerances = JsonConverter.ParseJson(tolerances_json)
    MaterialId = material_id
    SpecType = spec_type
    Revision = Revision_
    Properties.Add key:="Material_Id", Item:=Me.MaterialId
    Properties.Add key:="Spec_Type", Item:=Me.SpecType
    Properties.Add key:="Revision", Item:=Me.Revision
End Sub

Public Function PropertiesJson() As String
' Store spec Properties dictionary as json
    PropertiesJson = JsonConverter.ConvertToJson(Properties)
End Function

Public Function TolerancesJson() As String
    TolerancesJson = JsonConverter.ConvertToJson(Me.Tolerances)
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
