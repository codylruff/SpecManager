VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------
'CLASS VARIABLES
'----------------
Public Properties         As Dictionary
Private SpecType_         As String
Private Revision_         As String
'----------------
' AUTO-PROPERTIES
'----------------
Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
'--METHODS--'
Public Sub AddProperty(property As String)
' Adds a property to the list of properties
    Me.Properties.Add property, vbNullString
End Sub

Public Function PropertiesJson() As String
    PropertiesJson = JsonConverter.ConvertToJson(Me.Properties)
End Function

Public Sub JsonToObject(properties_json As String, spec_type As String, Revision_ As String)
    Set Properties = JsonConverter.ParseJson(properties_json)
    SpecType = spec_type
    Revision = Revision_
    Properties.Item("Revision") = Revision_
    Properties.Item("Spec_Type") = spec_type
End Sub

' Constructor
Private Sub Class_Initialize()
    Set Properties = New Dictionary
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Properties = Nothing
End Sub
