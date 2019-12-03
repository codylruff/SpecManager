VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecificationTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQueryable

'----------------
'CLASS VARIABLES
'----------------
Public Properties         As Object
Private SpecType_         As String
Private Revision_         As String
Private ProductLine_      As String
'----------------
' AUTO-PROPERTIES
'----------------
Public Property Let ProductLine(value As String): ProductLine_ = value: End Property
Public Property Get ProductLine() As String: ProductLine = ProductLine_: End Property
Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
'--METHODS--'
Public Sub AddProperty(property As String)
' Adds a property to the list of properties
    Me.Properties.Add property, nullstr
End Sub

Public Sub RemoveProperty(property As String)
    On Error Resume Next
    Me.Properties.Remove property
End Sub

Public Function PropertiesJson() As String
    PropertiesJson = JsonVBA.ConvertToJson(Me.Properties)
End Function

Public Sub JsonToObject(properties_json As String, spec_type As String, Revision_ As String, product_line As String)
    Set Properties = JsonVBA.ParseJson(properties_json)
    SpecType = spec_type
    Revision = Revision_
    ProductLine = product_line
End Sub

' Constructor
Private Sub Class_Initialize()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Properties = Nothing
End Sub

' IQueryable Methods
Function IQueryable_GetValueLabels() As String
    IQueryable_GetValueLabels = "Time_Stamp, Properties_Json, Revision, Spec_Type, Product_Line"
End Function

Function IQueryable_GetValues() As String
    IQueryable_GetValues = "'" & CStr(Now()) & "'," & _
                           "'" & Me.PropertiesJson & "', " & _
                           "'" & Me.Revision & "', " & _
                           "'" & Me.SpecType & "', " & _
                           "'" & Me.ProductLine & "'"
End Function