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

Implements IQueryable

' Events
Public Event PropertyAdded()
Public Event PropertyChanged(property As Variant, old_val As Variant, new_val As Variant)

' Properties dictionary
Public Properties   As Object
' FIELDS and PROPERTIES---
Private SpecType_   As String
Private Revision_   As String
Private MaterialId_ As String
Private IsStandard_ As Boolean
Private IsLatest_   As Boolean
Private Template_   As SpecificationTemplate
Private MaterialDescription_ As String
Private ProcessId_  As String
' Template
Public Property Set Template(obj As SpecificationTemplate)
    Set Template_ = obj
End Property
Public Property Get Template() As SpecificationTemplate
    Set Template = Template_
End Property

Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
Public Property Get MaterialId() As String: MaterialId = MaterialId_: End Property
Public Property Let MaterialId(value As String): MaterialId_ = UCase(value): End Property
Public Property Get MaterialDescription() As String: MaterialDescription = MaterialDescription_: End Property
Public Property Let MaterialDescription(val As String): MaterialDescription_ = val: End Property
Public Property Let ProcessId(value As String): ProcessId_ = value: End Property
Public Property Get ProcessId() As String: ProcessId = ProcessId_: End Property

' METHODS---
Public Function GetProperty(property As String) As Variant
    If Me.Properties.Exists(property) Then
        GetProperty = Me.Properties(property)
    Else
        GetProperty = nullstr
    End If
End Function

Public Sub AddProperty(property As String)
' Adds a property to the list of properties
    Me.Properties.Add property, nullstr
End Sub

Public Sub ChangeProperty(property As Variant, new_val As Variant)
' Use this procedure to change a property raising the property changed event.
    Dim old_val As Variant
    old_val = Me.Properties(property)
    Me.Properties(property) = new_val
    'RaiseEvent PropertyChanged(property, old_val, new_val)
End Sub

Public Sub JsonToObject(properties_json As String)
' Map json to the spec
    Dim props_temp As Object
    Dim prop As Variant
    Set props_temp = JsonVBA.ParseJson(properties_json)
    If Not IsNothing(Me.Template) Then
        Set Me.Properties = Template.Properties
        On Error Resume Next
        For Each prop In Me.Properties
            Me.Properties(prop) = props_temp(prop)
        Next prop
        On Error GoTo 0
    Else
        Set Me.Properties = props_temp
    End If

End Sub

Public Function PropertiesJson() As String
' Store spec Properties dictionary as json
    PropertiesJson = JsonVBA.ConvertToJson(Me.Properties)
End Function

Public Function header() As Variant
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
    header = arr
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
        arr(i) = Properties.item(prop)
        i = i + 1
    Next prop
    ToArray = arr
End Function

Public Function ToString() As String
    Dim print_string As String
    Dim prop As Variant
    print_string = nullstr
    For Each prop In header
        print_string = Utils.printf(print_string & " {0} {1}", " |", CStr(prop))
    Next prop
    print_string = print_string & vbNewLine
    For Each prop In ToArray
        print_string = Utils.printf(print_string & " {0} {1}", " |", CStr(prop))
    Next prop
    ToString = print_string
End Function

' Constructor
Private Sub Class_Initialize()
    Set Template = New SpecificationTemplate
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Template = Nothing
    Set Properties = Nothing
End Sub

' IQueryable Methods
Function IQueryable_GetValueLabels() As String
    IQueryable_GetValueLabels = "Material_Id, Time_Stamp, Properties_Json, Revision, Spec_Type"
End Function

Function IQueryable_GetValues() As String
    IQueryable_GetValues = "'" & Me.MaterialId & "', " & _
                           "'" & CStr(Now()) & "', " & _
                           "'" & Me.PropertiesJson & "', " & _
                           "'" & Me.Revision & "', " & _
                           "'" & Me.SpecType & "'"
End Function
