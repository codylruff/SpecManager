VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQueryable

Private m_User As String
Private m_Time_Stamp As String
Private m_description As String
Private m_work_order As String
Private m_MaterialId As String
Private m_SpecType As String
Private m_Revision As String

Public Property Let Revision(value As String): m_Revision = value: End Property
Public Property Get Revision() As String: Revision = m_Revision: End Property
Public Property Let SpecType(value As String): m_SpecType = value: End Property
Public Property Get SpecType() As String: SpecType = m_SpecType: End Property
Public Property Let MaterialId(value As String): m_MaterialId = value: End Property
Public Property Get MaterialId() As String: MaterialId = m_MaterialId: End Property
Public Property Let work_order(value As String): m_work_order = value: End Property
Public Property Get work_order() As String: work_order = m_work_order: End Property
Public Property Let description(value As String): m_description = value: End Property
Public Property Get description() As String: description = m_description: End Property
Public Property Let Time_Stamp(value As String): m_Time_Stamp = value: End Property
Public Property Get Time_Stamp() As String: Time_Stamp = m_Time_Stamp: End Property
Public Property Let User(value As String): m_User = value: End Property
Public Property Get User() As String: User = m_User: End Property

' IQueryable Methods
Function IQueryable_GetValueLabels() As String
    IQueryable_GetValueLabels = "User, Time_Stamp, Action_Description, Work_Order, Material_Id, Spec_Type, Spec_Revision"
End Function

Function IQueryable_GetValues() As String
    IQueryable_GetValues = "'" & m_User & "', " & _
               "'" & m_Time_Stamp & "', " & _
               "'" & m_description & "', " & _
               "'" & m_work_order & "', " & _
               "'" & m_MaterialId & "', " & _
               "'" & m_SpecType & "', " & _
               "'" & m_Revision & "'"
End Function
