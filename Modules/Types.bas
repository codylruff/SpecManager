Attribute VB_Name = "Types"
Option Explicit

Public Type UserAction
    User As String
    Time_Stamp As String
    Description As String
    work_order As String
    spec As Specification
End Type

Public Type Log
    Buffer As VBA.Collection
    log_type As LogType
    Id As String
End Type

Public Type Rect
    Left        As Long  ' x1
    Top         As Long  ' y1
    Right       As Long  ' x2
    Bottom      As Long  ' y2
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type FourBytes
    A As Byte
    B As Byte
    C As Byte
    D As Byte
End Type

Public Type OneLong
    L As Long
End Type
