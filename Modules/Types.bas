Attribute VB_Name = "Types"
Option Explicit

Public Type Log
    Buffer As VBA.Collection
    Log_Type As LogType
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

