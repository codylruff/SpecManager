VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Type SQL
    Select As VBA.Collection
    From As String
    Insert As Dictionary
    Where As Dictionary
End Type

Public Function ExecuteSQL() As Long
    Dim SQLstmt As String
    If SelectAll Then
        SQLstmt = “Select * From” & SQL.From
From
ExecuteSQL = 0
End Function

