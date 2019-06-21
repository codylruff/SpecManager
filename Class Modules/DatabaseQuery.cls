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

private Recordset_ As DatabaseRecord

public property Set Recordset(obj As DatabaseRecord)
    Set Recordset_ = obj
End Property

public property Get Recordset() As DatabaseRecord
    Set Recordset = Recordset_
End Property

Public Type SQL
    Select As VBA.Collection
    From As String
    Insert As Dictionary
    Where As Dictionary
End Type

Private Function ExecuteSQL() As Long
    Dim SQLstmt As String
    If SelectAll Then
        SQLstmt = “Select * From” & SQL.From
    ExecuteSQL = 0
End Function

Private Function ExecuteSQLSelect(db As IDatabase, path As String, SQLstmt As String) As Long
' Returns an table like array
    Set recordset = New DatabaseRecord
    On Error GoTo NullRecordException
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    source.openDb path
    source.selectQry SQLstmt
    recordset.Data = source.Data
    recordset.Header = source.Header
    recordset.Rows = source.NumRows
    recordset.columns = source.NumColumns
    ExecuteSQLSelect = 0
    Exit Function
NullRecordException:
    Logger.Log "SQL Select Error : NullRecordException!"
    ExecuteSQLSelect = -1
End Function

