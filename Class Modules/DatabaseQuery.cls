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

Private Recordset_ As DatabaseRecord
Private source As IDatabase
Private db_path As String

Public Property Set Recordset(obj As DatabaseRecord)
    Set Recordset_ = obj
End Property

Public Property Get Recordset() As DatabaseRecord
    Set Recordset = Recordset_
End Property

Public Type sql
    Select As VBA.Collection
    SelectAll As Boolean
    From As String
    Insert As Dictionary
    Where As Dictionary
    Query As String
End Type

Public Function ConnectToDataSource(path As String) As Long
    Set source = Factory.CreateSQLiteDatabase()
    db_path = path
    ConnectToDataSource = 0
End Function

Public Function execute() As Long
    If sql.Select.Count = 0 Then
        execute = ExecuteSQL(SQLstmt)
    Else
        execute = ExecuteSQLSelect(SQLstmt)
    End If
End Function

Private Function BuildSqlQuery() As Boolean
    SqlQuery = vbNullString
    
    If sql.SelectAll Then
        Append “SELECT * From” & sql.From”
        BuildSqlQuery = True
        Exit Function
    Else
        Append “SELECT “, IIf(sql.SelectAll, “* “, sql.Select), “ From “, sql.From
    End If

    
End Function

Private Function ExecuteSQL() As Long
        
    ExecuteSQL = 0
End Function

Private Function ExecuteSQLSelect(SQLstmt As String) As Long
' Returns an table like array
    Set Recordset = New DatabaseRecord
    On Error GoTo NullRecordException
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    source.openDb db_path
    source.selectQry SQLstmt
    With Recordset
        .Data = source.Data
        .header = source.header
        .Rows = source.NumRows
        .Columns = source.NumColumns
    End With
    ExecuteSQLSelect = 0
    Exit Function
NullRecordException:
    Logger.Log "SQL Select Error : NullRecordException!"
    ExecuteSQLSelect = -1
End Function

