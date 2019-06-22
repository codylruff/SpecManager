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
private source As IDatabase
private db_path As String

public property Set Recordset(obj As DatabaseRecord)
    Set Recordset_ = obj
End Property

public property Get Recordset() As DatabaseRecord
    Set Recordset = Recordset_
End Property

Public Type SQL
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

Public Function Execute() As Long
    If sql.Select.Count = 0 Then
        Execute = ExecuteSQL(SQLstmt)
    Else
        Execute = ExecuteSQLSelect(SQLstmt)
    End If
End Function

private Function BuildSqlQuery() As Boolean
    SqlQuery = vbNullString
    
    If sql.SelectAll Then
        Append “Select * From” & sql.From”
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
    Set recordset = New DatabaseRecord
    On Error GoTo NullRecordException
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    source.openDb db_path
    source.selectQry SQLstmt
    With recordset
        .Data = source.Data
        .Header = source.Header
        .Rows = source.NumRows
        .columns = source.NumColumns
    End With
    ExecuteSQLSelect = 0
    Exit Function
NullRecordException:
    Logger.Log "SQL Select Error : NullRecordException!"
    ExecuteSQLSelect = -1
End Function

