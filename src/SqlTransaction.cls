VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DbConnection As SQLiteDatabase 'IDatabase
Private ConnectionString_ As String
' ConnectionString
Private Property Let ConnectionString(value As String)
    ConnectionString_ = value
End Property
Public Property Get ConnectionString() As String
    ConnectionString = ConnectionString_
End Property

Public Sub Connect(path As String)
' Connect to sqlite database
    If ConnectionString_ = nullstr Then
        DbConnection.openDb path ' IDatabase.openDb
        Logger.Log "Connection established.", SqlLog
        ConnectionString = path
    Else
        Logger.Log "Connection already established.", SqlLog
    End If
End Sub

Public Function Begin() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    Logger.Log "Starting SQL Transaction", SqlLog
    ExecuteSQL "BEGIN TRANSACTION"
    Logger.Log "SQL Transaction Initialized", SqlLog
    Begin = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    Logger.Log "SQL BEGIN Error : DbTransactionFailException", SqlLog
    Begin = DB_TRANSACTION_ERR
End Function

Public Function Rollback() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL "ROLLBACK TRANSACTION"
    Logger.Log "SQL Transaction Rolled Back", SqlLog
    Rollback = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    Logger.Log "SQL ROLLBACK Error : DbTransactionFailException", SqlLog
    Rollback = DB_TRANSACTION_ERR
End Function

Public Function Commit() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL "COMMIT TRANSACTION"
    Logger.Log "SQL Transaction Commited", SqlLog
    Commit = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    Logger.Log "SQL COMMIT Error : DbTransactionFailException", SqlLog
    Commit = DB_TRANSACTION_ERR
End Function

Function ExecuteSQLSelect(SQLstmt As String) As DataFrame
' Returns an table like array
    Dim df As DataFrame
    Set df = New DataFrame
    On Error GoTo NullRecordException
    Logger.Log "-----------------------------------", SqlLog
    Logger.Log SQLstmt
    With Me.DbConnection
        .selectQry SQLstmt ' IDatabase.selectQry -> Null or Boolean
        df.Data = .Data ' IDatabase.Data -> Variant
        df.header = .header ' IDatabase.header -> Variant
        df.Rows = .NumRows ' Idatabase.NumRows -> Long
        df.Columns = .NumColumns ' IDatabase.NumColumns -> Long
    End With
    Set ExecuteSQLSelect = df
    Exit Function
NullRecordException:
    Logger.Log "SQL Select Error : NullRecordException!", SqlLog
    Set ExecuteSQLSelect = New DataFrame
End Function

Sub ExecuteSQL(SQLstmt As String)
' Performs update or insert querys returns error on select.
    Logger.Log "-----------------------------------", SqlLog
    Logger.Log SQLstmt, SqlLog
    If Left(SQLstmt, 6) = "SELECT" Then
        Logger.Log "Use ExecuteSQLSelect() for SELECT query", SqlLog
        Exit Sub
    Else
        'db.openDb (path)
        Me.DbConnection.execute SQLstmt ' IDatabase.execute
    End If
End Sub


' Constructor
Private Sub Class_Initialize()
    Set DbConnection = New SQLiteDatabase
End Sub

' Deconstructor
Private Sub Class_Terminate()
    Set DbConnection = Nothing
    Logger.Log "Connection Terminated.", SqlLog
End Sub
