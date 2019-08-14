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

Public DbConnection As SQLiteDatabase
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
    If ConnectionString_ = vbNullString Then
        DbConnection.openDb path
        App.logger.Log "Connection established."
        ConnectionString = path
    Else
        App.logger.Log "Connection already established."
    End If
End Sub

Public Function Begin() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    App.logger.Log "Starting SQL Transaction", SqlLog
    ExecuteSQL "BEGIN TRANSACTION"
    App.logger.Log "SQL Transaction Initialized", SqlLog
    Begin = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL BEGIN Error : DbTransactionFailException", SqlLog
    Begin = DB_TRANSACTION_FAILURE
End Function

Public Function Rollback() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL "ROLLBACK TRANSACTION"
    App.logger.Log "SQL Transaction Rolled Back", SqlLog
    Rollback = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL ROLLBACK Error : DbTransactionFailException", SqlLog
    Rollback = DB_TRANSACTION_FAILURE
End Function

Public Function Commit() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL "COMMIT TRANSACTION"
    App.logger.Log "SQL Transaction Commited", SqlLog
    Commit = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL COMMIT Error : DbTransactionFailException", SqlLog
    Commit = DB_TRANSACTION_FAILURE
End Function

Function ExecuteSQLSelect(SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record As DatabaseRecord
    Set record = New DatabaseRecord
    On Error GoTo NullRecordException
    App.logger.Log "-----------------------------------", SqlLog
    App.logger.Log SQLstmt, SqlLog
    With Me.DbConnection
        .selectQry SQLstmt
        record.Data = .Data
        record.header = .header
        record.Rows = .NumRows
        record.Columns = .NumColumns
    End With
    Set ExecuteSQLSelect = record
    Exit Function
NullRecordException:
    App.logger.Log "SQL Select Error : NullRecordException!", SqlLog
    Set ExecuteSQLSelect = New DatabaseRecord
End Function

Sub ExecuteSQL(SQLstmt As String)
' Performs update or insert querys returns error on select.
    App.logger.Log "-----------------------------------", SqlLog
    App.logger.Log SQLstmt, SqlLog
    If Left(SQLstmt, 6) = "SELECT" Then
        App.logger.Log "Use ExecuteSQLSelect() for SELECT query", SqlLog
        Exit Sub
    Else
        'db.openDb (path)
        Me.DbConnection.execute SQLstmt
    End If
End Sub


' Constructor
Private Sub Class_Initialize()
    Set DbConnection = New SQLiteDatabase
End Sub

' Deconstructor
Private Sub Class_Terminate()
    Set DbConnection = Nothing
End Sub
