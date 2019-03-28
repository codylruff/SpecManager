Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================

Function GetTemplateRecord(ByRef spec_type As String) As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Searching for " & spec_type & " template . . . "
    SQLstmt = "SELECT * FROM template_specifications" & _
              " WHERE Spec_Type= '" & spec_type & "'"
    Set GetTemplateRecord = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
End Function

Function GetSpecificationRecords(ByRef MaterialId As String) As DatabaseRecord
' Get a record(s) from the database
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Searching for " & MaterialId & " specifications . . . "
    SQLstmt = "SELECT * FROM  standard_specifications" & _
              " WHERE Material_Id= '" & MaterialId & "'"
    Set GetSpecificationRecords = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
End Function

Function PushTemplate(ByRef template As SpecTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO template_specifications " & _
              "(Time_Stamp, Properties_Json, Revision, Spec_Type) " & vbNewLine & _
              "VALUES ('" & CStr(Now()) & "'," & _
                      "'" & template.PropertiesJson & "', " & _
                      "'" & template.Revision & "', " & _
                      "'" & template.SpecType & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
    PushTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException"
    PushTemplate = DB_PUSH_FAILURE
End Function

Function UpdateTemplate(ByRef template As SpecTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "UPDATE template_specifications " & vbNewLine & _
              "SET " & _
              "Time_Stamp ='" & CStr(Now()) & "', " & _
              "Properties_Json ='" & template.PropertiesJson & "', " & _
              "Revision ='" & template.Revision & "'" & vbNewLine & _
              "WHERE Spec_Type ='" & template.SpecType & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
    UpdateTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL UPDATE Error : DbPushFailException"
    UpdateTemplate = DB_PUSH_FAILURE
End Function

Function PushSpec(ByRef spec As Specification) As Long
' Push a new records
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO standard_specifications " & _
              "(Material_Id, Time_Stamp, Properties_Json, Tolerances_Json, Revision, Spec_Type) " & vbNewLine & _
              "VALUES ('" & spec.MaterialId & "', " & _
                      "'" & CStr(Now()) & "', " & _
                      "'" & spec.PropertiesJson & "', " & _
                      "'" & spec.TolerancesJson & "', " & _
                      "'" & spec.Revision & "', " & _
                      "'" & spec.SpecType & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
    PushSpec = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException"
    PushSpec = DB_PUSH_FAILURE
End Function

Function GetTemplateTypes() As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Get all template types . . . "
    SQLstmt = "SELECT * FROM template_specifications"
    Set GetTemplateTypes = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
End Function

Private Function ExecuteSQLSelect(db As IDatabase, path As String, SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record As DatabaseRecord
    Set record = New DatabaseRecord
    On Error GoTo NullRecordException
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    db.openDb path
    db.selectQry SQLstmt
    record.data = db.data
    record.header = db.header
    record.rows = db.NumRows
    record.columns = db.NumColumns
    Set ExecuteSQLSelect = record
    Exit Function
NullRecordException:
    Logger.Log "SQL Select Error : NullRecordException!"
    Set ExecuteSQLSelect = New DatabaseRecord
End Function

Private Sub ExecuteSQL(db As IDatabase, path As String, SQLstmt As String)
' Performs update or insert querys returns error on select.
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    If Left(SQLstmt, 6) = "SELECT" Then
        Logger.Log ("Use ExecuteSQLSelect() for SELECT query")
        Exit Sub
    Else
        db.openDb (path)
        db.execute (SQLstmt)
    End If
End Sub

Public Sub exampleSelect()
  '----------------------------------------------'
  Dim qry As Variant
  Dim db As IDatabase
  '----------------------------------------------'
  Set db = Factory.CreateSQLiteDatabase
  db.openDb SQLITE_PATH
  db.selectQry "select * from standard_specifications " 'limit 100"  'faz o select na base de dados e printa as colunas do print'
  '----------------------------------------------'
  DbTest.Range(Cells(1, 1), Cells(1, db.NumColumns)).value = db.header 'cola cabecalho
  DbTest.Range(Cells(2, 1), Cells(db.NumRows + 1, db.NumColumns)).value = db.data 'cola os dados
  '----------------------------------------------'
End Sub

