Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================

Function GetUser(ByVal Name As String) As DatabaseRecord
' Get a user from the database
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Searching for user name . . . "
    SQLstmt = "SELECT * FROM user_privledges " & _
              "WHERE Name ='" & Name & "'"
    Set GetUser = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function PushNewUser(new_user As Account) As Long
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    SQLstmt = "INSERT INTO user_privledges " & _
              "(Name, Privledge_Level, Product_Line) " & _
              "VALUES ('" & new_user.Name & "', " & _
                      "'" & new_user.PrivledgeLevel & "', " & _
                      "'" & new_user.ProductLine & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    PushNewUser = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException"
    PushNewUser = DB_PUSH_FAILURE
End Function

Function GetTemplateRecord(ByRef spec_type As String) As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Searching for " & spec_type & " template . . . "
    SQLstmt = "SELECT * FROM template_specifications" & _
              " WHERE Spec_Type= '" & spec_type & "'"
    Set GetTemplateRecord = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function GetSpecificationRecords(ByRef MaterialId As String) As DatabaseRecord
' Get a record(s) from the database
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Searching for " & MaterialId & " specifications . . . "
    SQLstmt = "SELECT * FROM  standard_specifications" & _
              " WHERE Material_Id= '" & MaterialId & "'"
    Set GetSpecificationRecords = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function PushTemplate(ByRef template As SpecificationTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO template_specifications " & _
              "(Time_Stamp, Properties_Json, Revision, Spec_Type, Product_Line) " & _
              "VALUES ('" & CStr(Now()) & "'," & _
                      "'" & template.PropertiesJson & "', " & _
                      "'" & template.Revision & "', " & _
                      "'" & template.SpecType & "', " & _
                      "'" & template.ProductLine & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    PushTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException"
    PushTemplate = DB_PUSH_FAILURE
End Function

Function UpdateTemplate(ByRef template As SpecificationTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "UPDATE template_specifications " & _
              "SET " & _
              "Time_Stamp ='" & CStr(Now()) & "', " & _
              "Properties_Json ='" & template.PropertiesJson & "', " & _
              "Revision ='" & template.Revision & "' " & _
              "WHERE Spec_Type ='" & template.SpecType & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    UpdateTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL UPDATE Error : DbPushFailException"
    UpdateTemplate = DB_PUSH_FAILURE
End Function

Function PushSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications") As Long
' Push a new records
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO " & tbl & " " & _
              "(Material_Id, Time_Stamp, Properties_Json, Tolerances_Json, Revision, Spec_Type) " & _
              "VALUES ('" & spec.MaterialId & "', " & _
                      "'" & CStr(Now()) & "', " & _
                      "'" & spec.PropertiesJson & "', " & _
                      "'" & spec.TolerancesJson & "', " & _
                      "'" & spec.Revision & "', " & _
                      "'" & spec.SpecType & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    PushSpec = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException"
    PushSpec = DB_PUSH_FAILURE
End Function

Function DeleteTemplate(ByRef template As SpecificationTemplate) As Long
' Deletes a record
    Dim SQLstmt As String
    On Error GoTo DbDeleteFailException
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM template_specifications " & _
              "WHERE Spec_Type ='" & template.SpecType & "' AND Revision ='" & template.Revision & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    DeleteTemplate = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    Logger.Log "SQL DELETE Error : DbDeleteFailException"
    DeleteTemplate = DB_DELETE_FAILURE
End Function

Function DeleteSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications") As Long
' Push a new records
    Dim SQLstmt As String
    On Error GoTo DbDeleteFailException
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM " & tbl & " " & _
              "WHERE Material_Id ='" & spec.MaterialId & "' AND Revision ='" & spec.Revision & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    DeleteSpec = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    Logger.Log "SQL DELETE Error : DbDeleteFailException"
    DeleteSpec = DB_DELETE_FAILURE
End Function

Function GetTemplateTypes() As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    Logger.Log "Get all template types . . . "
    SQLstmt = "SELECT * FROM template_specifications"
    Set GetTemplateTypes = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function SelectAllSpecifications(spec_type As String) As VBA.Collection
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    ' build the sql query
    Logger.Log "Selecting all specifications . . . "
    SQLstmt = "SELECT * FROM standard_specifications WHERE Spec_Type ='" & spec_type & "'"
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
    record.SetDictionary
    Set SelectAllSpecifications = record.records
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
    record.Data = db.Data
    record.Header = db.Header
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
        db.Execute (SQLstmt)
    End If
End Sub
