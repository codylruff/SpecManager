Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================

Function PushUserAction(action As UserAction) As Long
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    SQLstmt = "INSERT INTO user_actions " & _
              "(User, Time_Stamp, Action_Description, Work_Order, Material_Id, Spec_Type, Spec_Revision) " & _
              "VALUES ('" & action.User & "', " & _
                      "'" & action.Time_Stamp & "', " & _
                      "'" & action.Description & "', " & _
                      "'" & action.work_order & "', " & _
                      "'" & action.spec.MaterialId & "', " & _
                      "'" & action.spec.SpecType & "', " & _
                      "'" & action.spec.Revision & "')"

    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    PushUserAction = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    App.logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushUserAction = DB_PUSH_FAILURE
End Function

Function GetSpecification(ByVal material_id As String, ByVal spec_type As String) As DatabaseRecord
' Gets a single specifcation from the database
    Dim SQLstmt As String
    App.logger.Log "Searching for " & spec_type & " : " & material_id
    SQLstmt = "SELECT * FROM standard_specifications " & _
              "LEFT JOIN materials ON standard_specifications.Material_Id = materials.Material_Id " & _
              "WHERE standard_specifications.Material_Id ='" & material_id & _
              "' AND " & "standard_specifications.Spec_Type ='" & spec_type & "'"

    Set GetSpecification = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function GetUser(ByVal Name As String) As DatabaseRecord
' Get a user from the database
    Dim SQLstmt As String
    ' build the sql query
    App.logger.Log "Searching for user name . . . "
    SQLstmt = "SELECT * FROM user_privledges " & _
              "WHERE Name ='" & Name & "'"

    Set GetUser = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function FlagUserForSecretChange(Name As String) As Long
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' build the sql query
    App.logger.Log "Updating user secret . . . "
    SQLstmt = "UPDATE user_privledges " & _
              "SET New_Secret_Required = " & 1 & _
              " WHERE Name ='" & Name & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    FlagUserForSecretChange = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    App.logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
    FlagUserForSecretChange = DB_PUSH_FAILURE
End Function

Function ChangeUserSecret(Name As String, new_secret As String) As Long
' Get a user from the database
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' build the sql query
    App.logger.Log "Updating user secret . . . "
    SQLstmt = "UPDATE user_privledges " & _
              "SET Secret ='" & new_secret & "', New_Secret_Required = " & 0 & _
              " WHERE Name ='" & Name & "'"
    
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    ChangeUserSecret = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    App.logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
    ChangeUserSecret = DB_PUSH_FAILURE
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
    App.logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushNewUser = DB_PUSH_FAILURE
End Function

Function GetTemplateRecord(ByRef spec_type As String) As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    App.logger.Log "Searching for " & spec_type & " template . . . "
    SQLstmt = "SELECT * FROM template_specifications" & _
              " WHERE Spec_Type= '" & spec_type & "'"

    Set GetTemplateRecord = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function GetSpecificationRecords(ByRef MaterialId As String) As DatabaseRecord
' Get a record(s) from the database
    Dim SQLstmt As String
    ' build the sql query
    App.logger.Log "Searching for " & MaterialId & " specifications . . . "
    SQLstmt = "SELECT * FROM  standard_specifications " & _
              "LEFT JOIN materials ON standard_specifications.Material_Id = materials.Material_Id " & _
              "WHERE standard_specifications.Material_Id= '" & MaterialId & "'"
              
    Set GetSpecificationRecords = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function PushTemplate(ByRef Template As SpecificationTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO template_specifications " & _
              "(Time_Stamp, Properties_Json, Revision, Spec_Type, Product_Line) " & _
              "VALUES ('" & CStr(Now()) & "'," & _
                      "'" & Template.PropertiesJson & "', " & _
                      "'" & Template.Revision & "', " & _
                      "'" & Template.SpecType & "', " & _
                      "'" & Template.ProductLine & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    PushTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    App.logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushTemplate = DB_PUSH_FAILURE
End Function

Function UpdateTemplate(ByRef Template As SpecificationTemplate)
' Push new template record
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    ' Create SQL statement from objects
    SQLstmt = "UPDATE template_specifications " & _
              "SET " & _
              "Time_Stamp ='" & CStr(Now()) & "', " & _
              "Properties_Json ='" & Template.PropertiesJson & "', " & _
              "Revision ='" & Template.Revision & "' " & _
              "WHERE Spec_Type ='" & Template.SpecType & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    UpdateTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    App.logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
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
    App.logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushSpec = DB_PUSH_FAILURE
End Function

Function DeleteTemplate(ByRef Template As SpecificationTemplate) As Long
' Deletes a record
    Dim SQLstmt As String
    On Error GoTo DbDeleteFailException
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM template_specifications " & _
              "WHERE Spec_Type ='" & Template.SpecType & "' AND Revision ='" & Template.Revision & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    DeleteTemplate = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    App.logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
    DeleteTemplate = DB_DELETE_FAILURE
End Function

Function DeleteSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications") As Long
' Push a new records
    Dim SQLstmt As String
    On Error GoTo DbDeleteFailException
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM " & tbl & " " & _
              "WHERE Material_Id ='" & spec.MaterialId & "' AND Revision ='" & spec.Revision & "'" & _
              " AND Spec_Type ='" & spec.SpecType & "'"
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt
    DeleteSpec = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    App.logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
    DeleteSpec = DB_DELETE_FAILURE
End Function

Function GetTemplateTypes() As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    App.logger.Log "Get all template types . . . "
    SQLstmt = "SELECT * FROM template_specifications"
    Set GetTemplateTypes = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
End Function

Function SelectAllSpecifications(spec_type As String) As VBA.Collection
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    ' build the sql query
    App.logger.Log "Selecting all specifications . . . "
    SQLstmt = "SELECT * FROM standard_specifications WHERE Spec_Type ='" & spec_type & "'"
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, DATABASE_PATH, SQLstmt)
    Set SelectAllSpecifications = record.records
End Function

Private Function BeginTransaction() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, "BEGIN TRANSACTION"
    BeginTransaction = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL BEGIN Error : DbTransactionFailException", SqlLog
    BeginTransaction = DB_TRANSACTION_FAILURE
End Function

Private Function RollbackTransaction() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, "ROLLBACK TRANSACTION"
    RollbackTransaction = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL ROLLBACK Error : DbTransactionFailException", SqlLog
    RollbackTransaction = DB_TRANSACTION_FAILURE
End Function

Private Function CommitTransaction() As Long
' Begin a transaction in sqlite
    On Error GoTo DbTransactionFailException
    ExecuteSQL Factory.CreateSQLiteDatabase, DATABASE_PATH, "COMMIT TRANSACTION"
    CommitTransaction = DB_TRANSACTION_SUCCESS
DbTransactionFailException:
    App.logger.Log "SQL COMMIT Error : DbTransactionFailException", SqlLog
    CommitTransaction = DB_TRANSACTION_FAILURE
End Function

Private Function ExecuteSQLSelect(db As IDatabase, path As String, SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record As DatabaseRecord
    Set record = New DatabaseRecord
    On Error GoTo NullRecordException
    App.logger.Log "-----------------------------------", SqlLog
    App.logger.Log SQLstmt, SqlLog
    db.openDb path
    db.selectQry SQLstmt
    record.Data = db.Data
    record.Header = db.Header
    record.Rows = db.NumRows
    record.Columns = db.NumColumns
    Set ExecuteSQLSelect = record
    Exit Function
NullRecordException:
    App.logger.Log "SQL Select Error : NullRecordException!", SqlLog
    Set ExecuteSQLSelect = New DatabaseRecord
End Function

Private Sub ExecuteSQL(db As IDatabase, path As String, SQLstmt As String)
' Performs update or insert querys returns error on select.
    App.logger.Log "-----------------------------------", SqlLog
    App.logger.Log SQLstmt, SqlLog
    If Left(SQLstmt, 6) = "SELECT" Then
        App.logger.Log "Use ExecuteSQLSelect() for SELECT query", SqlLog
        Exit Sub
    Else
        db.openDb (path)
        db.Execute (SQLstmt)
    End If
End Sub
