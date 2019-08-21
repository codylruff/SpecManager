Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================
Function PushIQueryable(obj As IQueryable, Optional trans As SqlTransaction) As Long
' Push an object, that implements the IQueryable interface, to the database
    Dim transaction As SqlTransaction
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO user_actions " & _
            "(" & obj.GetValueLabels & ") " & _
            "VALUES (" & obj.GetValues & ")"
    trans.ExecuteSQL (SQLstmt)
    PushIQueryable = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushIQueryable = DB_PUSH_FAILURE
End Function

Function PushUserAction(action_string As String, Optional trans As SqlTransaction) As Long
    Dim transaction As SqlTransaction
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO user_actions " & _
            "(User, Time_Stamp, Action_Description, Work_Order, Material_Id, Spec_Type, Spec_Revision) " & _
            "VALUES (" & action_string & ")"
    trans.ExecuteSQL (SQLstmt)
    PushUserAction = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushUserAction = DB_PUSH_FAILURE
End Function

Function GetSpecification(ByVal material_id As String, ByVal spec_type As String, Optional trans As SqlTransaction) As DatabaseRecord
' Gets a single specifcation from the database
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    Logger.Log "Searching for " & spec_type & " : " & material_id
    SQLstmt = "SELECT * FROM standard_specifications " & _
              "LEFT JOIN materials ON standard_specifications.Material_Id = materials.Material_Id " & _
              "WHERE standard_specifications.Material_Id ='" & material_id & _
              "' AND " & "standard_specifications.Spec_Type ='" & spec_type & "'"

    Set GetSpecification = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function GetUser(ByVal Name As String, Optional trans As SqlTransaction) As DatabaseRecord
' Get a user from the database
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Searching for user name . . . "
    SQLstmt = "SELECT * FROM user_privledges " & _
              "WHERE Name ='" & Name & "'"

    Set GetUser = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function FlagUserForSecretChange(Name As String, Optional trans As SqlTransaction) As Long
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Updating user secret . . . "
    SQLstmt = "UPDATE user_privledges " & _
              "SET New_Secret_Required = " & 1 & _
              " WHERE Name ='" & Name & "'"
    trans.ExecuteSQL (SQLstmt)
    FlagUserForSecretChange = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
    FlagUserForSecretChange = DB_PUSH_FAILURE
End Function

Function ChangeUserSecret(Name As String, new_secret As String, Optional trans As SqlTransaction) As Long
' Get a user from the database
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Updating user secret . . . "
    SQLstmt = "UPDATE user_privledges " & _
              "SET Secret ='" & new_secret & "', New_Secret_Required = " & 0 & _
              " WHERE Name ='" & Name & "'"
    
    trans.ExecuteSQL (SQLstmt)
    ChangeUserSecret = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
    ChangeUserSecret = DB_PUSH_FAILURE
End Function

Function PushNewUser(new_user As Account, Optional trans As SqlTransaction) As Long
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO user_privledges " & _
              "(Name, Privledge_Level, Product_Line) " & _
              "VALUES ('" & new_user.Name & "', " & _
                      "'" & new_user.PrivledgeLevel & "', " & _
                      "'" & new_user.ProductLine & "')"

    trans.ExecuteSQL (SQLstmt)
    PushNewUser = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushNewUser = DB_PUSH_FAILURE
End Function

Function GetTemplateRecord(ByRef spec_type As String, Optional trans As SqlTransaction) As DatabaseRecord
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Searching for " & spec_type & " template . . . "
    SQLstmt = "SELECT * FROM template_specifications" & _
              " WHERE Spec_Type= '" & spec_type & "'"

    Set GetTemplateRecord = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function GetSpecificationRecords(ByRef MaterialId As String, Optional trans As SqlTransaction) As DatabaseRecord
' Get a record(s) from the database
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Searching for " & MaterialId & " specifications . . . "
    SQLstmt = "SELECT * FROM  standard_specifications " & _
              "LEFT JOIN materials ON standard_specifications.Material_Id = materials.Material_Id " & _
              "WHERE standard_specifications.Material_Id= '" & MaterialId & "'"
              
    Set GetSpecificationRecords = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function PushTemplate(ByRef Template As SpecificationTemplate, Optional trans As SqlTransaction)
' Push new template record
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO template_specifications " & _
              "(Time_Stamp, Properties_Json, Revision, Spec_Type, Product_Line) " & _
              "VALUES ('" & CStr(Now()) & "'," & _
                      "'" & Template.PropertiesJson & "', " & _
                      "'" & Template.Revision & "', " & _
                      "'" & Template.SpecType & "', " & _
                      "'" & Template.ProductLine & "')"
    trans.ExecuteSQL (SQLstmt)
    PushTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushTemplate = DB_PUSH_FAILURE
End Function

Function UpdateTemplate(ByRef Template As SpecificationTemplate, Optional trans As SqlTransaction)
' Push new template record
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "UPDATE template_specifications " & _
              "SET " & _
              "Time_Stamp ='" & CStr(Now()) & "', " & _
              "Properties_Json ='" & Template.PropertiesJson & "', " & _
              "Revision ='" & Template.Revision & "' " & _
              "WHERE Spec_Type ='" & Template.SpecType & "'"
    trans.ExecuteSQL (SQLstmt)
    UpdateTemplate = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL UPDATE Error : DbPushFailException", SqlLog
    UpdateTemplate = DB_PUSH_FAILURE
End Function

Function PushSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications", Optional trans As SqlTransaction) As Long
' Push a new records
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO " & tbl & " " & _
              "(Material_Id, Time_Stamp, Properties_Json, Revision, Spec_Type) " & _
              "VALUES ('" & spec.MaterialId & "', " & _
                      "'" & CStr(Now()) & "', " & _
                      "'" & spec.PropertiesJson & "', " & _
                      "'" & spec.Revision & "', " & _
                      "'" & spec.SpecType & "')"
    trans.ExecuteSQL (SQLstmt)
    PushSpec = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushSpec = DB_PUSH_FAILURE
End Function

Function DeleteTemplate(ByRef Template As SpecificationTemplate, Optional trans As SqlTransaction) As Long
' Deletes a record
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbDeleteFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM template_specifications " & _
              "WHERE Spec_Type ='" & Template.SpecType & "' AND Revision ='" & Template.Revision & "'"
    trans.ExecuteSQL (SQLstmt)
    DeleteTemplate = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    Logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
    DeleteTemplate = DB_DELETE_FAILURE
End Function

Function DeleteSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications", Optional trans As SqlTransaction) As Long
' Push a new records
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    On Error GoTo DbDeleteFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM " & tbl & " " & _
              "WHERE Material_Id ='" & spec.MaterialId & "' AND Revision ='" & spec.Revision & "'" & _
              " AND Spec_Type ='" & spec.SpecType & "'"
    trans.ExecuteSQL (SQLstmt)
    DeleteSpec = DB_DELETE_SUCCESS
    Exit Function
DbDeleteFailException:
    Logger.Log "SQL DELETE Error : DbDeleteFailException", SqlLog
    DeleteSpec = DB_DELETE_FAILURE
End Function

Function GetTemplateTypes(Optional trans As SqlTransaction) As DatabaseRecord
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Get all template types . . . "
    SQLstmt = "SELECT * FROM template_specifications"
    Set GetTemplateTypes = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function SelectAllSpecifications(spec_type As String, Optional trans As SqlTransaction) As VBA.Collection
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    Dim record As DatabaseRecord
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    'Logger.Log "Selecting all specifications . . . "
    SQLstmt = "SELECT * FROM standard_specifications WHERE Spec_Type ='" & spec_type & "'"
    Set record = trans.ExecuteSQLSelect(SQLstmt)
    Set SelectAllSpecifications = record.records
End Function

Public Function BeginTransaction(Optional path As String) As SqlTransaction
' Begin a transaction in sqlite
    Dim trans As SqlTransaction
    Set trans = Factory.CreateSqlTransaction(IIf(path = vbNullString, DATABASE_PATH, path))
    If trans.Begin <> DB_TRANSACTION_FAILURE Then
        Set BeginTransaction = trans
    Else
        Set BeginTransaction = Nothing
    End If
End Function

' Function trans.ExecuteSQLSelect(DatabaseRecord
' ' Returns an table like array
'     Dim record As DatabaseRecord
'     Set record = New DatabaseRecord
'     On Error GoTo NullRecordException
'     Logger.Log "-----------------------------------", SqlLog
'     Logger.Log SQLstmt, SqlLog
'     'db.openDb path
'     db.selectQry SQLstmt
'     record.Data = db.Data
'     record.Header = db.Header
'     record.Rows = db.NumRows
'     record.Columns = db.NumColumns
'     Set trans.ExecuteSQLSelect ption:
'     Logger.Log "SQL Select Error : NullRecordException!", SqlLog
'     Set trans.ExecuteSQLSelect ecuteSQL(ByRef db As IDatabase, SQLstmt As String)
' ' Performs update or insert querys returns error on select.
'     Logger.Log "-----------------------------------", SqlLog
'     Logger.Log SQLstmt, SqlLog
'     If Left(SQLstmt, 6) = "SELECT" Then
'         Logger.Log "Use trans.ExecuteSQLSelect(b
'     Else
'         'db.openDb (path)
'         db.Execute (SQLstmt)
'     End If
' End Sub
