Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================
Function PushIQueryable(ByRef obj As IQueryable, Table As String, Optional ByRef trans As SqlTransaction) As Long
' Push an object, that implements the IQueryable interface, to the database
    Dim transaction As SqlTransaction
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO " & Table & _
            "(" & obj.GetValueLabels & ") " & _
            "VALUES (" & obj.GetValues & ")"
    trans.ExecuteSQL (SQLstmt)
    PushIQueryable = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", RuntimeLog
    Logger.Log SQLstmt, RuntimeLog
    PushIQueryable = DB_PUSH_FAILURE
End Function

Function PushValue(ByVal key_name As String, ByVal key_id As Variant, ByVal column_name As String, ByVal column_value As Variant, Table As String, Optional ByRef trans As SqlTransaction) As Long
' Push a value to the database
    Dim transaction As SqlTransaction
    Dim SQLstmt As String
    On Error GoTo DbPushFailException
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO " & Table & _
            "(" & key_name & ", " & column_name & ") " & _
            "VALUES ('" & key_id & "', '" & column_value & "')"
    trans.ExecuteSQL (SQLstmt)
    PushValue = DB_PUSH_SUCCESS
    Exit Function
DbPushFailException:
    Logger.Log "SQL INSERT Error : DbPushFailException", SqlLog
    PushValue = DB_PUSH_FAILURE
End Function

Function GetColumn(ByVal key_name As String, ByVal key_id As String, ByVal column_name As String, ByVal tbl As String, Optional ByRef trans As SqlTransaction) As DataFrame
' Gets a single specifcation from the database
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    Logger.Log "Searching for " & column_name & ", for : " & key_id
    SQLstmt = "SELECT " & column_name & " FROM " & tbl & _
              " WHERE " & key_name & " ='" & key_id & "'"

    Set GetColumn = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function GetSpecification(ByVal material_id As String, ByVal spec_type As String, Optional ByRef trans As SqlTransaction) As DataFrame
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

Function GetUser(ByVal Name As String, Optional ByRef trans As SqlTransaction) As DataFrame
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

Function FlagUserForSecretChange(Name As String, Optional ByRef trans As SqlTransaction) As Long
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

Function ChangeUserSecret(Name As String, new_secret As String, Optional ByRef trans As SqlTransaction) As Long
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

Function GetTemplateRecord(ByRef spec_type As String, Optional ByRef trans As SqlTransaction) As DataFrame
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

Function GetSpecificationRecords(ByRef MaterialId As String, Optional ByRef trans As SqlTransaction) As DataFrame
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

Function UpdateTemplate(ByRef Template As SpecificationTemplate, Optional ByRef trans As SqlTransaction)
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

Function DeleteTemplate(ByRef Template As SpecificationTemplate, Optional ByRef trans As SqlTransaction) As Long
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

Function DeleteSpec(ByRef spec As Specification, Optional tbl As String = "standard_specifications", Optional ByRef trans As SqlTransaction) As Long
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

Function GetTemplateTypes(Optional ByRef trans As SqlTransaction) As DataFrame
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

Function SelectAllSpecifications(spec_type As String, Optional ByRef trans As SqlTransaction) As VBA.Collection
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
    Dim df As DataFrame
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    'Logger.Log "Selecting all specifications . . . "
    SQLstmt = "SELECT * FROM standard_specifications WHERE Spec_Type ='" & spec_type & "'"
    Set df = trans.ExecuteSQLSelect(SQLstmt)
    Set SelectAllSpecifications = df.records
End Function

Function SelectAllWhere(wheres As Variant, vals As Variant, Table As String, Optional fields As String = "*", Optional ByRef trans As SqlTransaction) As DataFrame
' Selects all records matching criteria
    Dim conditions As String
    Dim SQLstmt As String
    Dim i As Long
    Dim transaction As SqlTransaction
    Dim df As DataFrame
    If Not UBound(wheres) = UBound(vals) Then
        Logger.Log "wheres and vals must be the same length!"
        Exit Function
    ElseIf Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    conditions = " WHERE " & wheres(0) & "='" & vals(0) & "'"
    If UBound(wheres) > 1 Then
        For i = 1 To UBound(wheres) - 1
            conditions = conditions & ", AND " & wheres(i) & "='" & vals(i) & "'"
        Next i
    End If
    SQLstmt = "SELECT " & fields & " FROM " & Table & conditions
    Set df = trans.ExecuteSQLSelect(SQLstmt)
    Set SelectAllWhere = df
End Function

Public Function BeginTransaction(Optional path As String) As SqlTransaction
' Begin a transaction in sqlite
    Dim trans As SqlTransaction
    Set trans = Factory.CreateSqlTransaction(IIf(path = nullstr, DATABASE_PATH, path))
    If trans.Begin <> DB_TRANSACTION_FAILURE Then
        Set BeginTransaction = trans
    Else
        Set BeginTransaction = Nothing
    End If
End Function
