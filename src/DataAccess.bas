Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================
'TODO Add Error Handling to the remaining functions in the module.
Function PushIQueryable(ByRef obj As IQueryable, Table As String, Optional ByRef trans As SqlTransaction, _
                        Optional db_path As String = DATABASE_PATH) As Long
' Push an object, that implements the IQueryable interface, to the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(db_path)
    End If
    SQLstmt = "INSERT INTO " & Table & _
            "(" & obj.GetValueLabels & ") " & _
            "VALUES (" & obj.GetValues & ")"
    trans.ExecuteSQL (SQLstmt)
    PushIQueryable = DB_PUSH_SUCCESS
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL INSERT Error : DbPushFailException", RuntimeLog
    Logger.Log SQLstmt, RuntimeLog
    PushIQueryable = DB_PUSH_ERR
    GoTo Finally
End Function

Function PushValue(ByVal key_name As String, ByVal key_id As Variant, ByVal column_name As String, _
                    ByVal column_value As Variant, Table As String, Optional ByRef trans As SqlTransaction) As Long
' Push a value to the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    SQLstmt = "INSERT INTO " & Table & _
            "(" & key_name & ", " & column_name & ") " & _
            "VALUES ('" & key_id & "', '" & column_value & "')"
    trans.ExecuteSQL (SQLstmt)
    PushValue = DB_PUSH_SUCCESS
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL INSERT Error : DbPushFailException", RuntimeLog
    PushValue = DB_PUSH_ERR
    GoTo Finally
End Function

Function GetColumn(ByVal key_name As String, ByVal key_id As String, ByVal column_name As String, _
                    ByVal tbl As String, Optional ByRef trans As SqlTransaction) As DataFrame
' Gets a single specifcation from the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    Logger.Log "Searching for " & column_name & ", for : " & key_id
    SQLstmt = "SELECT " & column_name & " FROM " & tbl & _
              " WHERE " & key_name & " ='" & key_id & "'"

    Set GetColumn = trans.ExecuteSQLSelect(SQLstmt)
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Set GetColumn = Nothing
    Logger.Error "DataAccess.GetColumn() Failed"
End Function

Function GetDocument(ByVal material_id As String, ByVal spec_type As String, machine_id As String, _
                        Optional ByRef trans As SqlTransaction) As DataFrame
' Gets a single specifcation document from the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    Logger.Log "Searching for " & spec_type & " : " & material_id
    SQLstmt = "SELECT * FROM standard_specifications " & _
              "LEFT JOIN materials ON standard_specifications.Material_Id = materials.Material_Id " & _
              "WHERE standard_specifications.Material_Id ='" & material_id & _
              "' AND " & "standard_specifications.Spec_Type ='" & spec_type & "'" & _
              " AND " & "standard_specifications.Machine_Id ='" & machine_id & "'"

    Set GetDocument = trans.ExecuteSQLSelect(SQLstmt)
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Error "DataAccess.GetDocument() Failed"
    Set GetDocument = Nothing
    GoTo Finally
End Function

Function GetUser(ByVal Name As String, Optional ByRef trans As SqlTransaction) As DataFrame
' Get a user from the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Searching for user name . . . "
    SQLstmt = "SELECT * FROM user_privledges " & _
              "WHERE Name ='" & Name & "'"

    Set GetUser = trans.ExecuteSQLSelect(SQLstmt)
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Set GetUser = Nothing
    Logger.Error "DataAccess.GetUser Failed"
    GoTo Finally
End Function

Function FlagUserForSecretChange(ByVal Name As String, Optional ByRef trans As SqlTransaction) As Long
' Flags a user in the database as needing a password change.
On Error GoTo Catch
    Dim SQLstmt As String
    Dim transaction As SqlTransaction
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
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL UPDATE Error : DbPushFailException", RuntimeLog
    FlagUserForSecretChange = DB_PUSH_ERR
    GoTo Finally
End Function

Function ChangeUserSecret(ByVal Name As String, ByVal new_secret_hash As String, Optional ByRef trans As SqlTransaction) As Long
' Get a user from the database
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Updating user secret (hash). . . "
    SQLstmt = "UPDATE user_privledges " & _
              "SET Secret ='" & new_secret_hash & "', New_Secret_Required = " & 0 & _
              " WHERE Name ='" & Name & "'"
    
    trans.ExecuteSQL (SQLstmt)
    ChangeUserSecret = DB_PUSH_SUCCESS
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL UPDATE Error : DbPushFailException", RuntimeLog
    ChangeUserSecret = DB_PUSH_ERR
    GoTo Finally
End Function

Function GetTemplateRecord(ByVal spec_type As String, Optional ByRef trans As SqlTransaction) As DataFrame
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

Function GetDocumentRecords(ByVal MaterialId As String, Optional ByRef trans As SqlTransaction) As DataFrame
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
              
    Set GetDocumentRecords = trans.ExecuteSQLSelect(SQLstmt)
End Function

Function UpdateTemplate(ByVal Template As Template, Optional ByRef trans As SqlTransaction)
' Push new template record
On Error GoTo Catch
    Dim SQLstmt As String
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
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL UPDATE Error : DbPushFailException", RuntimeLog
    UpdateTemplate = DB_PUSH_ERR
    GoTo Finally
End Function

Function DeleteTemplate(ByVal Template As Template, Optional ByRef trans As SqlTransaction) As Long
' Deletes a record
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM template_specifications " & _
              "WHERE Spec_Type ='" & Template.SpecType & "' AND Revision ='" & Template.Revision & "'"
    trans.ExecuteSQL (SQLstmt)
    DeleteTemplate = DB_DELETE_SUCCESS
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL DELETE Error : DbDeleteFailException", RuntimeLog
    DeleteTemplate = DB_DELETE_ERR
    GoTo Finally
End Function

Function DeleteSpec(ByRef doc As Document, ByVal machine_id As String, Optional ByVal tbl As String = "standard_specifications", Optional ByRef trans As SqlTransaction) As Long
' Push a new records
On Error GoTo Catch
    Dim SQLstmt As String
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' Create SQL statement from objects
    SQLstmt = "DELETE FROM " & tbl & " " & _
              "WHERE Material_Id ='" & doc.MaterialId & "' AND Revision ='" & doc.Revision & "'" & _
              " AND Spec_Type ='" & doc.SpecType & "'" & _
              " AND " & "Machine_Id ='" & machine_id & "'"

    trans.ExecuteSQL (SQLstmt)
    DeleteSpec = DB_DELETE_SUCCESS
Finally:
    Set doc = Nothing
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL DELETE Error : DbDeleteFailException", RuntimeLog
    DeleteSpec = DB_DELETE_ERR
    GoTo Finally
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

Function SelectAllDocuments(ByVal spec_type As String, Optional ByRef trans As SqlTransaction) As VBA.Collection
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
    Set SelectAllDocuments = df.records
End Function

Private Function SelectAllWhere(ByVal wheres As Variant, ByVal vals As Variant, ByVal Table As String, Optional ByVal Fields As String = "*", Optional ByRef trans As SqlTransaction) As DataFrame
' Selects all records matching criteria
'FIXME This does not work correctly it is set to private until it does.
    Dim conditions As String
    Dim SQLstmt As String
    Dim i As Long
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
    SQLstmt = "SELECT " & Fields & " FROM " & Table & conditions
    Set df = trans.ExecuteSQLSelect(SQLstmt)
    Set SelectAllWhere = df
End Function

Public Function BeginTransaction(Optional ByVal path As String) As SqlTransaction
' Begin a transaction in sqlite
' REVIEW Does this get called anywhere?
    Dim trans As SqlTransaction
    Set trans = Factory.CreateSqlTransaction(IIf(path = nullstr, DATABASE_PATH, path))
    If trans.Begin <> DB_TRANSACTION_ERR Then
        Set BeginTransaction = trans
    Else
        Set BeginTransaction = Nothing
    End If
End Function

Function UpdateValue(ByVal key_name As String, ByVal key_id As Variant, ByVal column_name As String, ByVal column_value As Variant, Table As String, Optional ByRef trans As SqlTransaction) As Long
    Dim SQLstmt As String
On Error GoTo Catch
    If Utils.IsNothing(trans) Then
        Set trans = Factory.CreateSqlTransaction(DATABASE_PATH)
    End If
    ' build the sql query
    Logger.Log "Updating " & key_id & " . . . "
    SQLstmt = "UPDATE " & Table & " " & _
              "SET " & column_name & " = '" & column_value & "' " & _
              "WHERE " & key_name & " ='" & key_id & "'"
    Debug.Print SQLstmt
    trans.ExecuteSQL (SQLstmt)
    UpdateValue = DB_PUSH_SUCCESS
Finally:
    Set trans = Nothing
    Exit Function
Catch:
    Logger.Log "SQL UPDATE Error : DbPushFailException", RuntimeLog
    UpdateValue = DB_PUSH_ERR
    GoTo Finally
End Function
