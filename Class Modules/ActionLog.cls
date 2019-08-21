VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type UserAction
    User As String
    Time_Stamp As String
    description As String
    work_order As String
    MaterialId As String
    SpecType As String
    Revision As String
End Type

Private log_enabled As Boolean

Public Sub Enable(): log_enabled = True: End Sub
Public Sub Disable(): log_enabled = False: End Sub

Private Sub Class_Initialize()
' Constructor
    Enable
End Sub

Private Function ToString(action As UserAction) As String
    ToString = "'" & action.User & "', " & _
               "'" & action.Time_Stamp & "', " & _
               "'" & action.description & "', " & _
               "'" & action.work_order & "', " & _
               "'" & action.MaterialId & "', " & _
               "'" & action.SpecType & "', " & _
               "'" & action.Revision & "'"
End Function

Private Property Get Username() As String
    Username = UCase(VBA.Environ("Username"))
End Property

Public Sub DocumentPrinted(spec As Specification, work_order As String)
    Dim ret_val As Long
    If Not log_enabled Then Exit Sub
    With spec
        VerifyActionLogged DataAccess.PushUserAction(ToString(CreateUserAction(Username, "Document Printed", _
                                                    .MaterialId, .SpecType, .Revision, work_order)))
    End With
End Sub

Public Sub LogUserAction(Optional description As String)
' Generic User Action description only
    If Not log_enabled Then Exit Sub
    VerifyActionLogged DataAccess.PushUserAction(ToString(CreateUserAction(Username, description)))
End Sub


Public Sub CrudOnSpecification(spec As Specification, message As String)
' Log any crud operation on a specification
    If Not log_enabled Then Exit Sub
    With spec
        VerifyActionLogged DataAccess.PushUserAction(ToString(CreateUserAction(Username, message, .MaterialId, .SpecType, .Revision)))
    End With
End Sub

Public Sub CrudOnTemplate(Template As SpecificationTemplate, message As String)
' Log any crud operation on a template
    If Not log_enabled Then Exit Sub
    With Template
        VerifyActionLogged DataAccess.PushUserAction(ToString(CreateUserAction(Username, message, "Template", .SpecType, .Revision)))
    End With
End Sub

Private Function CreateUserAction(user_name As String, action_description As String, _
            Optional material_id As String, Optional spec_type As String, _
                Optional rev As String, Optional work_order As String) As UserAction
' Creates a user action for logging into the database.
    Dim action As UserAction
    With action
        .User = user_name
        .Time_Stamp = Format(CStr(Now()), "dd-MMM-yyyy HH:nn:ss")
        .description = action_description
        .work_order = work_order
        .MaterialId = material_id
        .SpecType = spec_type
        .Revision = rev
    End With
    CreateUserAction = action
End Function

Private Sub VerifyActionLogged(ret_val As Long)
' logs the return value from the db when an action is written
    If ret_val <> DB_PUSH_SUCCESS Then
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        Logger.Log "Action Was Not Saved!"
    Else
        Logger.Log "Data Access returned: " & ret_val, DebugLog
        Logger.Log "Action Succesfully Saved."
    End If
End Sub
