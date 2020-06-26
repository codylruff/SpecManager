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

Private log_enabled As Boolean

Public Sub enable(): log_enabled = True: End Sub
Public Sub Disable(): log_enabled = False: End Sub

Private Sub Class_Initialize()
' Constructor
    enable
End Sub

Private Property Get Username() As String
    Username = UCase(VBA.Environ("Username"))
End Property

Public Sub DocumentPrinted(doc As Document, work_order As String)
    Dim ret_val As Long
    If Not log_enabled Then Exit Sub
    With doc
        VerifyActionLogged DataAccess.PushIQueryable(CreateUserAction(Username, "Document Printed", _
                                                    .MaterialId, .SpecType, .Revision, work_order), "user_actions")
    End With
End Sub

Public Sub LogUserAction(Optional description As String)
' Generic User Action description only
    If Not log_enabled Then Exit Sub
    VerifyActionLogged DataAccess.PushIQueryable(CreateUserAction(Username, description), "user_actions")
End Sub


Public Sub CrudOnDocument(doc As Document, message As String)
' Log any crud operation on a specification
    If Not log_enabled Then Exit Sub
    With doc
        VerifyActionLogged DataAccess.PushIQueryable(CreateUserAction(Username, message, .MaterialId, .SpecType, .Revision), "user_actions")
    End With
End Sub

Public Sub CrudOnTemplate(Template As Template, message As String)
' Log any crud operation on a template
    If Not log_enabled Then Exit Sub
    With Template
        VerifyActionLogged DataAccess.PushIQueryable(CreateUserAction(Username, message, "Template", .SpecType, .Revision), "user_actions")
    End With
End Sub

Private Function CreateUserAction(user_name As String, action_description As String, _
            Optional material_id As String, Optional spec_type As String, _
                Optional rev As String, Optional work_order As String) As UserAction
' Creates a user action for logging into the database.
    Dim action As New UserAction
    With action
        .User = user_name
        .Time_Stamp = Format(CStr(Now()), "dd-MMM-yyyy HH:nn:ss")
        .description = action_description
        .work_order = work_order
        .MaterialId = material_id
        .SpecType = spec_type
        .Revision = rev
    End With
    Set CreateUserAction = action
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
