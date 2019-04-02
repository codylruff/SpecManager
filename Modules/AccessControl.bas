Attribute VB_Name = "AccessControl"
Option Explicit

Private user As Account

Function Account_Initialize() As Account
    If GetPrivledges(UCase(VBA.Environ("Username"))) <> DB_SELECT_SUCCESS Then
        Logger.Log "Creating new user " & VBA.Environ("Username")
        Set Account_Initialize = AccessControl.AutoAddNewUser
    Else
        Logger.Log "Selected User : {" & user.ToString & "}"
        Set Account_Initialize = user
    End If
End Function

Function GetPrivledges(Name As String) As Long
    Dim record As DatabaseRecord
    Set record = DataAccess.GetUser(Name)
    If record Is Nothing Then
        GetPrivledges = DB_SELECT_FAILURE
    Else
        Set user = New Account
        record.SetDictionary
        user.Name = record.Fields("Name")
        user.PrivledgeLevel = record.Fields("Privledge_Level")
        user.ProductLine = record.Fields("Product_Line")

        GetPrivledges = DB_SELECT_SUCCESS
    End If
End Function

Function AutoAddNewUser() As Account
    Dim new_user As Account
    Set new_user = New Account
    new_user.Name = Environ("Username")
    ' Users have read only access by default
    new_user.PrivledgeLevel = USER_READONLY
    Logger.Log DataAccess.PushNewUser(new_user)
    Set AutoAddNewUser = new_user
End Function
