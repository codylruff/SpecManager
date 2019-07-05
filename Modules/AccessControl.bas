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
    ' obsoleted
    If record.Fields Is Nothing Then
        GetPrivledges = DB_SELECT_FAILURE
    Else
        Set user = New Account
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
    new_user.ChangeSetting "name", Environ("Username")
    ' Users have read only access by default
    new_user.PrivledgeLevel = USER_READONLY
    new_user.ChangeSetting "privledge_level", USER_READONLY
    new_user.ProductLine = "User"
    new_user.ChangeSetting "product_line", "User"
    new_user.SaveUserJson
    Logger.Log DataAccess.PushNewUser(new_user)
    Set AutoAddNewUser = new_user
End Function

Public Sub ConfigControl()
'Initializes the password form for config access.
    Dim w As Window
    If Environ("UserName") <> "CRuff" Then
        formPassword.Show
    Else
        Application.DisplayAlerts = True
        If Windows.Count <> 1 Then
            For Each w In Windows
                If w.Parent.Name = ThisWorkbook.Name Then w.Visible = True
            Next w
        Else
            Application.Visible = True
        End If
        'Application.VBE.MainWindow.Visible = True
        GuiCommands.ShowAllSheets SAATI_Data_Manager.ThisWorkbook
        'Application.SendKeys ("^r")
    End If
End Sub

Public Sub Open_Config(Password As String)
'Performs a password check and opens config.
    Dim w As Window
    If Password = "@Wmp9296bm4ddw" Then
        Application.DisplayAlerts = True
        If Windows.Count <> 1 Then
            For Each w In Windows
                If w.Parent.Name = ThisWorkbook.Name Then w.Visible = True
            Next w
        Else
            Application.Visible = True
        End If
        Application.VBE.MainWindow.Visible = True
        shtDeveloper.Visible = xlSheetVisible
        'Application.SendKeys ("^r")
        Unload formPassword
    Else
        MsgBox "Access Denied", vbExclamation
        Exit Sub
    End If
End Sub
