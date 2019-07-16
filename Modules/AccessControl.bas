Attribute VB_Name = "AccessControl"
Option Explicit

Private user As Account

Function Account_Initialize() As Account
    If GetPrivledges(UCase(VBA.Environ("Username"))) <> DB_SELECT_SUCCESS Then
        Logger.Log "Creating new user " & VBA.Environ("Username"), UserLog
        Set Account_Initialize = AccessControl.AutoAddNewUser
    Else
        Logger.Log "Selected User : {" & user.ToString & "}", UserLog
        Set Account_Initialize = user
    End If
    
End Function

Function GetPrivledges(Name As String) As Long
    Dim record As DatabaseRecord
    Set record = DataAccess.GetUser(Name)
    ' obsoleted
    If record.Rows = 0 Then
        GetPrivledges = DB_SELECT_FAILURE
    Else
        Set user = New Account
        user.Name = record.records(1).Item("Name")
        user.PrivledgeLevel = record.records(1).Item("Privledge_Level")
        user.ProductLine = record.records(1).Item("Product_Line")
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
    Logger.Log DataAccess.PushNewUser(new_user), UserLog
    Set AutoAddNewUser = new_user
End Function

Public Sub ConfigControl()
'Initializes the password form for config access.
    Dim w As Window
    Dim origCalcMode As xlCalculation
    
    If Open_Config(App.gDll.CreateInputBox(Password, "Access Control", "Enter your password :")) Then
        
        ' Toggle Gui Functions to speed up
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        origCalcMode = Application.Calculation
        Application.Calculation = xlCalculationManual
        
        If Windows.Count <> 1 Then
            For Each w In Windows
                If w.Parent.Name = ThisWorkbook.Name Then w.Visible = True
            Next w
        Else
            Application.Visible = True
        End If
        ' Show all worksheets
        GuiCommands.ShowAllSheets SAATI_Data_Manager.ThisWorkbook
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = origCalcMode
    End If
End Sub

Private Function Open_Config(Password As String) As Boolean
'Performs a password check and opens config.
    Dim w As Window
    If Password = "@Wmp9296bm4ddw" Then
        Open_Config = True
    Else
        PromptHandler.AccessDenied
        Open_Config = False
    End If
End Function
