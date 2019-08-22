VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQueryable

Public PrivledgeLevel           As Long
Public ProductLine              As String
Private Name_                   As String
Public Settings                 As Object
Public SecretSHA1               As String
Public FlaggedForPasswordChange As Boolean

Public Property Let Name(value As String)
    Name_ = UCase(value)
End Property

Public Property Get Name() As String
    Name = Name_
End Property

Public Function ToString() As String
  ToString = "User Name : " & Me.Name & _
             ", Privledge Level : " & Me.PrivledgeLevel & _
             ", Product Line : " & Me.ProductLine
End Function

Public Function GetSecret() As String
' User secret as SHA1 has
    GetSecret = SecretSHA1
End Function

Public Sub LoadUserJson()
' Advanced example: Read .json file and load into a dictionary
    Set Settings = JsonVBA.GetJsonObject(ThisWorkbook.path & "\config\user.json")
    If Me.Settings Is Nothing Then CreateSettingsJson
End Sub

Public Sub ChangeSetting(setting As String, new_value As Variant)
    Settings.item(setting) = new_value
    SaveUserJson
End Sub

Public Sub SaveUserJson()
' Saves user configuration settings to user.json saved in the Spec Manager installation folder
    JsonVBA.WriteJsonObject ThisWorkbook.path & "\config\user.json", Me.Settings
End Sub

Private Sub CreateSettingsJson()
    If Settings Is Nothing Then
        Set Settings = CreateObject("Scripting.Dictionary")
    End If
    With Settings
        .Add "name", Environ("Username")
        .Add "default_printer", vbNullString
        .Add "default_log_level", vbNullString
        .Add "privledge_level", vbNullString
        .Add "product_line", vbNullString
        .Add "repo_path", vbNullString
    End With
End Sub

' Constructor
Private Sub Class_Initialize()
    Set Settings = CreateObject("Scripting.Dictionary")
End Sub

' Deconstructor
Private Sub Class_Terminate()
    Set Settings = Nothing
End Sub

' IQueryable Methods
Function IQueryable_GetValueLabels() As String
    IQueryable_GetValueLabels = "Name, Privledge_Level, Product_Line"
End Function

Function IQueryable_GetValues() As String
    IQueryable_GetValues = "'" & Me.Name & "', " & _
                           "'" & Me.PrivledgeLevel & "', " & _
                           "'" & Me.ProductLine & "'"
End Function
