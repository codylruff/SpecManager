VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public version As String
Public current_user As Account
Public standard As Specification
Public current_spec As Specification
Public templates As VBA.Collection
Public specs As Object
Public console As ConsoleBox
Public current_template As SpecificationTemplate

Sub Start()
    ' Log in with environ(username)
    Set current_user = AccessControl.Account_Initialize
    ' Load user settings
    current_user.LoadUserJson
    ' Create global object instances for specs and templates
    Set current_spec = New Specification
    Set standard = New Specification
    Set templates = New VBA.Collection
    Set specs = New Dictionary
    Set console = New ConsoleBox
    Set current_template = New SpecificationTemplate
End Sub

Public Sub ResetInteractiveObject()
    Set current_spec = New Specification
    Set standard = New Specification
    Set specs = New Dictionary
    Set console = New ConsoleBox
    Set current_template = New SpecificationTemplate
End Sub

Public Sub InitializeTestSuite()
    Set current_user = New Account
    current_user.Name = "TestSuite"
    current_user.PrivledgeLevel = 25
    current_user.ProductLine = "Admin"
End Sub

Public Sub DeinitializeTestSuite()
    Me.ResetInteractiveObject
    Set current_user = AccessControl.Account_Initialize
End Sub

Public Sub Shutdown()
    If Not current_user Is Nothing Then current_user.SaveUserJson
    Set current_user = Nothing
    Set current_spec = Nothing
    Set standard = Nothing
    Set specs = Nothing
    Set console = Nothing
    Set current_template = Nothing
End Sub


