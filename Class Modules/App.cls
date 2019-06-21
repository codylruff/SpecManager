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
' Events
Public Event BeforeRefresh
Public Event AfterRefresh
Public Event BeforeStart
Public Event AfterStart
Public Event BeforeShutdown
' Fields
Public version As String
Public current_user As Account
Public current_spec As Specification
Public templates As VBA.Collection
Public specs As Object
Public console As ConsoleBox
Public current_template As SpecificationTemplate
' Methods
Sub Start()
    ' Log in with environ(username)
    Set current_user = AccessControl.Account_Initialize
    ' Load user settings
    current_user.LoadUserJson
    ' Create global object instances for specs and templates
    Set current_spec = New Specification
    Set templates = SpecManager.GetAllTemplates
    Set specs = New Dictionary
    Set console = New ConsoleBox
    Set current_template = New SpecificationTemplate
End Sub

Public Sub RefreshObjects()
    Set current_spec = New Specification
    Set templates = SpecManager.GetAllTemplates
    Set specs = New Dictionary
    Set console = New ConsoleBox
    Set current_template = New SpecificationTemplate
End Sub

Public Sub InitializeTestSuiteCredentials()
    Set current_user = New Account
    current_user.Name = "TestSuite"
    current_user.PrivledgeLevel = 25
    current_user.ProductLine = "Admin"
End Sub

Public Sub DeInitializeTestSuiteCredentials()
    Me.RefreshObjects
    Set current_user = AccessControl.Account_Initialize
End Sub

Public Sub Shutdown()
    If Not current_user Is Nothing Then current_user.SaveUserJson
    Set current_user = Nothing
    Set current_spec = Nothing
    Set specs = Nothing
    Set console = Nothing
    Set current_template = Nothing
End Sub


