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
Public Event BeforeRefresh()
Public Event AfterRefresh()
Public Event BeforeStart()
Public Event AfterStart()
Public Event BeforeShutdown()

' Fields
Public version                   As String
Public current_user              As Account
Public current_doc              As Document
Public templates                 As VBA.Collection
Public specs                     As Object
Public printer                   As DocumentPrinter
Public current_template          As Template
Public current_machine_id        As String
Public TestingMode               As Boolean
Private DocumentsByUID_          As Object

Public Property Get DocumentsByUID() As Object
    
    'Set DocumentsByUID_ = SortDocumentsByUID
    'Set DocumentsByUID = DocumentsByUID_
    Set DocumentsByUID = specs
End Property

' Methods

Private Function SortDocumentsByUID() As Object
' Sort specs by UID and return it as a new dictionary
    Dim docs_by_uid As Object
    Dim doc As Variant

    Set docs_by_uid = Factory.CreateDictionary

    ' Create a dictionary based on UID
    For Each doc In specs
        docs_by_uid.Add Key:=specs(doc).UID, item:=specs(doc)
    Next doc
    
    Set SortDocumentsByUID = docs_by_uid
End Function

Sub Start()
    ' Log in with environ(username)
    Set current_user = AccessControl.Account_Initialize
    If current_user.FlaggedForPasswordChange Then AccessControl.ChangeSecret current_user
    ' Load user settings
    current_user.LoadUserJson
    ' Create global object instances for specs and templates
    Set current_doc = New Document
    Set templates = SpecManager.GetAllTemplates
    Set specs = CreateObject("Scripting.Dictionary")
    Set DocumentsByUID_ = CreateObject("Scripting.Dictionary")
    Set printer = New DocumentPrinter
    Set current_template = New Template
    TestingMode = False
End Sub

Public Sub RefreshObjects()
    Set current_doc = New Document
    Set specs = CreateObject("Scripting.Dictionary")
    Set DocumentsByUID_ = CreateObject("Scripting.Dictionary")
    Set printer = New DocumentPrinter
    Set current_template = New Template
End Sub

Public Sub InitializeTestSuiteCredentials()
    Set current_user = New Account
    current_user.Name = "TestSuite"
    current_user.PrivledgeLevel = 25
    current_user.ProductLine = "Admin"
    App.TestingMode = True
    ActionLog.Disable
    Logger.EnableSaving
End Sub

Public Sub DeInitializeTestSuiteCredentials()
    App.TestingMode = False
    Me.RefreshObjects
    ActionLog.enable
    Set current_user = AccessControl.Account_Initialize
    Logger.DisableSaving
End Sub

Public Sub Shutdown()
    If Not current_user Is Nothing Then
        current_user.SaveUserJson
    End If
    Set current_user = Nothing
    Set current_doc = Nothing
    Set specs = Nothing
    Set DocumentsByUID_ = Nothing
    Set printer = Nothing
    Set current_template = Nothing
    
    Logger.SaveAllLogs
End Sub

