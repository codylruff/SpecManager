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

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Type Settings
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    DisplayStatusBar As Boolean
    Calculation As Long
End Type

' Events
Public Event BeforeRefresh()
Public Event AfterRefresh()
Public Event BeforeStart()
Public Event AfterStart()
Public Event BeforeShutdown()

' Fields
Public kDll                      As New KrishDll
Public gMsg                      As String
Public gAppGuid                  As String
Public version                   As String
Public current_user              As Account
Public current_spec              As Specification
Public forms                     As VBA.Collection
Public templates                 As VBA.Collection
Public specs                     As Object
Public printer                   As DocumentPrinter
Public current_template          As SpecificationTemplate
Public current_machine_id        As String
Public TestingMode               As Boolean
Public PerformanceModeEnabled    As Boolean
Private settings_                As Settings
Private DocumentsByUID_          As Object

Public Property Get DocumentsByUID() As Object
    
    'Set DocumentsByUID_ = SortDocumentsByUID
    'Set DocumentsByUID = DocumentsByUID_
    Set DocumentsByUID = specs
End Property

Public Property Get GUI() As KrishDll
' Wrapper for gDll krishDll library handle
    Set GUI = kDll
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
    Set current_spec = New Specification
    Set templates = SpecManager.GetAllTemplates
    Set specs = CreateObject("Scripting.Dictionary")
    Set DocumentsByUID_ = CreateObject("Scripting.Dictionary")
    Set forms = SetUpForms()
    Set printer = New DocumentPrinter
    Set current_template = New SpecificationTemplate
    TestingMode = False
    PerformanceModeEnabled = False
End Sub

Private Function SetUpForms() As VBA.Collection
' Prepare forms for use
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    With coll
        .Add Factory.CreateProtectionPlanningForm, "Protection Planning"
        .Add Factory.CreateFiltrationPlanningForm, "Filtration Planning"
        .Add Factory.CreateAdminForm, "Admin Control Panel"
        .Add Factory.SpecificationConfigForm, "Specification Config"
    End With
    Set SetUpForms = coll
End Function

Public Sub RefreshObjects()
    Set current_spec = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
    Set DocumentsByUID_ = CreateObject("Scripting.Dictionary")
    Set printer = New DocumentPrinter
    Set current_template = New SpecificationTemplate
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
    Set current_spec = Nothing
    Set specs = Nothing
    Set DocumentsByUID_ = Nothing
    Set printer = Nothing
    Set current_template = Nothing
    Set forms = Nothing
    Logger.SaveAllLogs
End Sub

Public Sub Restart()
    Shutdown
    Start
End Sub

Public Sub SaveSettings()
    With Application
        settings_.ScreenUpdating = .ScreenUpdating
        settings_.DisplayAlerts = .DisplayAlerts
        settings_.DisplayStatusBar = .DisplayStatusBar
        settings_.Calculation = .Calculation
    End With
End Sub

Public Sub PerformanceMode(enable As Boolean, Optional DisableEvents As Boolean = False)
' This Sub turns off un-necessary gui functions to increase performance
    PerformanceModeEnabled = enable

    If enable Then
        Logger.Log "Performance Mode : ON", RuntimeLog
        ' Save current excel-gui settings
        SaveSettings
        ' Enabled performance mode
        With Application
            If .ScreenUpdating Then .ScreenUpdating = False
            If .DisplayAlerts Then .DisplayAlerts = False
            If .DisplayStatusBar Then .DisplayStatusBar = False
            'If .EnableEvents And DisableEvents Then .EnableEvents = False
            If .Calculation = xlCalculationAutomatic Then .Calculation = xlCalculationManual
        End With
    Else
        Logger.Log "Performance Mode : OFF", RuntimeLog
        With Application
            .ScreenUpdating = settings_.ScreenUpdating
            .DisplayAlerts = settings_.DisplayAlerts
            .DisplayStatusBar = settings_.DisplayStatusBar
            .Calculation = settings_.Calculation
        End With
    End If

    
End Sub

