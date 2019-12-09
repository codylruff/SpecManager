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
Public kDll                 As New KrishDll
Public gMsg                 As String
Public gAppGuid             As String
Public version              As String
Public current_user         As Account
Public current_spec         As Specification
Public templates            As VBA.Collection
Public specs                As Object
Public printer              As DocumentPrinter
Public current_template     As SpecificationTemplate
Public TestingMode          As Boolean
Private pPerformanceMode    As Boolean
Private settings_           As Settings


Public Property Get GUI() As KrishDll
' Wrapper for gDll krishDll library handle
    Set GUI = kDll
End Property

' Methods
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
    Set printer = New DocumentPrinter
    Set current_template = New SpecificationTemplate
    TestingMode = False
    pPerformanceMode = False
End Sub

Public Sub RefreshObjects()
    Set current_spec = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
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
    Set printer = Nothing
    Set current_template = Nothing
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
    If enable And Not pPerformanceMode Then
        pPerformanceMode = True
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
        If pPerformanceMode Then
            pPerformanceMode = False
            Logger.Log "Performance Mode : OFF", RuntimeLog
            With Application
                .ScreenUpdating = settings_.ScreenUpdating
                .DisplayAlerts = settings_.DisplayAlerts
                .DisplayStatusBar = settings_.DisplayStatusBar
                .Calculation = settings_.Calculation
            End With
        End If
    End If

    
End Sub

