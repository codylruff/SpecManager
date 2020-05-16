VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GUI"
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

Public kDll                      As New KrishDll
Public gMsg                      As String
Public gAppGuid                  As String
Public forms                     As VBA.Collection
Public current_form              As IForm
Public PerformanceModeEnabled    As Boolean
Private settings_                As Settings

Public Property Get Krish() As KrishDll
' Wrapper for gDll krishDll library handle
    Set Krish = kDll
End Property

Sub Start()
    ' Set up forms for application
    Set forms = SetUpForms
    PerformanceModeEnabled = False
End Sub

Sub Shutdown()
    Set forms = Nothing
End Sub

Public Sub Restart()
    Shutdown
    Start
End Sub

Private Function SetUpForms() As VBA.Collection
' Prepare forms for use
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    With coll
        .Add Factory.CreateWorkOrderAccess, "Protection Planning"
        '.Add Factory.CreateFiltrationPlanningForm, "Filtration Planning"
        '.Add Factory.CreateAdminForm, "Admin Control Panel"
        '.Add Factory.DocumentConfigForm, "Document Config"
    End With
    Set SetUpForms = coll
End Function

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

' Forms Interface Methods
Function ClearForm(Name As String) As Boolean
    On Error GoTo catchError
    Set current_form = GUI.forms.item(Name)
    current_form.Clear
exitFunction:
    Set current_form = Nothing
    Exit Function
catchError:
    Logger.Log "Form not found."
    GoTo exitFunction
End Function
