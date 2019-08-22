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
Public gDll                 As New KrishDll
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
End Sub

Public Sub DeInitializeTestSuiteCredentials()
    App.TestingMode = False
    Me.RefreshObjects
    ActionLog.enable
    Set current_user = AccessControl.Account_Initialize
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

' -----------------------
' VBA-TOOLS Methods
' -----------------------
Public Function testJson()
'---------------------------------------------------------------------------------------
' Procedure : testJson
' Author    : KRISH
' Date      : 21/05/2018
' Purpose   : returns true or false whether an array has item or not
' Returns   :
'---------------------------------------------------------------------------------------
'
    Dim T As New TDialog
    T.title = "test"
    T.message = "this is a ' message"
    T.dialogType = vbInformation
    T.AddButton "button1", 1
    T.AddButton "button2", 2, True
    
    Debug.Print T.ToJSON
    
    
End Function

'' Samples
Public Function testOpenFile()
    Dim Filters() As String
    FnArrayAddItem Filters, "all files|*.*;"
    
    Dim Files As String
    
    Files = gDll.OpenFileA("test", Filters, False)
    
    On Error Resume Next
    Debug.Print Files
    'Debug.Print UBound(Files), Files(0)
    
End Function

Public Function testJsonObject()
    Dim Actor As Variant
    Dim ActorAddress As Variant
    
    Set Actor = gDll.JsonParse("{ 'Name': 'John Wick', 'Address': { 'City': 'New York', 'State': 'NY' }, 'Age': 53 }")
    
    Set ActorAddress = gDll.JsonGetObject(Actor, "Address")
    Debug.Print gDll.JsonGetValue(Actor, "Name") & " is from " & gDll.JsonGetValue(ActorAddress, "City") & " and he is " & gDll.JsonGetValue(Actor, "Age") & " years old."
    Debug.Print gDll.JsonGetValue(Actor, "Address")
    
End Function

Public Function TestJsonSerializeObject()
    Dim Holder As Object
    Set Holder = CreateObject("Scripting.Dictionary")
    Holder.Add "Key1", "Value1"
    Holder.Add "Key2", "Value2"
    
    Debug.Print gDll.DLL.JsonSerialize(Holder)
    
End Function

Public Function OpenSettings(Parameter As String)
    gDll.Toast "Hey you called OpenSettings() function with a parameter " & Parameter, vbInformation
End Function

Public Function ExecuteMe(Parameter1 As String, parameter2 As String)
    gDll.Toast gDll.sFormat("Hey you called OpenSettings() function with a parameter {0} and {1}", Parameter1, parameter2), vbInformation
End Function

Public Function PicturePathFromUrl(url As String)
    PicturePathFromUrl = gDll.PicturePathFromUrl(url, False, "", True)
End Function

