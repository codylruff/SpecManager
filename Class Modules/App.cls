VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public current_user As Account
Public standard As Specification
Public current_spec As Specification
Public specs As Object
Public console As ConsoleBox
Public current_template As SpecTemplate
Public warp_order As Warp

Private Sub Class_Initialize()
    Set current_user = AccessControl.Account_Initialize
    current_user.LoadUserJson
    Set current_spec = New Specification
    Set standard = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
    Set console = New ConsoleBox
    Set current_template = New SpecTemplate
    Set warp_order = New Warp
End Sub

Public Sub ResetInteractiveObject()
    Set current_spec = New Specification
    Set standard = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
    Set console = New ConsoleBox
    Set current_template = New SpecTemplate
    Set warp_order = New Warp
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

Private Sub Class_Terminate()
    Set current_user = Nothing
    Set current_spec = Nothing
    Set standard = Nothing
    Set specs = Nothing
    Set console = Nothing
    Set current_template = Nothing
    Set warp_order = Nothing
End Sub


