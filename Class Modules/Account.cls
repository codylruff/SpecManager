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

Public PrivledgeLevel   As Long
Public ProductLine      As String
Public Protection       As Long
Public Filtration       As Long
Public Chemical         As Long
Private Name_           As String
Public LogLevel         As Long
Public DefaultPrinter   As String

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

Public Sub LoadUserJson()
' Advanced example: Read .json file and load into sheet (Windows-only)
' (add reference to Microsoft Scripting Runtime)

  Dim FSO As New Object
  Dim JsonTS As Object
  Dim JsonText As String
  Dim Parsed As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  ' Read .json file
  Set JsonTS = fso.OpenTextFile("user.json", ForReading)
  JsonText = JsonTS.ReadAll
  JsonTS.Close
  ' Parse json to Dictionary
  Set Parsed = JsonConverter.ParseJson(JsonText)
  LogLevel = Parsed.Item("log_level")
  DefaultPrinter = Parsed.Item("default_printer")
End Sub

Public Sub SaveUserJson()
' Saves user configuration settings to user.json saved in the Spec Manager installation folder


End Sub