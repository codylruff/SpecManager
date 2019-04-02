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
