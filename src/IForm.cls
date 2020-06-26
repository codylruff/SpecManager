VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
Public Property Get Name() As String
End Property

Public Property Set Console(val As Console)
End Property

Public Property Get Console() As Console
End Property

Public Function Clear()
End Function

Public Property Get Sheet() As Worksheet
End Property

Public Property Get HelpText() As String
End Property

Public Property Get Fields() As Variant
' Includes all editiable named ranges or fields.
End Property

Public Function CmdByName(cmd As String) As Boolean
' Allows form handler to look up commands without needing to know them before runtime.
End Function

Public Function ReformatSheet() As Boolean
' This allows the form it self to be properly formatted in case a user messes it up.
End Function

Public Function FieldValue(field As String) As Variant
' Reads the value of a field in the form
End Function

Public Function ChangeFieldValue(field As String, val As Variant)
' Changes the value of a field in the form
End Function
