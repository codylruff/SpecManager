VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
Implements IForm
' DEFAULT PROPERTY VALUES
Private Const FORM_NAME As String = "View Document Form"
Private Const HELP_TEXT As String = "TODO" ' Thinking of loading this from a text file??
Private m_IForm_Console As Console
Private m_IForm_Fields() As Variant
'--------------
' I_PROPERTIES
'--------------
Public Property Set IForm_Console(value As Console): Set m_IForm_Console = value: End Property
Public Property Get IForm_Console() As Console: Set IForm_Console = m_IForm_Console: End Property

Public Property Get IForm_Sheet() As Worksheet
    Set IForm_Sheet = shtView
End Property

Public Property Get IForm_Name() As String
    IForm_Name = FORM_NAME
End Property

Public Property Get IForm_HelpText() As String
    IForm_HelpText = HELP_TEXT
End Property

Public Property Get IForm_Fields() As Variant
    IForm_Fields = m_IForm_Fields
End Property

'----------
' I_METHODS
'----------
Public Function IForm_Clear()
    ' Console must be cleared through object
    IForm_Console.Clear
    ' Fields can be cleared manually.
    For i = 0 To UBound(IForm_Fields) - 1
        IForm_Sheet.Range(IForm_Fields(i)).value = nullstr
    Next i
End Function

Public Function IForm_CmdByName(cmd As String) As Boolean
    Dim ret_val As Boolean
    Select Case cmd
        Case Default
            ret_val = False
    End Select
    IForm_CmdByName = ret_val
End Function

Public Function IForm_ReformatSheet() As Boolean
    ' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
    IForm_ReformatSheet = False
End Function

Public Function IForm_FieldValue(field As String) As Variant
' Reads the value of a field in the form
' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
    IForm_FieldValue = False
End Function

Public Function IForm_ChangeFieldValue(field As String, val As Variant)
' Changes the value of a field in the form
' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
End Function
'----------
' COMMANDS
'----------
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    Set IForm_Console = Factory.CreateConsole(Me.IForm_Sheet)
    ' Initialize the fields array to save on Array() calls.
    m_IForm_Fields = Array("")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    'REVIEW Is there anything that needs to be destroyed here?
End Sub
