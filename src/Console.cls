VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
Private m_Text As String
Private m_Sheet As Worksheet
Public Property Set Sheet(obj As Worksheet): Set m_Sheet = obj: End Property
Public Property Get Sheet() As Worksheet: Set Sheet = m_Sheet: End Property
Public Property Let text(val As String): Me.Address.value = val: End Property
Public Property Get text() As String: text = Me.Address.value: End Property
Public Property Get Address() As Range
  On Error GoTo Catch
  Set Address = Me.Sheet.Range("console")
  GoTo Finally
Catch:
  Logger.Error "Could not find a console in the referenced worksheet."
  Set Address = Nothing
Finally:
  'PASS
End Property
Public Sub Clear()
  Me.text = nullstr
End Sub
Public Sub WriteLine(msg As String)
' Console.WriteLine
  Me.text = Me.text & msg & vbNewLine
   
End Sub
