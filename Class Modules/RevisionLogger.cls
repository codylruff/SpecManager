VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RevisionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRevisionLog As Object
Private pBuffer As VBA.Collection

Private WithEvents pSpecification As Specification
Attribute pSpecification.VB_VarHelpID = -1

Public Sub RecordAction()
    ' PASS
End Sub

Public Sub ListenTo(spec As Specification)
    Set pSpecification = spec
End Sub

Private Sub pSpecification_PropertyChanged(property As Variant, old_val As Variant, new_val As Variant)
      pBuffer.Add "Changed : " & CStr(property) & _
                  ", " & CStr(old_val) & " --> " & CStr(new_val)
End Sub

Private Sub Class_Initialize()
    Set pRevisionLog = CreateObject("Scripting.Dictionary")
End Sub
