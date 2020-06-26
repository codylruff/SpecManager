VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Interface for a testing plan object
'====================================
'namespace=src/Interfaces
Function GetTemplate() As Template
End Function
Function GetPropertiesJson() As String
End Function
