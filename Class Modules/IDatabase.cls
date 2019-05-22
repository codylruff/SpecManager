VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Interfaces")
'====================================
'CLASS DESCRIPTION:
' Interface for databases
'====================================
' public properties
Public Property Get Header() As Variant: End Property ' Array representing column headers
Public Property Get Data() As Variant: End Property ' Array representing data from table
Public Property Get NumRows() As Long: End Property
Public Property Get NumColumns() As Long: End Property

Public Sub openDb(sqliPath As String)
' opens a database
End Sub

Public Sub Execute(ssql As String)
' Execute sql statement
End Sub

Public Sub selectQry(ssql As String)
' Execute select statement
End Sub
