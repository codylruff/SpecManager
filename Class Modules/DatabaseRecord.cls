VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_header    As Variant
Private m_data      As Variant
Private rows_       As Long
Private columns_    As Long
Public Fields       As Object
Private m_records   As VBA.Collection

Private Property Set records(ByRef obj As VBA.Collection)
    Set m_records = obj
End Property
Public Property Get records() As VBA.Collection
    If m_records Is Nothing Then
       InitializeRecords
    End If
    Set records = m_records
End Property
Public Property Get Header() As Variant: Header = m_header: End Property
Public Property Let Header(ByVal Value As Variant): m_header = Value: End Property
Public Property Get Data() As Variant: Data = m_data: End Property
Public Property Let Data(ByVal Value As Variant): m_data = Value: End Property
Public Property Let rows(Value As Long): rows_ = Value: End Property
Public Property Get rows() As Long: rows = rows_: End Property
Public Property Let columns(Value As Long): columns_ = Value: End Property
Public Property Get columns() As Long: columns = columns_: End Property

Private Function InitializeRecords()
' Returns a dictionary representation of this object
    Dim dict As Object
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    Dim i, j As Longi
    For j = 1 To rows
        Set dict = Factory.CreateDictionary()
        Logger.Log "Checking Row " & j & " . . . "
        For i = 1 To columns
        Logger.Log "Column " & i & " : " & Header(1, i) & ", Row " & j & " : " & Data(j, i)
            dict.Add Key:=Header(1, i), Item:=Data(j, i)
        Next i
        coll.Add dict
    Next j
    Set Fields = dict
    Set m_records = coll
End Function
