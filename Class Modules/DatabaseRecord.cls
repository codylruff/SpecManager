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
Public records      As Collection

Public Property Get header() As Variant: header = m_header: End Property
Public Property Let header(ByVal value As Variant): m_header = value: End Property
Public Property Get data() As Variant: data = m_data: End Property
Public Property Let data(ByVal value As Variant): m_data = value: End Property
Public Property Let rows(value As Long): rows_ = value: End Property
Public Property Get rows() As Long: rows = rows_: End Property
Public Property Let columns(value As Long): columns_ = value: End Property
Public Property Get columns() As Long: columns = columns_: End Property

Public Function SetDictionary()
' Returns a dictionary representation of this object
    Dim dict As Object
    Dim coll As Collection
    Set coll = New Collection
    Dim i, j As Long
    For j = 1 To rows
        Set dict = Factory.CreateDictionary()
        For i = 1 To columns
        Logger.Log "Column " & i & " : " & header(1, i) & ", Row " & j & " : " & data(j, i)
            dict.Add key:=header(1, i), Item:=data(j, i)
        Next i
        coll.Add dict
    Next j
    Set Fields = dict
    Set records = coll
End Function
