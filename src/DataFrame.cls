VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_header    As Variant
Private m_data      As Variant
Private rows_       As Long
Private columns_    As Long
Public m_fields     As Object
Private m_records   As VBA.Collection

Private Property Set Fields(ByRef obj As Object)
    Set m_fields = obj
End Property
Public Property Get Fields() As Object
    On Error Resume Next
    If m_fields Is Nothing Then
       InitializeRecords
    End If
    Set Fields = m_fields
End Property

Private Property Set records(ByRef obj As VBA.Collection)
    Set m_records = obj
End Property
Public Property Get records() As VBA.Collection
    If m_records Is Nothing Then
       InitializeRecords
    End If
    Set records = m_records
End Property

Public Property Get header() As Variant: header = m_header: End Property
Public Property Let header(ByVal value As Variant): m_header = value: End Property
Public Property Get Data() As Variant: Data = m_data: End Property
Public Property Let Data(ByVal value As Variant): m_data = value: End Property
Public Property Let Rows(value As Long): rows_ = value: End Property
Public Property Get Rows() As Long: Rows = rows_: End Property
Public Property Let Columns(value As Long): columns_ = value: End Property
Public Property Get Columns() As Long: Columns = columns_: End Property

Public Function At(row_id As Long, column_id As Long) As Variant
    At = m_data(row_id, column_id)
End Function

Public Function ToString() As String
    Dim print_string As String
    Dim prop As Variant
    Dim i As Long
    Dim k As Long
    print_string = nullstr
    For Each prop In header
        print_string = Utils.printf(print_string & " {0} {1}", " |", CStr(prop))
    Next prop
    print_string = print_string & vbNewLine
    For i = 1 To UBound(Data, 1)
        For k = 1 To UBound(Data, 2)
            print_string = Utils.printf(print_string & " {0} {1}", " |", CStr(At(i, k)))
        Next k
    Next i
    ToString = print_string
End Function

Private Function InitializeRecords()
' Returns a dictionary representation of this object
    Dim dict As Object
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    Dim i, j As Long
    On Error Resume Next
    For j = 1 To Rows
        Set dict = Factory.CreateDictionary()
        Logger.Log "Checking Row " & j & " . . . ", DebugLog
        For i = 1 To Columns
            Logger.Log "Column " & i & " : " & header(1, i) & ", Row " & j & " : " & Data(j, i), DebugLog
            If dict.Exists(header(1, i)) Then
                If dict(header(1, i)).value = nullstr Then
                    dict.Add Key:=header(1, i), item:=Data(j, i)
                End If
            Else
                dict.Add Key:=header(1, i), item:=Data(j, i)
            End If
        Next i
        coll.Add dict
    Next j
    Set m_fields = dict
    Set m_records = coll
End Function

Private Sub Class_Initialize()
    Set m_fields = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_fields = Nothing
End Sub
