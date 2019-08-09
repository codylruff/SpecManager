VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Consider changing the name of this class to data frame (like pandas)
Private list As ListObject
Private CustomSorters As VBA.Collection

Property Set ListObject(ByVal obj As ListObject)
    Set list = obj
End Property

Property Get ListObject() As ListObject
    Set ListObject = list
End Property

Public Property Get Columns() As ListColumns
    Set ListColumns = list.ListColumns
End Property

Property Get Rows() As ListRows
    Set ListRows = list.ListRows
End Property

Property Get Name() As String
    Name = list.Name
End Property

Property Get DisplayName() As String
    DisplayName = list.DisplayName
End Property

Property Get Range() As Range
    Set Range = list.Range
End Property

Property Get Data() As Range
    Set Data = list.DataBodyRange
End Property

Property Get HeaderRowRange() As Range
    Set HeaderRowRange = list.HeaderRowRange
End Property

Property Get InsertRowRange() As Range
    Set InsertRowRange = list.InsertRowRange
End Property

Property Get TotalsRowRange() As Range
    Set TotalsRowRange = list.TotalsRowRange
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = list.Parent
End Property

Property Get Workbook() As Workbook
    Set Workbook = list.Parent.Parent
End Property

Property Get TableName() As String
    TableName = "[" & Me.Worksheet.Name & "$" & Me.Range.Address(False, False) & "]"
End Property

Public Sub AddCustomSorter(sort_dict As Object, sort_name As String)
' Create a custom table sorter given a dictionary of labels with sort index (integer)
    CustomSorters.Add sort_dict, Key:=sort_name
End Sub

Public Sub CustomSort(sort_name As String, column_name As String)
' Retrieve custom sorter by name and apply it to this table.
' NEEDS TESTING
    Dim sorter As Object
    Dim r As Long
    Dim C As Long

    If CustomSorters.Contains(sort_name) Then
        Set sorter = CustomSorters.item(sort_name)
    Else
        Exit Sub
    End If
    
    C = Me.Columns(column_name).index
    For r = 0 To Me.Rows.Count - 1
        On Error Resume Next
        Me.Data(r, C).value = sorter(Me.Data(r, C).value)
    Next r

    On Error GoTo 0
    list.Sort = Me.Columns(column_name)

End Sub

Private Sub Class_Initialze()
    Set CustomSorters = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set CustomSorters = Nothing
End Sub
