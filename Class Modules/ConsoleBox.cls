VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Classes")
'====================================
'CLASS DESCRIPTION:
' Takes a user form and text box name
' and uses it as a console output.
'====================================
'Factory Module Constructor function:
'Function CreateConsoleBox(frm As UserForm) As ConsoleBox
'----------------
'CLASS VARIABLES
'----------------
Private FormId_         As MSForms.UserForm
Private OutputHistory_  As VBA.Collection
'------------
' AUTO-PROPERTIES
'------------
Public Property Set OutputHistory(Value As VBA.Collection): Set OutputHistory_ = Value: End Property
Public Property Get OutputHistory() As VBA.Collection: Set OutputHistory = OutputHistory_: End Property
Public Property Set FormId(Value As MSForms.UserForm): Set FormId_ = Value: End Property
Public Property Get FormId() As MSForms.UserForm: Set FormId = FormId_: End Property
Public Property Get CurrentText() As String
     CurrentText = FormId.txtConsole.Text
End Property
'----------
' METHODS
'----------
Public Sub PrintObjectToSheet(spec As Specification, sht As Worksheet)
' Print object to specified worksheet starting in cell "A1"
    Dim Key As Variant
    Dim i As Integer
    sht.Cells.ClearContents
    With sht.Range("A" & 1)
        .Value = "Testing Requirements : " & spec.MaterialId
        .Font.Bold = True
        .Font.Size = 13
    End With
    With sht
        i = 3
        For Each Key In spec.Properties
            If Not spec.Properties(Key) = vbNullString Then
                .Range("A" & i).Value = CStr(Key) & " : "
                .Range("B" & i).Value = spec.Properties(CStr(Key))
                .Range("A" & i).Font.Size = 12
                .Range("B" & i).Font.Size = 12
            Else
                .Range("A" & i).Value = CStr(Key)
                .Range("B" & i).Value = "-"
                .Range("A" & i).Font.Size = 12
                .Range("B" & i).Font.Size = 12
            End If
            i = i + 1
        Next Key
        .columns("A:G").AutoFit
        i = i + 2
    End With
    Dim PicPath As String
    PicPath = "S:\Data Manager\Specifications\weave-patterns\101.jpg"
   Dim img As Picture
   Sheets("SpecificationForm").Activate
   ActiveSheet.Range("A" & i).Activate
   Set img = ActiveSheet.Pictures.Insert(PicPath)
End Sub

Public Sub PrintObject(obj As Object)
' Print object to console
    Dim Key As Variant
    ' Clear the console
    FormId.txtConsole.Text = vbNullString
    For Each Key In obj.Properties
        If Not obj.Properties(Key) = vbNullString Then
            PrintLine CStr(Key), obj.Properties(CStr(Key))
        Else
            PrintLine CStr(Key), "-"
        End If
    Next Key
End Sub

Public Sub PrintLine(ParamArray var() As Variant)
    Dim i As Integer
    Dim s As String
    Const Padding = 25
    s = vbNullString
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    OutputHistory.Add s
    Me.FormId.txtConsole.Text = FormId.txtConsole.Text & s & vbNewLine
End Sub

Private Sub Class_Initialize()
' Constructor
    Set OutputHistory = New VBA.Collection
End Sub



