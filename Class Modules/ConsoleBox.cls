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
     CurrentText = Me.FormId.txtConsole.Text
End Property
'----------
' METHODS
'----------
Public Sub PrintObjectToSheet(spec As Specification, sht As Worksheet)
' Print object to specified worksheet starting in cell "A1"
    Dim Key As Variant
    Dim i As Integer
    sht.Cells.ClearContents
    With sht
        .Range("A" & 1).Value = spec.SpecType & " : "
        .Range("B" & 1).Value = spec.MaterialId
        .Range("A" & 1).Font.Bold = True
        .Range("A" & 1).Font.Size = 20
        .Range("B" & 1).Font.Bold = True
        .Range("B" & 1).Font.Size = 20
        i = 3
        For Each Key In spec.Properties
            If Not spec.Properties(Key) = vbNullString Then
                .Range("A" & i).Value = CStr(Key) & " : "
                .Range("B" & i).Value = CStr(spec.Properties(CStr(Key)))
                .Range("A" & i).Font.Size = 17
                .Range("B" & i).Font.Size = 17
                .Range("B" & i).HorizontalAlignment = xlHAlignCenter
                If spec.Properties(Key) = Chr(252) Or spec.Properties(Key) = Chr(251) Then
                    .Range("B" & i).Font.Name = "Wingdings"
                End If
            Else
                .Range("A" & i).Value = CStr(Key)
                .Range("B" & i).Value = "-"
                .Range("A" & i).Font.Size = 17
                .Range("B" & i).Font.Size = 17
            End If
            i = i + 1
        Next Key
        .columns("A:G").AutoFit
        i = i + 2
    End With
    Dim xPic As Object
    Dim PicPath As String
    Dim img As Picture
    If spec.SpecType = "Ballistic Testing Requirements" Then
        
        For Each xPic In sht.Pictures
            xPic.Delete
        Next
        
        PicPath = "S:\Data Manager\Specifications\ballistic-package\15x15-corner.jpg"
        
        sht.Activate
        ActiveSheet.Range("A" & i).Activate
        Set img = ActiveSheet.Pictures.Insert(PicPath)
    ElseIf spec.SpecType = "Testing Requirements" Then
        For Each xPic In sht.Pictures
            xPic.Delete
        Next
        PicPath = "S:\Data Manager\Specifications\weave-patterns\plain.jpg"
        sht.Activate
        ActiveSheet.Range("A" & i).Activate
        Set img = ActiveSheet.Pictures.Insert(PicPath)
    End If
End Sub

Public Sub PrintObject(obj As Object)
' Print object to console
    Dim Key As Variant
    ' Clear the console
    FormId.txtConsole.Text = vbNullString
    'PrintLine obj.SpecType, obj.MaterialId
    'PrintLine "Revision", obj.Revision
    
    For Each Key In obj.Properties
        If Not obj.Properties(Key) = vbNullString Then
            PrintLine CStr(Key), FormatAsciiToYesNo(obj.Properties(CStr(Key)))
        Else
            PrintLine CStr(Key), "-"
        End If
    Next Key
End Sub

Private Function FormatAsciiToYesNo(char As String) As String

    If char = Chr(252) Then
        char = "Required"
    ElseIf char = Chr(251) Then
        char = "-"
    End If
    FormatAsciiToYesNo = char
End Function

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



