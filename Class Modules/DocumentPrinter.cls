VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentPrinter"
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
'Function CreateDocumentPrinter(frm As UserForm) As DocumentPrinter
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
     CurrentText = Me.FormId.txtConsole.text
End Property
Public Property Let CurrentText(Value As String)
    CurrentText = Value
End Property
'----------
' METHODS
'----------
Public Sub ClearConsole()
' Clears the console of any text
    Me.CurrentText = vbNullString
    Logger.Log "Console Cleared"
End Sub

Private Sub InsertHeader(ws As Worksheet, spec As Specification, Optional production_order As String = vbNullString)
' Header Content and formatting for Spec Docs
    ' Insert Header
    With ws.PageSetup
        .LeftHeader = spec.SpecType
        .CenterHeader = spec.MaterialId
        .RightHeader = spec.Revision & vbNewLine & _
                       CStr(Now()) & vbNewLine & _
                       production_order
    End With
End Sub

Private Sub InsertFooter(ws As Worksheet, Optional disclaimer As String = vbNullString, _
            Optional leftFooterText = vbNullString, Optional rightFooterText = vbNullString)
' Footer Content and formatting for Spec Docs

    If IsMissing(disclaimer) Then
        disclaimer = "This document is the property of SAATI Americas - " & _
                     "Printed and/or Unsigned Copies Are Uncontrolled."
    End If
    
    ' Insert Footer
    With ws.PageSetup
        .LeftFooter = leftHeaderText
        .CenterFooter = disclaimer
        .RightFooter = rightHeaderText
    End With

End Sub

Private Function GetWeavePatternImagePath(material_id As String) As String
' Given a material_id retrieve the file path for the weave pattern image
    Dim weave_id As String

    Select Case Mid(material_id, 13, 1)
        Case "9"
            weave_id = "6HS"
        Case "A"
            weave_id = "Twill_2x2"
        Case "B"
            weave_id = "Twill_4x4"
        Case "C"
            weave_id = "Basket_2x2"
        Case "D"
            weave_id = "Basket_3x3"
        Case "E"
            weave_id = "Basket_4x4"
        Case "G"
            weave_id = "5HS"
        Case "H"
            weave_id = "8HS"
        Case "O"
            weave_id = "Twill_3x1"
        Case "P"
            weave_id = "Twill_2x1"
        Case "Q"
            weave_id = "Twill_7x1"
        Case "T"
            weave_id = "Plain"
        Case "U"
            weave_id = "Tape_Basket_2x2"
        Case "R"
            weave_id = "Tape_Plain"
        Case "S"
            weave_id = "Tape_Twill_2x2"
        Case "V"
            weave_id = "Twill_3x1"
        Case "7"
            weave_id = "Double_Cloth"
        Case Else
            If Arrays.IsInArray(Mid(material_id, 6, 3), Array("336", "930")) Then
                weave_id = "6HS"
            Else 
                weave_id = "Not_Available"
            End If
    End Select

    Logger.Log "Inserting image : " & weave_id & ".jpg"
    GetWeavePatternImagePath = "S:\Data Manager\Specifications\weave-pattern\" & weave_id & ".jpg"

End Function

Private Function GetBallisticPackageImagePath() As String
' Returns the path of the correct ballistic package image given the requirements
    Dim img_path As String
    img_path = "S:\Data Manager\Specifications\ballistic-package\15x15-corner.jpg"
    GetBallisticPackageImagePath = img_path
End Function

Private Sub InsertImage(sht As Worksheet, img_path As String, row_id As Integer)
' Inserts the given image into the given sheet
    Dim xPic As Object
    Dim img As Picture
    For Each xPic In sht.Pictures
        xPic.Delete
    Next
    sht.Activate
    ActiveSheet.Range("A" & row_id).Activate
    Set img = ActiveSheet.Pictures.Insert(img_path)
End Sub

Private Sub LoadRbaToSheet(specs_dict As Object)
' Load values from rba properties into cooresponding named range
    Dim nr As Variant
    Dim rng As Range
    Dim spec As Specification
    Dim qc_sample As String
    Dim ballistic_sample As String
    ' Check for sample sizes if required
    If specs_dict.Exists("Testing Requirements") Then
        qc_sample = specs_dict("Testing Requirements").GetProperty("QC Sample Length")
    Else
        qc_sample = "1.5"
    End If
    If specs_dict.Exists("Ballistic Testing Requirements") Then
        ballistic_sample = specs_dict("Ballistic Testing Requirements").GetProperty("Ballistic Sample Length")
    Else
        ballistic_sample = "NONE"
    End If
    ' Insert values into worksheet
    Set spec = specs_dict("Weaving RBA")
    Utils.Insert shtRBA.Range("revision"), spec.Revision
    Utils.Insert shtRBA.Range("qc_sample_size"), qc_sample
    Utils.Insert shtRBA.Range("ballistic_sample_size"), ballistic_sample
    On Error Resume Next
    For Each nr In spec.Properties
        Set rng = shtRBA.Range(CStr(nr))
        Utils.Insert rng, spec.GetProperty(CStr(nr))
    Next nr
    On Error GoTo 0
End Sub

Public Sub PrintObjectToSheet(spec As Specification, sht As Worksheet, Optional prod_order As String = vbNullString, _
            Optional i As Integer = 3)
' Print object to specified worksheet starting in cell "A1"
    Dim Key As Variant
    Dim path As String
    ' Disable updates to the gui during processing
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ' This tree selects any specs which have an associated custom worksheet template
    Select Case spec.SpecType
        Case "Weaving RBA"
            LoadRbaToSheet App.specs
        Case Else
            ' All specifications which do not have a custom worksheet template
            ' will be handled by this case.
            ' Insert/Format body of the specification
            With sht
                ' Clear any existing information from the sheet
                .Cells.ClearContents
                ' Spec Header

                ' .Range("A" & 1).Value = spec.SpecType & " : "
                ' .Range("B" & 1).Value = spec.MaterialId
                ' .Range("A" & 1).Font.Bold = True
                ' .Range("A" & 1).Font.Size = 18
                ' .Range("B" & 1).Font.Bold = True
                ' .Range("B" & 1).Font.Size = 18
                ' If prod_order <> vbNullString Then
                '     .Range("A" & 2).Value = "Production Order : "
                '     .Range("B" & 2).Value = prod_order
                '     .Range("A" & 2).Font.Bold = True
                '     .Range("A" & 2).Font.Size = 18
                '     .Range("B" & 2).Font.Bold = True
                '     .Range("B" & 2).Font.Size = 18
                '     .Range("B" & 2).HorizontalAlignment = xlHAlignLeft
                '     i = 4
                ' Else
                '     i = 3
                ' End If
                ' Spec Properties
                For Each Key In spec.Properties
                    If Not spec.GetProperty(CStr(Key)) = vbNullString Then
                        .Range("A" & i).Value = CStr(Key) & " : "
                        .Range("B" & i).Value = CStr(spec.GetProperty(CStr(Key)))
                        .Range("A" & i).Font.Size = 17
                        .Range("B" & i).Font.Size = 17
                        .Range("B" & i).Font.Bold = True
                        .Range("B" & i).HorizontalAlignment = xlHAlignLeft
                        If spec.GetProperty(CStr(Key)) = Chr(252) Or spec.GetProperty(CStr(Key)) = Chr(171) Then
                            .Range("B" & i).Font.Name = "Wingdings"
                        ElseIf spec.GetProperty(CStr(Key)) = Chr(210) Then
                            .Range("B" & i).Font.Name = "Wingdings 2"
                        Else
                            .Range("B" & i).Font.Name = "Calibri"
                        End If
                    Else
                        .Range("A" & i).Value = CStr(Key)
                        .Range("B" & i).Value = "-"
                        .Range("A" & i).Font.Size = 17
                        .Range("B" & i).Font.Size = 17
                        .Range("B" & i).Font.Bold = True
                    End If
                    i = i + 1
                Next Key
                .Columns("A:G").AutoFit
                i = i + 2
            End With
            ' Insert images per requirements
            If spec.SpecType = "Ballistic Testing Requirements" Then
                InsertImage sht, GetBallisticPackageImagePath, i
            ElseIf spec.SpecType = "Testing Requirements" Then
                InsertImage sht, GetWeavePatternImagePath(spec.MaterialId), i
            End If
    
    End Select
    ' Enable updates to the gui after processing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub PrintObject(obj As Object)
' Print object to console
    Dim Key As Variant
    ' Clear the console
    FormId.txtConsole.text = vbNullString
    'PrintLine obj.SpecType, obj.MaterialId
    'PrintLine "Revision", obj.Revision
    
    For Each Key In obj.Properties
        If ShouldPrint(CStr(Key)) Then
            If Not obj.Properties(CStr(Key)) = vbNullString Then
                PrintLine CStr(Key), FormatAscii(obj.Properties(CStr(Key)))
            Else
                PrintLine CStr(Key), "-"
            End If
        End If
    Next Key
End Sub

Public Function ShouldPrint(prop As String) As Boolean
    If Left(prop, 3) = "fd_" Then
        ShouldPrint = False
    ElseIf Left(prop, 3) = "di_" Then
        ShouldPrint = False
    ElseIf Left(prop, 3) = "ld_" Then
        ShouldPrint = False
    Else
        ShouldPrint = True
    End If
End Function

Public Sub ListObjects(dict As Object)
' Given a dictionary of specifications lists the specs by type in the console
    Dim T As Variant
    For Each T In dict
        Me.FormId.txtConsole.text = FormId.txtConsole.text & CStr(T) & vbNewLine
    Next T
End Sub

Private Function FormatAscii(char As String) As String

    If char = Chr(252) Then
        char = "Required"
    ElseIf char = Chr(171) Then
        char = "Critical"
    ElseIf char = Chr(210) Then
        char = "X"
    ElseIf IsNumeric(Left(char, 1)) Then
        ' Pass
    ElseIf char = vbNullString Then
        char = "-"
    End If
    FormatAscii = char
End Function

Public Sub PrintLine(ParamArray var() As Variant)
    Dim i As Integer
    Dim s As String
    Const Padding = 45
    s = vbNullString
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    OutputHistory.Add s
    Me.FormId.txtConsole.text = FormId.txtConsole.text & s & vbNewLine
End Sub

Private Sub Class_Initialize()
' Constructor
    Set OutputHistory = New VBA.Collection
End Sub



