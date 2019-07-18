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
'CONSTANTS
'----------------
Private Const CHAR_REQUIRED          As Long = 252              ' Unicode character used to denote a required test
Private Const CHAR_CRITICAL          As Long = 171              ' Unicode character used to denote a critical test
Private Const CHAR_NOT_REQUIRED      As Long = 210              ' Unicode character used to denote that a test is not-required
Private Const CHAR_NULL_VALUE        As Long = 45               ' Unicode character used to denote a null property value
Private Const FONT_SIZE_HEADING      As Long = 17               ' Font size for document headings
Private Const FONT_SIZE_VALUE        As Long = 17               ' Font size for property values
Private Const FONT_SIZE_PROPERTY     As Long = 17               ' Font size for property names
Private Const FONT_REQUIRED          As String = "Wingdings"    ' Font used to denote a required test
Private Const FONT_CRITICAL          As String = "Wingdings"    ' Font used to denote a critical test
Private Const FONT_NOT_REQUIRED      As String = "Wingdings 2"  ' Font used to denote that a test is not-required
Private Const FONT_DEFAULT           As String = "Calibri"      ' Font used by default
Private Const DEFAULT_QC_SAMPLE      As String = "1.5"          ' Minimum protection qc sample size
'----------------
'CLASS VARIABLES
'----------------
Private FormId_         As MSForms.UserForm
Private OutputHistory_  As VBA.Collection
'------------
' AUTO-PROPERTIES
'------------
Public Property Set OutputHistory(value As VBA.Collection): Set OutputHistory_ = value: End Property
Public Property Get OutputHistory() As VBA.Collection: Set OutputHistory = OutputHistory_: End Property
Public Property Set FormId(value As MSForms.UserForm): Set FormId_ = value: End Property
Public Property Get FormId() As MSForms.UserForm: Set FormId = FormId_: End Property
Public Property Get CurrentText() As String
     CurrentText = Me.FormId.txtConsole.text
End Property
Public Property Let CurrentText(value As String)
    CurrentText = value
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
    ' Clear Header
    Utils.ClearHeaderFooter ws, Header:=True, footer:=False
    ' Insert Header
    With ws.PageSetup
        .LeftHeader = spec.SpecType & " (v" & spec.Template.Revision & ")" & Chr(10) & _
                      spec.Template.ProductLine
        .CenterHeader = "Material : " & spec.MaterialId & Chr(10) & _
                        "Production Order : " & production_order
        .RightHeader = "Revision : " & spec.Revision & Chr(10) & _
                       "Printed : " & CStr(Now())
    End With
End Sub

Private Sub InsertFooter(ws As Worksheet, Optional disclaimer As String = vbNullString, _
            Optional leftFooterText = vbNullString, Optional rightFooterText = vbNullString)
' Footer Content and formatting for Spec Docs
    'Clear footer
    Utils.ClearHeaderFooter ws, Header:=False, footer:=True
    ' Initialize disclaimer text
    If disclaimer = vbNullString Then
        disclaimer = "This document is the property of SAATI Americas - " & _
                     "Printed and/or Unsigned Copies Are Uncontrolled."
    End If
    ' Insert Footer
    With ws.PageSetup
        .LeftFooter = leftFooterText
        .CenterFooter = disclaimer
        .RightFooter = rightFooterText
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

Private Sub LoadRbaToSheet(specs_dict As Object, prod_order As String)
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
        qc_sample = DEFAULT_QC_SAMPLE
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
    Utils.Insert shtRBA.Range("work_order"), prod_order
    Utils.Insert shtRBA.Range("date_printed"), Format(CStr(Now()), "mm/dd/yyyy")
    On Error Resume Next
    For Each nr In spec.Properties
        Set rng = shtRBA.Range(CStr(nr))
        Utils.Insert rng, spec.GetProperty(CStr(nr))
    Next nr
    On Error GoTo 0

End Sub

Private Sub LoadTestingRequirementsToSheet(spec As Specification, sht As Worksheet, prod_order As String, _
            package_variant As DocumentPackageVariant)
' Load QC Testing document per  given variant
    Dim i As Integer
    LoadGenericSpecificationDocument spec, sht, prod_order
    ' Change Critical tests to Required tests for Tie-Backs
    If package_variant = WeavingTieBack Then
        For i = 2 To spec.Properties.Count
            With sht.Range("B" & i)
                If .value = Chr(CHAR_CRITICAL) Then
                    .value = Chr(CHAR_REQUIRED)
                    .Font.Name = FONT_REQUIRED
                End If
            End With
        Next i
    End If
    ' Add document legend
    sht.Range("E2").value = "NOTES :"
    sht.Range("E3").value = "Required At Start-Up = "
    sht.Range("F3").value = Chr(CHAR_CRITICAL)
    sht.Range("F3").Font.Name = FONT_CRITICAL
    sht.Range("E4").value = "Required For Release = "
    sht.Range("F4").value = Chr(CHAR_REQUIRED)
    sht.Range("F4").Font.Name = FONT_REQUIRED
    sht.Range("E5").value = "Test Not Required = "
    sht.Range("F5").value = Chr(CHAR_NOT_REQUIRED)
    sht.Range("F5").Font.Name = FONT_NOT_REQUIRED
    sht.Columns("A:G").AutoFit
    ' Insert fabric construction image
    InsertImage sht, GetWeavePatternImagePath(spec.MaterialId), spec.Properties.Count + 4
End Sub

Private Sub LoadBallisticTestingRequirementsToSheet(spec As Specification, sht As Worksheet, prod_order As String)
' Load Ballistic Testing Requirements to Sheet
    LoadGenericSpecificationDocument spec, sht, prod_order
    ' Insert package build image
    InsertImage sht, GetBallisticPackageImagePath, spec.Properties.Count + 4
End Sub

Public Sub PrintObjectToSheet(spec As Specification, sht As Worksheet, Optional prod_order As String = vbNullString, _
           Optional package_variant As DocumentPackageVariant, Optional i As Integer = 2)
' Print object to specified worksheet starting in cell "A1"
    ' Disable updates to the gui during processing
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ' This tree selects any specs which have an associated custom worksheet template
    Select Case spec.SpecType
        Case "Weaving RBA"
            LoadRbaToSheet App.specs, prod_order
        Case "Testing Requirements"
            LoadTestingRequirementsToSheet spec, sht, prod_order, package_variant
        Case "Ballistic Testing Requirements"
            LoadBallisticTestingRequirementsToSheet spec, sht, prod_order
        Case Else
            ' All specifications which do not have a custom worksheet template
            ' will be handled by this case.
            ' Insert/Format body of the specification
            LoadGenericSpecificationDocument spec, sht, prod_order
    End Select
    ' Enable updates to the gui after processing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub PrintObjectToConsole(obj As Object)
' Print object to console
    Dim Key As Variant
    ' Clear the console
    FormId.txtConsole.text = vbNullString
    'WriteLine obj.SpecType, obj.MaterialId
    'WriteLine "Revision", obj.Revision
    
    For Each Key In obj.Properties
        If ShouldPrint(CStr(Key)) Then
            If Not obj.Properties(CStr(Key)) = vbNullString Then
                WriteLine CStr(Key), FormatAscii(obj.Properties(CStr(Key)))
            Else
                WriteLine CStr(Key), Chr(CHAR_NULL_VALUE)
            End If
        End If
    Next Key
End Sub

Public Function ShouldPrint(prop As String) As Boolean
' Acts as a filter for values that are difficult to visualize in tabular form
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
' Given a dictionary of documents, list the docs by type in the console
    Dim T As Variant
    For Each T In dict
        Me.FormId.txtConsole.text = FormId.txtConsole.text & CStr(T) & vbNewLine
    Next T
End Sub

Private Function FormatAscii(char As String) As String

    If char = Chr(CHAR_REQUIRED) Then
        char = "Required"
    ElseIf char = Chr(CHAR_CRITICAL) Then
        char = "Critical"
    ElseIf char = Chr(CHAR_NOT_REQUIRED) Then
        char = "X"
    ElseIf IsNumeric(Left(char, 1)) Then
        ' Pass
    ElseIf char = vbNullString Then
        char = Chr(CHAR_NULL_VALUE)
    End If
    FormatAscii = char
End Function

Public Sub WriteLine(ParamArray var() As Variant)
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

Private Sub PrintPropertyToSheet(ws As Worksheet, prop_name As Variant, prop_val As Variant, row As Integer)
' Prints a given property to the worksheet
    With ws
        If Not prop_val = vbNullString Then
            .Range("A" & row).value = CStr(prop_name) & " : "
            .Range("A" & row).Font.Size = FONT_SIZE_PROPERTY
            .Range("B" & row).value = prop_val
            .Range("B" & row).Font.Size = FONT_SIZE_VALUE
            .Range("B" & row).Font.Bold = True
            .Range("B" & row).HorizontalAlignment = xlHAlignLeft
            If prop_val = Chr(CHAR_REQUIRED) Then
                .Range("B" & row).Font.Name = FONT_REQUIRED
            ElseIf prop_val = Chr(CHAR_CRITICAL) Then
                .Range("B" & row).Font.Name = FONT_CRITICAL
            ElseIf prop_val = Chr(CHAR_NOT_REQUIRED) Then
                .Range("B" & row).Font.Name = FONT_NOT_REQUIRED
            Else
                .Range("B" & row).Font.Name = FONT_DEFAULT
            End If
        Else
            .Range("A" & row).value = CStr(prop_name)
            .Range("A" & row).Font.Size = FONT_SIZE_PROPERTY
            .Range("B" & row).value = Chr(CHAR_NULL_VALUE)
            .Range("B" & row).Font.Size = FONT_SIZE_VALUE
            .Range("B" & row).Font.Bold = True
        End If
    End With
End Sub

Private Sub LoadGenericSpecificationDocument(spec As Specification, sht As Worksheet, _
            prod_order As String, Optional i As Integer = 2)
' Handles all non-custom generic spec-docs
    Dim Key As Variant
        ' Clear any existing information from the sheet
        sht.Cells.ClearContents
        ' Spec Header
        InsertHeader sht, spec, prod_order
        InsertFooter sht
        ' Spec Properties
        For Each Key In spec.Properties
            PrintPropertyToSheet sht, Key, spec.GetProperty(CStr(Key)), i
            ' If Not spec.GetProperty(CStr(Key)) = vbNullString Then
            '     .Range("A" & i).value = CStr(Key) & " : "
            '     .Range("A" & i).Font.Size = FONT_SIZE_PROPERTY
            '     .Range("B" & i).value = CStr(spec.GetProperty(CStr(Key)))
            '     .Range("B" & i).Font.Size = FONT_SIZE_VALUE
            '     .Range("B" & i).Font.Bold = True
            '     .Range("B" & i).HorizontalAlignment = xlHAlignLeft
            '     If spec.GetProperty(CStr(Key)) = Chr(CHAR_REQUIRED) Then
            '         .Range("B" & i).Font.Name = FONT_REQUIRED
            '     ElseIf spec.GetProperty(CStr(Key)) = Chr(CHAR_CRITICAL) Then
            '         .Range("B" & i).Font.Name = FONT_CRITICAL
            '     ElseIf spec.GetProperty(CStr(Key)) = Chr(CHAR_NOT_REQUIRED) Then
            '         .Range("B" & i).Font.Name = FONT_NOT_REQUIRED
            '     Else
            '         .Range("B" & i).Font.Name = FONT_DEFAULT
            '     End If
            ' Else
            '     .Range("A" & i).value = CStr(Key)
            '     .Range("A" & i).Font.Size = FONT_SIZE_PROPERTY
            '     .Range("B" & i).value = Chr(CHAR_NULL_VALUE)
            '     .Range("B" & i).Font.Size = FONT_SIZE_VALUE
            '     .Range("B" & i).Font.Bold = True
            ' End If
            i = i + 1
        Next Key
        sht.Columns("A:G").AutoFit

End Sub

Public Sub ToPDF(ws As Worksheet, Optional destination_path As String = vbNullString)
' Exports the given ws as a PDF and saves it to S:\Data Manager\Specifications

    On Error GoTo SaveFileError
    ' Parse args
    If destination_path = vbNullString Then
        With App.current_spec
            destination_path = PUBLIC_DIR & "\Specifications\" & .MaterialId & "_" & .SpecType & "_rev" & .Revision
        End With
    End If

    ' If file already exists then load it.
    If Utils.FileExists(destination_path & ".pdf") Then
        ActiveWorkbook.FollowHyperlink destination_path & ".pdf"
        Exit Sub
    End If
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=destination_path, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    Logger.Log "PDF Saved : " & destination_path
    ws.Visible = xlSheetHidden
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail"
End Sub
