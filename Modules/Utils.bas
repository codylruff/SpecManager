Attribute VB_Name = "Utils"
Option Explicit
'=================================
' DESCRIPTION: Util Module holds
' miscellenous helper functions.
'=================================
' ------------------------------------------------
' WINDOWS API FUNCTIONS DO NOT CHANGE
' ------------------------------------------------
#If Win64 Then
Public Declare PtrSafe Function SendMessageA Lib "USER32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function GetDesktopWindow Lib "USER32" () As LongPtr
Private Declare PtrSafe Function InvalidateRect Lib "USER32" (ByVal hWnd As LongPtr, lpRect As Long, ByVal bErase As Long) As LongPtr
Private Declare PtrSafe Function UpdateWindow Lib "USER32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindow Lib "USER32" (ByVal hWnd As LongPtr) As LongPtr
#Else
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function InvalidateRect Lib "USER32" (ByVal hWnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "USER32" (ByVal hWnd As Long) As Long
#End If
' -------------------------------------------------
Public Function DropKeys(ByRef dict As Object, keys As Variant) As Object
' If a key exists the key and item will be removed and the modified dictionary returned.
    Dim i As Integer
    For i = LBound(keys) To UBound(keys)
        If dict.Exists(keys(i)) Then dict.Remove (keys(i))
    Next i
    Set DropKeys = dict
End Function

Public Function OpenWorkbook(path) As Workbook
    Set OpenWorkbook = Workbooks.Open(path, 0)
End Function

Public Function ConvertNumericToAlpha(col As Long) As String
    Dim vArr
    vArr = Split(Cells(1, col).Address(True, False), "$")
    ConvertNumericToAlpha = vArr(0)
End Function

Sub DeleteNames(wb As Workbook)
    Dim RangeName As Name
    On Error Resume Next
    For Each RangeName In Names
        wb.Names(RangeName.Name).Delete
    Next
    On Error GoTo 0
End Sub

Public Function CreateNamedRange(wb As Workbook, RangeName As String, sht As Worksheet, r As Long, c As Long) As Variant
    Dim CellName As String
    Dim cell As Range
    CellName = ConvertNumericToAlpha(c) & r
    
    Set cell = sht.Range(CellName)
    wb.Names.Add Name:=RangeName, RefersTo:=cell

    On Error GoTo ErrorHandler
    CreateNamedRange = Range(wb.Names(RangeName)).Value
    Exit Function
ErrorHandler:
    CreateNamedRange = False
End Function

Public Function RemoveWhiteSpace(target As String) As String
'tested
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(target, vbNullString)
    End With
End Function

Function ConvertToCamelCase(s As String) As String
'tested
' Converts sentence case to Camel Case
' numbers and symbols will be removed
    On Error GoTo RegExError
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^a-zA-Z]"
        .Global = True
        ConvertToCamelCase = Replace(StrConv(.Replace(s, " "), vbProperCase), " ", "")
    End With
    Exit Function
RegExError:
    Err.Raise SM_REGEX_ERROR
    Logger.Log "RegEx Error: ConvertToCamelCase"
End Function

Function SplitCamelCase(sString As String, Optional sDelim As String = " ") As String
'test failing
' Converts camel case to sentence case
On Error GoTo Error_Handler
    Dim oRegEx As Object
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .Pattern = "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))"
        .Global = True
        SplitCamelCase = .Replace(sString, "$1" & sDelim)
    End With
    
Error_Handler_Exit:
    On Error Resume Next
    Set oRegEx = Nothing
    Exit Function
 
Error_Handler:
    Logger.Log "RegEx Error: SplitCamelCase"
    Resume Error_Handler_Exit
End Function

Function GetLine(ParamArray var() As Variant) As String
'test
    Const Padding = 25
    Dim i As Integer
    Dim s As String
    s = vbNullString
    'If FormId.txtConsole = Nothing Then Exit Sub
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    GetLine = s & vbNewLine
End Function

Function CreateNewSheet(shtName As String, Optional DeleteOldSheet As Boolean = False) As Worksheet
'test
' Creates a new worksheet with the given name
    Application.DisplayAlerts = False
    Dim Exists As Boolean, i As Integer
    With ThisWorkbook
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = shtName Then
                Exists = True
            End If
        Next i
        If Exists = True Then
            If DeleteOldSheet Then
                .Sheets(shtName).Delete
            Else
                Set CreateNewSheet = .Sheets(shtName)
                Exit Function
            End If
        End If
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = shtName
    End With
    Set CreateNewSheet = ThisWorkbook.Sheets(shtName)
    Application.DisplayAlerts = True
End Function

Function RemoveSheet(ws As Worksheet) As Boolean

    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    RemoveSheet = True
    Exit Function
ErrorHandler:
    RemoveSheet = False
End Function

Sub ToggleExcelGui(B As Boolean)
' Disables unpleasent ui effects
    Application.ScreenUpdating = B
    Application.DisplayAlerts = B
End Sub

Function CheckForEmpties(frm) As Boolean
'Clears the values from a user form.
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                If ctl.Value = vbNullString Then
                    MsgBox "All boxes must be filed.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
            Case "ComboBox"
                If ctl.Value = vbNullString Then
                    MsgBox "Make a selection from the drop down menu.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
        End Select
    Next ctl
    CheckForEmpties = False
End Function

Sub UnloadAllForms(Optional dummyVariable As Byte)
'Unloads all open user forms
    Dim i As Long
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next
End Sub

Sub UpdateTable(shtName As String, tblName As String, Header As String, val)
'Adds an entry at the bottom of specified column header.
    Dim rng As Range
    Set rng = Sheets(shtName).Range(tblName & "[" & Header & "]")
    rng.End(xlDown).Offset(1, 0).Value = val
End Sub

Sub Update(rng As Range, val)
'Adds an entry at the bottom of specified column header.
    rng.End(xlDown).Offset(1, 0).Value = val
End Sub

Sub Insert(rng As Range, val)
'Inserts an entry into a specific named cell.
    rng.Value = val
End Sub

Public Function printf(mask As String, ParamArray tokens()) As String
'test
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

Public Sub PrintSheet(ws As Worksheet, Optional FitToPage As Boolean = False)
' Prints the sheet of the given name in the spec manager workbook
    If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    End If
    If App.current_user.Settings.Item("default_printer") = vbNullString Then
        ChangeActivePrinter
    End If
    If FitToPage Then
        If Application.PrintCommunication <> False Then Application.PrintCommunication = False
        With ws.PageSetup
            If .FitToPagesWide <> 1 Then .FitToPagesWide = 1
            If .FitToPagesTall <> True Then .FitToPagesTall = True
        End With
        Application.PrintCommunication = True
    End If
    'fncScreenUpdating State:=False
    ws.PrintOut ActivePrinter:=App.current_user.Settings.Item("default_printer")
    'fncScreenUpdating State:=True
End Sub

Public Function ArrayLength(arr As Variant) As Long
'test
    ArrayLength = UBound(arr) - LBound(arr) + 1
End Function

Sub ChangeActivePrinter()
'
' ChangeActivePrinter Macro

    Application.Dialogs(xlDialogPrinterSetup).Show
    Logger.Log "Setting default printer for Spec Manager : " & Application.ActivePrinter
    App.current_user.Settings.Item("default_printer") = Application.ActivePrinter
    App.current_user.SaveUserJson
'
End Sub

Public Function ToFileExtension(extension_type As Long) As String
'test
' Given an enum converts to the file extension string for vba files
    Select Case extension_type
        Case 1
            ToFileExtension = ".bas"
        Case 2
            ToFileExtension = ".cls"
        Case 3
            ToFileExtension = ".frm"
        Case Else
            ToFileExtension = ".txt"
    End Select
End Function

Sub SaveAll()
    Dim xWb As Workbook
    For Each xWb In Application.Workbooks
        If Not xWb.ReadOnly And Windows(xWb.Name).Visible Then
            xWb.Save
        End If
    Next
End Sub

Function TestForUnsavedChanges() As Boolean
    If ActiveWorkbook.Saved = False Then
        MsgBox "This workbook contains unsaved changes."
    End If
End Function

Public Function AskUser(question As String) As Boolean
    Dim answer As String
    If MsgBox(question, vbQuestion + vbYesNo, "???") = vbYes Then
        AskUser = True
    Else
        AskUser = False
    End If
End Function

Public Sub ToggleAutoRecover()
' This sub will switch the auto recover function on and off.

End Sub

' ==============================================
' RBA PARSER
' ==============================================
Public Sub ParseRBAs(path As String)
    Dim wb As Workbook
    Dim strFile As String
    Dim rba_dict As Object
    Dim nr As Name
    Dim rng As Object
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim ret_val As Long
    Set wb = OpenWorkbook(path)
    DeleteNames wb
    Set rba_dict = CreateObject("Scripting.Dictionary")
    Set rba_dict = AddRbaNames(rba_dict, wb, "fd", 73, 82, 2, 11)
    Set rba_dict = AddRbaNames(rba_dict, wb, "di", 73, 82, 15, 24)
    Set rba_dict = AddRbaNames(rba_dict, wb, "ld", 73, 82, 28, 37)
    Set rba_dict = AddMoreRbaNames(rba_dict, wb)
    ret_val = JsonVBA.WriteJsonObject(path & ".json", rba_dict)
    Set rba_dict = Nothing
    wb.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Public Function AddMoreRbaNames(dict As Object, wb As Workbook) As Object
    
    With wb.Names
        .Add Name:="actual_weft_count", RefersTo:=wb.Sheets("ENG").Range("AC26")
        .Add Name:="article_code", RefersTo:=wb.Sheets("ENG").Range("J14")
        .Add Name:="aux_selvedges_closing_degrees", RefersTo:=wb.Sheets("ENG").Range("AC36")
        .Add Name:="bottom_rapier_clamps", RefersTo:=wb.Sheets("ENG").Range("AC49")
        .Add Name:="bottom_spreader_bars", RefersTo:=wb.Sheets("ENG").Range("AC59")
        .Add Name:="central_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T69")
        .Add Name:="central_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z69")
        .Add Name:="central_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J69")
        .Add Name:="central_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF69")
        .Add Name:="central_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N69")
        .Add Name:="cutting_degrees", RefersTo:=wb.Sheets("ENG").Range("J38")
        .Add Name:="date", RefersTo:=wb.Sheets("ENG").Range("AC8")
        .Add Name:="dorn_left_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T66")
        .Add Name:="dorn_left_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z66")
        .Add Name:="dorn_left_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J66")
        .Add Name:="dorn_left_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF66")
        .Add Name:="dorn_left_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N66")
        .Add Name:="draw_in_harness", RefersTo:=wb.Sheets("ENG").Range("AC18")
        .Add Name:="draw_in_reed", RefersTo:=wb.Sheets("ENG").Range("AC20")
        .Add Name:="fabric_width", RefersTo:=wb.Sheets("ENG").Range("J12")
        .Add Name:="first_heddle", RefersTo:=wb.Sheets("ENG").Range("J30")
        .Add Name:="first_heddle_1", RefersTo:=wb.Sheets("ENG").Range("J30")
        .Add Name:="first_heddle_guide", RefersTo:=wb.Sheets("ENG").Range("J34")
        .Add Name:="harness_configuration", RefersTo:=wb.Sheets("ENG").Range("J22")
        .Add Name:="horizontal_back_rest_roller", RefersTo:=wb.Sheets("ENG").Range("J42")
        .Add Name:="last_heddle", RefersTo:=wb.Sheets("ENG").Range("AC30")
        .Add Name:="last_heddle_guide", RefersTo:=wb.Sheets("ENG").Range("AC34")
        .Add Name:="left_main_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T67")
        .Add Name:="left_main_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z67")
        .Add Name:="left_main_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J67")
        .Add Name:="left_main_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF67")
        .Add Name:="left_main_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N67")
        .Add Name:="left_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T64")
        .Add Name:="left_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z64")
        .Add Name:="left_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J64")
        .Add Name:="left_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF64")
        .Add Name:="left_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N64")
        .Add Name:="loom_number", RefersTo:=wb.Sheets("ENG").Range("AC10")
        .Add Name:="loom_type", RefersTo:=wb.Sheets("ENG").Range("AC12")
        .Add Name:="number_ends_wo_selvedges", RefersTo:=wb.Sheets("ENG").Range("J24")
        .Add Name:="number_harnesses", RefersTo:=wb.Sheets("ENG").Range("J20")
        .Add Name:="pinch_roller_felt_type", RefersTo:=wb.Sheets("ENG").Range("J55")
        .Add Name:="press_roller_type", RefersTo:=wb.Sheets("ENG").Range("J53")
        .Add Name:="rba_number", RefersTo:=wb.Sheets("ENG").Range("J8")
        .Add Name:="reed", RefersTo:=wb.Sheets("ENG").Range("J16")
        .Add Name:="reed_width", RefersTo:=wb.Sheets("ENG").Range("AC16")
        .Add Name:="right_main_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T68")
        .Add Name:="right_main_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z68")
        .Add Name:="right_main_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J68")
        .Add Name:="right_main_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF68")
        .Add Name:="right_main_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N68")
        .Add Name:="right_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T65")
        .Add Name:="right_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z65")
        .Add Name:="right_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J65")
        .Add Name:="right_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF65")
        .Add Name:="right_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N65")
        .Add Name:="sand_roller_type", RefersTo:=wb.Sheets("ENG").Range("AC53")
        .Add Name:="selvedges_type", RefersTo:=wb.Sheets("ENG").Range("AC22")
        .Add Name:="shed_closing_degrees", RefersTo:=wb.Sheets("ENG").Range("J36")
        .Add Name:="speed", RefersTo:=wb.Sheets("ENG").Range("AC14")
        .Add Name:="springs_type", RefersTo:=wb.Sheets("ENG").Range("J44")
        .Add Name:="style_number", RefersTo:=wb.Sheets("ENG").Range("J10")
        .Add Name:="temples_composition", RefersTo:=wb.Sheets("ENG").Range("AC44")
        .Add Name:="upper_rapier_clamps", RefersTo:=wb.Sheets("ENG").Range("J49")
        .Add Name:="upper_spreader_bars", RefersTo:=wb.Sheets("ENG").Range("J59")
        .Add Name:="vertical_back_rest_roller", RefersTo:=wb.Sheets("ENG").Range("AC42")
        .Add Name:="warp_tension", RefersTo:=wb.Sheets("ENG").Range("J26")
        .Add Name:="weave_pattern", RefersTo:=wb.Sheets("ENG").Range("J18")
        .Add Name:="weft_count_set_point", RefersTo:=wb.Sheets("ENG").Range("AC24")
    End With
    With dict
        .Add Key:="actual_weft_count", Item:=IIf(Range(wb.Names("actual_weft_count")).Value = vbNullString, vbNullString, Range(wb.Names("actual_weft_count")).Value)
        .Add Key:="article_code", Item:=IIf(Range(wb.Names("article_code")).Value = vbNullString, vbNullString, Range(wb.Names("article_code")).Value)
        .Add Key:="aux_selvedges_closing_degrees", Item:=IIf(Range(wb.Names("aux_selvedges_closing_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("aux_selvedges_closing_degrees")).Value)
        .Add Key:="bottom_rapier_clamps", Item:=IIf(Range(wb.Names("bottom_rapier_clamps")).Value = vbNullString, vbNullString, Range(wb.Names("bottom_rapier_clamps")).Value)
        .Add Key:="bottom_spreader_bars", Item:=IIf(Range(wb.Names("bottom_spreader_bars")).Value = vbNullString, vbNullString, Range(wb.Names("bottom_spreader_bars")).Value)
        .Add Key:="central_selvedges_drawing_in", Item:=IIf(Range(wb.Names("central_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_drawing_in")).Value)
        .Add Key:="central_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("central_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_ends_per_dent")).Value)
        .Add Key:="central_selvedges_number_ends", Item:=IIf(Range(wb.Names("central_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_number_ends")).Value)
        .Add Key:="central_selvedges_weave", Item:=IIf(Range(wb.Names("central_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_weave")).Value)
        .Add Key:="central_selvedges_yarn_count", Item:=IIf(Range(wb.Names("central_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_yarn_count")).Value)
        .Add Key:="cutting_degrees", Item:=IIf(Range(wb.Names("cutting_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("cutting_degrees")).Value)
        .Add Key:="date", Item:=IIf(Range(wb.Names("date")).Value = vbNullString, vbNullString, Range(wb.Names("date")).Value)
        .Add Key:="dorn_left_selvedges_drawing_in", Item:=IIf(Range(wb.Names("dorn_left_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_drawing_in")).Value)
        .Add Key:="dorn_left_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("dorn_left_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_ends_per_dent")).Value)
        .Add Key:="dorn_left_selvedges_number_ends", Item:=IIf(Range(wb.Names("dorn_left_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_number_ends")).Value)
        .Add Key:="dorn_left_selvedges_weave", Item:=IIf(Range(wb.Names("dorn_left_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_weave")).Value)
        .Add Key:="dorn_left_selvedges_yarn_count", Item:=IIf(Range(wb.Names("dorn_left_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_yarn_count")).Value)
        .Add Key:="draw_in_harness", Item:=IIf(Range(wb.Names("draw_in_harness")).Value = vbNullString, vbNullString, Range(wb.Names("draw_in_harness")).Value)
        .Add Key:="draw_in_reed", Item:=IIf(Range(wb.Names("draw_in_reed")).Value = vbNullString, vbNullString, Range(wb.Names("draw_in_reed")).Value)
        .Add Key:="fabric_width", Item:=IIf(Range(wb.Names("fabric_width")).Value = vbNullString, vbNullString, Range(wb.Names("fabric_width")).Value)
        .Add Key:="first_heddle", Item:=IIf(Range(wb.Names("first_heddle")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle")).Value)
        .Add Key:="first_heddle_1", Item:=IIf(Range(wb.Names("first_heddle_1")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle_1")).Value)
        .Add Key:="first_heddle_guide", Item:=IIf(Range(wb.Names("first_heddle_guide")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle_guide")).Value)
        .Add Key:="harness_configuration", Item:=IIf(Range(wb.Names("harness_configuration")).Value = vbNullString, vbNullString, Range(wb.Names("harness_configuration")).Value)
        .Add Key:="horizontal_back_rest_roller", Item:=IIf(Range(wb.Names("horizontal_back_rest_roller")).Value = vbNullString, vbNullString, Range(wb.Names("horizontal_back_rest_roller")).Value)
        .Add Key:="last_heddle", Item:=IIf(Range(wb.Names("last_heddle")).Value = vbNullString, vbNullString, Range(wb.Names("last_heddle")).Value)
        .Add Key:="last_heddle_guide", Item:=IIf(Range(wb.Names("last_heddle_guide")).Value = vbNullString, vbNullString, Range(wb.Names("last_heddle_guide")).Value)
        .Add Key:="left_main_selvedges_drawing_in", Item:=IIf(Range(wb.Names("left_main_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_drawing_in")).Value)
        .Add Key:="left_main_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("left_main_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_ends_per_dent")).Value)
        .Add Key:="left_main_selvedges_number_ends", Item:=IIf(Range(wb.Names("left_main_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_number_ends")).Value)
        .Add Key:="left_main_selvedges_weave", Item:=IIf(Range(wb.Names("left_main_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_weave")).Value)
        .Add Key:="left_main_selvedges_yarn_count", Item:=IIf(Range(wb.Names("left_main_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_yarn_count")).Value)
        .Add Key:="left_selvedges_drawing_in", Item:=IIf(Range(wb.Names("left_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_drawing_in")).Value)
        .Add Key:="left_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("left_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_ends_per_dent")).Value)
        .Add Key:="left_selvedges_number_ends", Item:=IIf(Range(wb.Names("left_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_number_ends")).Value)
        .Add Key:="left_selvedges_weave", Item:=IIf(Range(wb.Names("left_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_weave")).Value)
        .Add Key:="left_selvedges_yarn_count", Item:=IIf(Range(wb.Names("left_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_yarn_count")).Value)
        .Add Key:="loom_number", Item:=IIf(Range(wb.Names("loom_number")).Value = vbNullString, vbNullString, Range(wb.Names("loom_number")).Value)
        .Add Key:="loom_type", Item:=IIf(Range(wb.Names("loom_type")).Value = vbNullString, vbNullString, Range(wb.Names("loom_type")).Value)
        .Add Key:="number_ends_wo_selvedges", Item:=IIf(Range(wb.Names("number_ends_wo_selvedges")).Value = vbNullString, vbNullString, Range(wb.Names("number_ends_wo_selvedges")).Value)
        .Add Key:="number_harnesses", Item:=IIf(Range(wb.Names("number_harnesses")).Value = vbNullString, vbNullString, Range(wb.Names("number_harnesses")).Value)
        .Add Key:="pinch_roller_felt_type", Item:=IIf(Range(wb.Names("pinch_roller_felt_type")).Value = vbNullString, vbNullString, Range(wb.Names("pinch_roller_felt_type")).Value)
        .Add Key:="press_roller_type", Item:=IIf(Range(wb.Names("press_roller_type")).Value = vbNullString, vbNullString, Range(wb.Names("press_roller_type")).Value)
        .Add Key:="rba_number", Item:=IIf(Range(wb.Names("rba_number")).Value = vbNullString, vbNullString, Range(wb.Names("rba_number")).Value)
        .Add Key:="reed", Item:=IIf(Range(wb.Names("reed")).Value = vbNullString, vbNullString, Range(wb.Names("reed")).Value)
        .Add Key:="reed_width", Item:=IIf(Range(wb.Names("reed_width")).Value = vbNullString, vbNullString, Range(wb.Names("reed_width")).Value)
        .Add Key:="right_main_selvedges_drawing_in", Item:=IIf(Range(wb.Names("right_main_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_drawing_in")).Value)
        .Add Key:="right_main_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("right_main_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_ends_per_dent")).Value)
        .Add Key:="right_main_selvedges_number_ends", Item:=IIf(Range(wb.Names("right_main_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_number_ends")).Value)
        .Add Key:="right_main_selvedges_weave", Item:=IIf(Range(wb.Names("right_main_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_weave")).Value)
        .Add Key:="right_main_selvedges_yarn_count", Item:=IIf(Range(wb.Names("right_main_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_yarn_count")).Value)
        .Add Key:="right_selvedges_drawing_in", Item:=IIf(Range(wb.Names("right_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_drawing_in")).Value)
        .Add Key:="right_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("right_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_ends_per_dent")).Value)
        .Add Key:="right_selvedges_number_ends", Item:=IIf(Range(wb.Names("right_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_number_ends")).Value)
        .Add Key:="right_selvedges_weave", Item:=IIf(Range(wb.Names("right_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_weave")).Value)
        .Add Key:="right_selvedges_yarn_count", Item:=IIf(Range(wb.Names("right_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_yarn_count")).Value)
        .Add Key:="sand_roller_type", Item:=IIf(Range(wb.Names("sand_roller_type")).Value = vbNullString, vbNullString, Range(wb.Names("sand_roller_type")).Value)
        .Add Key:="selvedges_type", Item:=IIf(Range(wb.Names("selvedges_type")).Value = vbNullString, vbNullString, Range(wb.Names("selvedges_type")).Value)
        .Add Key:="shed_closing_degrees", Item:=IIf(Range(wb.Names("shed_closing_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("shed_closing_degrees")).Value)
        .Add Key:="speed", Item:=IIf(Range(wb.Names("speed")).Value = vbNullString, vbNullString, Range(wb.Names("speed")).Value)
        .Add Key:="springs_type", Item:=IIf(Range(wb.Names("springs_type")).Value = vbNullString, vbNullString, Range(wb.Names("springs_type")).Value)
        .Add Key:="style_number", Item:=IIf(Range(wb.Names("style_number")).Value = vbNullString, vbNullString, Range(wb.Names("style_number")).Value)
        .Add Key:="temples_composition", Item:=IIf(Range(wb.Names("temples_composition")).Value = vbNullString, vbNullString, Range(wb.Names("temples_composition")).Value)
        .Add Key:="upper_rapier_clamps", Item:=IIf(Range(wb.Names("upper_rapier_clamps")).Value = vbNullString, vbNullString, Range(wb.Names("upper_rapier_clamps")).Value)
        .Add Key:="upper_spreader_bars", Item:=IIf(Range(wb.Names("upper_spreader_bars")).Value = vbNullString, vbNullString, Range(wb.Names("upper_spreader_bars")).Value)
        .Add Key:="vertical_back_rest_roller", Item:=IIf(Range(wb.Names("vertical_back_rest_roller")).Value = vbNullString, vbNullString, Range(wb.Names("vertical_back_rest_roller")).Value)
        .Add Key:="warp_tension", Item:=IIf(Range(wb.Names("warp_tension")).Value = vbNullString, vbNullString, Range(wb.Names("warp_tension")).Value)
        .Add Key:="weave_pattern", Item:=IIf(Range(wb.Names("weave_pattern")).Value = vbNullString, vbNullString, Range(wb.Names("weave_pattern")).Value)
        .Add Key:="weft_count_set_point", Item:=IIf(Range(wb.Names("weft_count_set_point")).Value = vbNullString, vbNullString, Range(wb.Names("weft_count_set_point")).Value)
    End With
    Set AddMoreRbaNames = dict
End Function

Public Function AddRbaNames(dict As Object, wb As Workbook, tag As String, r_start As Long, r_end As Long, c_start As Long, c_end As Long) As Object
    Dim sht As Worksheet
    Dim nr As String
    Dim ret_val As Variant
    Set sht = wb.Sheets("ENG")
    Dim r, c, rw, cl As Long
    For r = r_start To r_end
        cl = 0
        For c = c_start To c_end
            rw = Abs(r_end - r)
            nr = tag & "_" & cl & rw
            ret_val = CreateNamedRange(wb, nr, sht, CLng(r), CLng(c))
            dict.Add nr, IIf(ret_val = vbNullString, vbNullString, ret_val)
            cl = cl + 1
        Next c
    Next r
    Set AddRbaNames = dict
End Function

#If Win64 Then
Public Function fncScreenUpdating(State As Boolean, Optional Window_hWnd As LongPtr = 0)
#Else
Public Function fncScreenUpdating(State As Boolean, Optional Window_hWnd As Long = 0)
#End If
    Const WM_SETREDRAW = &HB
    Const WM_PAINT = &HF
    
    If Window_hWnd = 0 Then
        Window_hWnd = GetDesktopWindow()
    Else
        If IsWindow(hWnd:=Window_hWnd) = False Then
            Exit Function
        End If
    End If
    
    If State = True Then
        Call SendMessage(hWnd:=Window_hWnd, wMsg:=WM_SETREDRAW, wParam:=1, lParam:=0)
        Call InvalidateRect(hWnd:=Window_hWnd, lpRect:=0, bErase:=True)
        Call UpdateWindow(hWnd:=Window_hWnd)
    
    Else
        Call SendMessage(hWnd:=Window_hWnd, wMsg:=WM_SETREDRAW, wParam:=0, lParam:=0)
    
    End If

End Function
