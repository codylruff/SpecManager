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
Public Function CleanString(ByVal target As String, find_strings As Variant, Optional remove_whitespace As Boolean = False) As String
    Dim i As Integer
    Dim clean_string As String
    clean_string = target
    For i = LBound(find_strings) To UBound(find_strings)
        clean_string = Replace(clean_string, find_strings(i), vbNullString)
    Next i
    If clean_string <> target Then
        CleanString = IIf(remove_whitespace, Utils.RemoveWhiteSpace(clean_string), clean_string)
    Else
        CleanString = target
    End If
End Function

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
