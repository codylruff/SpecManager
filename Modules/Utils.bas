Attribute VB_Name = "Utils"
Option Explicit
'=================================
' DESCRIPTION: Util Module holds
' miscellenous helper functions.
'=================================
Sub ToggleExcelGui(b As Boolean)
' Disables unpleasent ui effects
    Application.ScreenUpdating = b
    Application.DisplayAlerts = b
End Sub

Public Function RemoveWhiteSpace(target As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(target, vbNullString)
    End With
End Function

Function ConvertToCamelCase(s As String) As String
' Converts sentence case to Camel Case
    On Error GoTo RegExError
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^a-zA-Z]"
        .Global = True
        ConvertToCamelCase = Replace(StrConv(.Replace(s, " "), vbProperCase), " ", "")
    End With
    Exit Function
RegExError:
    Logger.Log "RegEx Error: ConvertToCamelCase"
End Function

Function SplitCamelCase(sString As String, Optional sDelim As String = " ") As String
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

Function CreateNewSheet(shtName As String) As Worksheet
' Creates a new worksheet with the given name
    Application.DisplayAlerts = False
    Dim exists As Boolean, i As Integer
    With ThisWorkbook
        For i = 1 To Worksheets.count
            If Worksheets(i).Name = shtName Then
                exists = True
            End If
        Next i
        If exists = True Then
            .Sheets(shtName).Delete
        End If
        .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = shtName
    End With
    Set CreateNewSheet = Sheets(shtName)
    Application.DisplayAlerts = True
End Function

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
    For i = VBA.UserForms.count - 1 To 0 Step -1
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
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

Public Sub PrintSheet(ws As Worksheet)
' Prints the sheet of the given name in the spec manager workbook
    If App.current_user.Settings.Item("default_printer") = vbNullString Then
        ChangeActivePrinter
    End If
    ws.PrintOut ActivePrinter:=App.current_user.Settings.Item("default_printer")
End Sub

Public Function ArrayLength(arr As Variant) As Long
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
