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
Public Declare PtrSafe Function SendMessageA Lib "user32" (ByVal Hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal Hwnd As LongPtr, lpRect As Long, ByVal bErase As Long) As LongPtr
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal Hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal Hwnd As LongPtr) As LongPtr
#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal Hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal Hwnd As Long) As Long
#End If
' -------------------------------------------------
Public Function GetLastRow(sht As Worksheet, start_column As String, start_row As Long) As Long
' Returns the last populated row in a range of data
        Dim current_cell As Range
        Dim i As Long
        Set current_cell = sht.Range(start_column & start_row)
        i = start_row
        Do Until IsEmpty(current_cell)
            Set current_cell = Range(start_column & i)
            i = i + 1
        Loop
        GetLastRow = i
End Function

Public Sub CopyRange(src_sht As Worksheet, dest_sht As Worksheet, src_address As String, dest_address As String)
' Copies from one range to another
    src_sht.Range(src_address).Copy destination:=dest_sht.Range(dest_address)
End Sub

Public Function GetRangeAddress(sht As Worksheet, first_cell As Variant, last_cell As Variant) As String
' Creates a range from the top left cell and bottom right cells.
    GetRangeAddress = sht.Range(first_cell, last_cell).Address
End Function

Public Function IsNothing(obj As Object) As Boolean
' Returns true if the object is unitialized
    If obj Is Nothing Then
        IsNothing = True
    Else
        IsNothing = False
    End If
End Function

Public Function Contains(col As Collection, Key As Variant) As Boolean
    Dim obj As Variant

    On Error GoTo err

    Contains = True
    obj = col(Key)
    Exit Function

err:
    Contains = False
End Function

Public Function SheetByName(sheet_name As String) As Worksheet
' Returns a sheet from ThisWorkbook.Sheets by name
    Set SheetByName = ThisWorkbook.Sheets(sheet_name)
End Function

Public Function FileExists(file_path As String) As Boolean
' Tests whether a file exists
    FileExists = IIf(Dir(file_path) <> "", True, False)
End Function

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

Public Function CreateNamedRange(wb As Workbook, RangeName As String, sht As Worksheet, r As Long, C As Long) As Variant
    Dim CellName As String
    Dim cell As Range
    CellName = ConvertNumericToAlpha(C) & r
    
    Set cell = sht.Range(CellName)
    wb.Names.Add Name:=RangeName, RefersTo:=cell

    On Error GoTo ErrorHandler
    CreateNamedRange = Range(wb.Names(RangeName)).value
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
    err.Raise SM_REGEX_ERROR
    App.logger.Log "RegEx Error: ConvertToCamelCase", DebugLog
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
    App.logger.Log "RegEx Error: SplitCamelCase", DebugLog
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

Function FormatTestResult(ParamArray var() As Variant) As String
    Const Padding = 36
    Dim i As Integer
    Dim s As String
    s = vbNullString
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    FormatTestResult = s
End Function


Function CreateNewSheet(shtName As String, Optional DeleteOldSheet As Boolean = False) As Worksheet
'test
' Creates a new worksheet with the given name
    ' Turn on Performance Mode
    App.PerformanceMode True

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
    ' Turn off Performance Mode
    App.PerformanceMode False
End Function

Function RemoveSheet(ws As Worksheet) As Boolean

    On Error GoTo ErrorHandler
    ' Turn on Performance Mode
    App.PerformanceMode True
    ws.Delete
    ' Turn off Performance Mode
    App.PerformanceMode False
    RemoveSheet = True
    Exit Function
ErrorHandler:
    RemoveSheet = False
End Function

Function CheckForEmpties(frm) As Boolean
'Clears the values from a user form.
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                If ctl.value = vbNullString Then
                    MsgBox "All boxes must be filed.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
            Case "ComboBox"
                If ctl.value = vbNullString Then
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
    rng.End(xlDown).Offset(1, 0).value = val
End Sub

Sub Update(rng As Range, val)
'Adds an entry at the bottom of specified column header.
    rng.End(xlDown).Offset(1, 0).value = val
End Sub

Sub Insert(rng As Range, val)
'Inserts an entry into a specific named cell.
    rng.value = val
End Sub

Function ReadNamedRange(Name As String) As Variant
End Function

Public Function printf(mask As String, ParamArray tokens()) As String
'test
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

Public Function ArrayLength(arr As Variant) As Long
'test
    ArrayLength = UBound(arr) - LBound(arr) + 1
End Function

Sub ChangeActivePrinter()
'
' ChangeActivePrinter Macro

    Application.Dialogs(xlDialogPrinterSetup).show
    App.logger.Log "Setting default printer for Spec Manager : " & Application.ActivePrinter
    App.current_user.Settings.item("default_printer") = Application.ActivePrinter
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

    ' Turn on Performance Mode
    App.PerformanceMode True

    For Each xWb In Application.Workbooks
        If Not xWb.ReadOnly And Windows(xWb.Name).Visible Then
            xWb.Save
        End If
    Next
    ' Turn off Performance Mode
    App.PerformanceMode False
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

Public Sub ClearHeaderFooter(ws As Worksheet, _
           Optional Header As Boolean = True, Optional footer As Boolean = True)
' Clears the contents of header and footer (optionally select one or the other)
    
    ' Clear Header
    If Header Then
        With ws.PageSetup
            .LeftHeader = vbNullString
            .CenterHeader = vbNullString
            .RightHeader = vbNullString
        End With
    End If
    ' Clear Footer
    If footer Then
        With ws.PageSetup
            .LeftFooter = vbNullString
            .CenterFooter = vbNullString
            .RightFooter = vbNullString
        End With
    End If

End Sub

Public Sub ToggleAutoRecover()
' This sub will switch the auto recover function on and off.
End Sub

Sub CreateWorksheetScopedNameRanges()
'PURPOSE: Create Worksheet Scoped versions of Workbook Scoped Named Ranges
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim nm As Name
Dim rng As Range
Dim FilterPhrase As String

'Filter Named Ranges that contain specific phrase
  'FilterPhrase = "Table"

'Loop Through each named Range
  For Each nm In ActiveWorkbook.Names
  
    'Is Name scoped at the Workbook level?
      If TypeOf nm.Parent Is Workbook Then
        Debug.Print nm.Name
        'Does Name meet Filter Phrase Requirement? If so, recreate named range with Worksheet Scope
          'If InStr(1, nm.Name, FilterPhrase) > 0 Then
            Set rng = Range(nm.RefersTo)
            rng.Parent.Names.Add Name:=nm.Name, RefersToR1C1:="=" & rng.Address(ReferenceStyle:=xlR1C1)
          'End If
          
      End If
      
  Next nm

End Sub

Public Function ConvertTwipsToPixels(lngTwips As Long, lngDirection As PixelDirection) As Long
'-----------------------------------------------------------------------------
' Pixel to Twips conversions
'-----------------------------------------------------------------------------
' cf http://support.microsoft.com/default.aspx?scid=kb;en-us;210590
' To call this function, pass the number of twips you want to convert,
' and another parameter indicating the horizontal or vertical measurement
' (0 for horizontal, non-zero for vertical). The following is a sample call:
'

   'Handle to device
   #If VBA7 Then
       Dim lngDC As LongPtr
   #Else
       Dim lngDC As Long
   #End If

   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = PixelDirection.Horizontal) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
   
End Function

Public Function ConvertTwipsToCm(Twips As Double, Optional iRound As Integer = 2) As Double
'---------------------------------------------------------------------------------------
' Procedure : ConvertTwipsToCm
' Author    : KRISH J
' Purpose   : The NOT version of ConvertCmToTwips
' Returns   : Twips value in double
'---------------------------------------------------------------------------------------
'

    ConvertTwipsToCm = Round(Twips / 567, iRound)
End Function

Public Function ConvertPointerToObject(ByVal lngThisPointer As Long) As Object
'---------------------------------------------------------------------------------------
' Procedure : ConvertPointerToObject
' Author    : KRISH J
' Date      : 20/12/2017
' Purpose   : Retrieves the object from memory pointer
'---------------------------------------------------------------------------------------
'
    Dim objThisObject As Object
    RtlMoveMemory objThisObject, lngThisPointer, POINTERSIZE
    Set ConvertPointerToObject = objThisObject
    RtlMoveMemory objThisObject, ZEROPOINTER, POINTERSIZE
End Function

Public Function ConvertPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
   'Handle to device
   #If VBA7 Then
       Dim lngDC As LongPtr
   #Else
       Dim lngDC As Long
   #End If
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertPixelsToTwips = (lngPixels * nTwipsPerInch) / lngPixelsPerInch
End Function

Public Function ConvertObjectToPointer(ByRef objThisObject As Object) As Long
'---------------------------------------------------------------------------------------
' Procedure : ConvertObjectToPointer
' Author    : KRISH J
' Date      : 20/12/2017
' Purpose   : Converts an object to a pointer in memory
'---------------------------------------------------------------------------------------
'
    Dim lngThisPointer As Long

    RtlMoveMemory lngThisPointer, objThisObject, POINTERSIZE
    ConvertObjectToPointer = lngThisPointer
End Function

Public Function ConvertFontToCm(stringLength As Long) As Double
'---------------------------------------------------------------------------------------
' Procedure : ConvertFontToCm
' Author    : KRISH J
' Purpose   : Returns 10x smaller unit than given lenght
' Returns   :
'---------------------------------------------------------------------------------------

    ConvertFontToCm = stringLength / 10
End Function

Public Function ConvertCmToTwips(ValueInCm As Double) As Double
'---------------------------------------------------------------------------------------
' Procedure : ConvertCmToTwips
' Author    : KRISH J
' Purpose   : Converts cm units into twips. 567 twips in one cm. 1440 twips in one inch
' Returns   : Twips value in double
'---------------------------------------------------------------------------------------
'

    ConvertCmToTwips = 567 * ValueInCm
End Function
