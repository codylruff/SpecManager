Attribute VB_Name = "Utils"
Option Explicit
Option Compare Text
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
Declare Function CoCreateGuid Lib "ole32" (ByRef GUID As Byte) As Long
' -------------------------------------------------

Public Function GenerateGUID() As String
    Dim ID(0 To 15) As Byte
    Dim N As Long
    Dim GUID As String
    Dim Res As Long
    Res = CoCreateGuid(ID(0))

    For N = 0 To 15
        GUID = GUID & IIf(ID(N) < 16, "0", "") & Hex$(ID(N))
        If Len(GUID) = 8 Or Len(GUID) = 13 Or Len(GUID) = 18 Or Len(GUID) = 23 Then
            GUID = GUID & "-"
        End If
    Next N
    GenerateGUID = GUID
End Function

Public Sub ColumnToRow(rng As Range)
' Take a column range and convert it to a row range starting in the first cell.
    Dim cell As Range

    For Each cell In rng.Cells
        Debug.Print cell.value
        If Not cell.row = 1 Then
            rng.Worksheet.Cells(1, cell.row).value = cell.value
            cell.value = Null
        End If
    Next cell

End Sub

Function ArrayContains(Arr As Variant, item As Variant) As Boolean
' Checks for an item within the given array and returns true or false.
    Dim i As Long
    For i = 0 To UBound(Arr)
        If Arr(i) = item Then
            ArrayContains = True
            Exit Function
        End If
    Next i
End Function

Public Function GetFiles(Optional dir_path As String, Optional pfilters As Variant) As Variant
' Given a dir return an array of file names
    Dim Arr() As String
    Dim file As Variant
    Dim i As Long
    Dim result As Integer
    Dim fDialog As FileDialog
    On Error Resume Next
    'IMPORTANT!
    Set fDialog = Application.FileDialog(3)
    fDialog.AllowMultiSelect = True
    
    'Optional FileDialog properties
    fDialog.title = "Select files"
    If dir_path = nullstr Then dir_path = "C:\Users\cruff\documents\projects\source"
    fDialog.InitialFileName = dir_path

    'Optional: Add filters
    fDialog.Filters.Clear
    If Not IsMissing(pfilters) Then
        Dim filters_string As String
        For i = LBound(pfilters) To UBound(pfilters)
            If i = UBound(pfilters) Then
                filters_string = filters_string & "*." & pfilters(i)
            Else
                filters_string = filters_string & "*." & pfilters(i) & "; "
            End If
        Next i
        If i < 2 Then filters_string = Replace(filters_string, ";", nullstr)
        fDialog.Filters.Add "Custom", filters_string
    End If
    'Show the dialog. -1 means success!
    If fDialog.show = -1 Then
        i = 0
        ReDim Arr(CLng(fDialog.SelectedItems.Count))
        For Each file In fDialog.SelectedItems
            Arr(i) = CStr(file)
            i = i + 1
        Next file
    End If
    GetFiles = Arr

End Function

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
    src_sht.Range(src_address).Copy Destination:=dest_sht.Range(dest_address)
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
        clean_string = Replace(clean_string, find_strings(i), nullstr)
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
    On Error Resume Next
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
        RemoveWhiteSpace = .Replace(target, nullstr)
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
    err.Raise REGEX_ERR
    Logger.Log "RegEx Error: ConvertToCamelCase", DebugLog
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
    Logger.Log "RegEx Error: SplitCamelCase", DebugLog
    Resume Error_Handler_Exit
End Function

Function GetLine(ParamArray var() As Variant) As String
'test
    Const Padding = 25
    Dim i As Integer
    Dim s As String
    s = nullstr
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
    s = nullstr
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
    If Not GUI.PerformanceModeEnabled Then App.PerformanceMode (True)

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
    If GUI.PerformanceModeEnabled Then App.PerformanceMode (False)
End Function

Function RemoveSheet(ws As Worksheet) As Boolean

    On Error GoTo ErrorHandler
    ' Turn on Performance Mode
    If Not GUI.PerformanceModeEnabled Then App.PerformanceMode (True)
    ws.Delete
    ' Turn off Performance Mode
    If GUI.PerformanceModeEnabled Then App.PerformanceMode (False)
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
                If ctl.value = nullstr Then
                    MsgBox "All boxes must be filed.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
            Case "ComboBox"
                If ctl.value = nullstr Then
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

Sub UpdateTable(shtName As String, tblName As String, header As String, val)
'Adds an entry at the bottom of specified column header.
    Dim rng As Range
    Set rng = Sheets(shtName).Range(tblName & "[" & header & "]")
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

Public Function ArrayLength(Arr As Variant) As Long
'test
    ArrayLength = UBound(Arr) - LBound(Arr) + 1
End Function

Sub ChangeActivePrinter()
'
' ChangeActivePrinter Macro

    Application.Dialogs(xlDialogPrinterSetup).show
    Logger.Log "Setting default printer for Spec Manager : " & Application.ActivePrinter
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
    If Not GUI.PerformanceModeEnabled Then App.PerformanceMode (True)

    For Each xWb In Application.Workbooks
        If Not xWb.ReadOnly And Windows(xWb.Name).Visible Then
            xWb.Save
        End If
    Next
    ' Turn off Performance Mode
    If GUI.PerformanceModeEnabled Then App.PerformanceMode (False)
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
           Optional header As Boolean = True, Optional footer As Boolean = True)
' Clears the contents of header and footer (optionally select one or the other)
    
    ' Clear Header
    If header Then
        With ws.PageSetup
            .LeftHeader = nullstr
            .CenterHeader = nullstr
            .RightHeader = nullstr
        End With
    End If
    ' Clear Footer
    If footer Then
        With ws.PageSetup
            .LeftFooter = nullstr
            .CenterFooter = nullstr
            .RightFooter = nullstr
        End With
    End If

End Sub

Public Function GetNames(wb As Workbook, Optional ws As String) As Variant
' Returns an array of names for the given workbook/worksheet
    Dim Arr() As String
    Dim arr_len As Long
    Dim nm As Variant
    Dim i As Long
    i = 0
    On Error Resume Next
    If ws = nullstr Then
        arr_len = wb.Names.Count
        ReDim Arr(arr_len, 1)
        For Each nm In wb.Names
            Arr(i, 0) = Split(nm.Name, "!")(1)
            Arr(i, 1) = wb.Range(nm.Name).value
            i = i + 1
        Next nm
    Else
        arr_len = wb.Sheets(ws).Names.Count
        ReDim Arr(arr_len, 1)
        For Each nm In wb.Sheets(ws).Names
            Arr(i, 0) = Split(nm.Name, "!")(1)
            Arr(i, 1) = CStr(wb.Sheets(ws).Range(nm.Name).value)
            i = i + 1
        Next nm
    End If
    On Error GoTo 0
    GetNames = Arr

End Function

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modSortACollectionOrDictionary --> Added this to Utils for compactness sake (cody.ruff.engr@gmail.com)
' By Chip Pearson www.cpearson.com, chip@cpearson.com
'
' This module contains functions for sorting and manipulating
' Collection and Dictionary objects. This module contains the following
' proceudres:
'
'       Set ArrayToCollection
'       ArrayToDictionary
'       CollectionToArray
'       CollectionToDictionary
'       CollectionToRange
'       CreateDictionaryKeyFromCollectionItem
'       DictionaryToArray
'       DictionaryToCollection
'       DictionaryToRange
'       KeyExistsInCollection
'       RangeToDictionary
'       RangeToCollection
'       SortCollection
'       SortDictionary
'
' NOTE: converted to latebinding
'
' NOTE: This module requires the modArraySupport code module, which is available for download
' at http://www.cpearson.com/excel/VBAArrays.htm and the modQSortInPlace module, which is
' available for download at http://www.cpearson.com/excel/qsort.htm.
' These modules are included in the example workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Function ArrayToCollection(Arr As Variant, ByRef coll As VBA.Collection) As Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set ArrayToCollection
' This function converts an array to a Collection. Arr may be either a 1-dimensional
' arrary or a two-dimensional array. If Arr is a 1-dimensional array, each element
' of the array is added to Coll without a key. If Arr is a 2-dimensional array,
' the first column is assumed to the be Item to be added, and the second column
' is assumed to be the Key for that item.
' Items are added to the Coll collection. Existing contents are preserved.
' This function returns True if successful, or False if an error occurs.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim KeyVal As String

''''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    Set ArrayToCollection = Nothing
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Work with either a 1-dimensional
' or 2-dimensional array. Any other
' number of dimensions will cause
' a error. Use On Error to
' trap for errors (most likely a
' duplicate key error).
'''''''''''''''''''''''''''''''''''
On Error GoTo ErrH:
Select Case NumberOfArrayDimensions(Arr:=Arr)
    Case 0
        '''''''''''''''''''''''''''''''
        ' Unallocated array. Exit with
        ' error.
        '''''''''''''''''''''''''''''''
        Set ArrayToCollection = Nothing
        Exit Function
        
    Case 1
        ''''''''''''''''''''''''''''''
        ' Arr is a single dimensional
        ' array. Load the elements of
        ' the array without keys.
        ''''''''''''''''''''''''''''''
        For Ndx = LBound(Arr) To UBound(Arr)
            coll.Add item:=Arr(Ndx)
        Next Ndx
    
    Case 2
        '''''''''''''''''''''''''''''
        ' Arr is a two-dimensional
        ' array. The first column
        ' is the Item and the second
        ' column is the Key.
        '''''''''''''''''''''''''''''
        For Ndx = LBound(Arr, 1) To UBound(Arr, 1)
            KeyVal = Arr(Ndx, 1)
            If Trim(KeyVal) = nullstr Then
                '''''''''''''''''''''''''''''''''
                ' Key is empty. Add to collection
                ' without a key.
                '''''''''''''''''''''''''''''''''
                coll.Add item:=Arr(Ndx, 1)
            Else
                '''''''''''''''''''''''''''''''''
                ' Key is not empty. Add with key.
                '''''''''''''''''''''''''''''''''
                coll.Add item:=Arr(Ndx, 0), Key:=KeyVal
            End If
        Next Ndx
    
    Case Else
        '''''''''''''''''''''''''''''
        ' The array has 3 or more
        ' dimensions. Return an
        ' error.
        '''''''''''''''''''''''''''''
        Set ArrayToCollection = Nothing
        Exit Function

End Select

Set ArrayToCollection = coll
Exit Function

ErrH:
    ''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' a duplicate key error. Return
    ' False.
    ''''''''''''''''''''''''''''''''
    Set ArrayToCollection = Nothing


End Function

Public Function ArrayToDictionary(Arr As Variant, dict As Object) As Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ArrayToDictionary
' This function loads the contents of a two dimensional array into the Dict dictionary
' object. Arr must be two dimensional. The first column is the Item to add to the Dict
' dictionary, and the second column is the Key value of the Item. The existing items
' in the dictionary are left intact.
' The function returns True if successful, or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim ItemVar As Variant
Dim KeyVal As String

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    Set ArrayToDictionary = Nothing
    Exit Function
End If

'''''''''''''''''''''''''''''''
' Ensure Arr is two dimensional
'''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then
    Set ArrayToDictionary = Nothing
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''''
' Loop through the arary and
' add the items to the Dictionary.
'''''''''''''''''''''''''''''''''''
On Error GoTo ErrH:
For Ndx = LBound(Arr, 1) To UBound(Arr, 1)
    dict.Add Key:=Arr(Ndx, LBound(Arr, 2) + 1), item:=Arr(Ndx, LBound(Arr, 2))
Next Ndx
    
'''''''''''''''''
' Return Success.
'''''''''''''''''
Set ArrayToDictionary = dict
Exit Function

ErrH:
Set ArrayToDictionary = Nothing

End Function

Public Function CollectionToArray(coll As VBA.Collection, Arr As Variant) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToArray
' This function converts a collection object to a single dimensional array.
' The elements of Collection may be any type of data except User Defined Types.
' The procedure will populate the array Arr with the elements of the collection.
' Only the collection items, not the keys, are stored in Arr. The function returns
' True if the the Collection was successfully converted to an array, or False
' if an error occcurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V As Variant
Dim Ndx As Long

''''''''''''''''''''''''''''''
' Ensure Coll is not Nothing.
''''''''''''''''''''''''''''''
If coll Is Nothing Then
    CollectionToArray = Null
    Exit Function
End If

''''''''''''''''''''''''''''''
' Ensure Arr is an array and
' is dynamic.
''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    CollectionToArray = Null
    Exit Function
End If
If IsArrayDynamic(Arr:=Arr) = False Then
    CollectionToArray = Null
    Exit Function
End If

''''''''''''''''''''''''''''
' Ensure Coll has at least
' one item.
''''''''''''''''''''''''''''
If coll.Count < 1 Then
    CollectionToArray = Null
    Exit Function
End If
    
''''''''''''''''''''''''''''''
' Redim Arr to the number of
' elements in the collection.
'''''''''''''''''''''''''''''
ReDim Arr(1 To coll.Count)
'''''''''''''''''''''''''''''
' Loop through the colletcion
' and add the elements of
' Collection to Arr.
'''''''''''''''''''''''''''''
For Ndx = 1 To coll.Count
    If IsObject(coll(Ndx)) = True Then
        Set Arr(Ndx) = coll(Ndx)
    Else
        Arr(Ndx) = coll(Ndx)
    End If
Next Ndx

CollectionToArray = Arr

End Function


Public Function CollectionToDictionary(coll As VBA.Collection, _
    dict As Object) As Object
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToDictionary
'
' This function converts a Collection Objct to a
' Dictionary object. This code requires a reference
' the Microsoft Scripting RunTime Library.
'
' It calls a private procedure named
' CreateDictionaryKeyFromCollectionItem that you supply
' to create a Dictionary Key from an Item in the Collection.
' This must return a String value that will be unique within
' the Dictionary.
'
' If an error occurs (e.g., a Key value returned by
' CreateDictionaryKeyFromCollectionItem already exists
' in the Dictionary object), Dictionary is set to Nothing.
' The function returns True if the conversion from Collection
' to Dictionary was successful, or False if an error occurred.
' If it returns False, the Dictionary is set to Nothing.
'
' The code destroys the existing contents of Dict
' and replaces them with the new elements. The Coll
' Collection is left intact with no changes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim ItemKey As String
Dim ItemVar As Variant

''''''''''''''''''''''''''''''''''''''''''''
' Ensure Coll is not Nothing.
''''''''''''''''''''''''''''''''''''''''''''
If (coll Is Nothing) Then
    Set CollectionToDictionary = Nothing
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''
' Reset Dict to a new, empty Dictionary
''''''''''''''''''''''''''''''''''''''''''''
Set dict = Nothing
Set dict = CreateObject("Scripting.Dictionary")
'''''''''''''''''''''''''''''''''''''''''''
' Ensure we have at least one element in
' the collection object.
'''''''''''''''''''''''''''''''''''''''''''
If coll.Count = 0 Then
    Set dict = Nothing
    Set CollectionToDictionary = Nothing
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''''''''''''
' Loop through the collection and convert
' each item in the collection to an item
' for the dictionary. Call
' CreateDictionaryKeyFromCollectionItem
' to get the Key to be used in the Dictionary
' item.
'''''''''''''''''''''''''''''''''''''''''''
For Ndx = 1 To coll.Count
    '''''''''''''''''''''''''''''''''''''''
    ' Coll may contain object variables.
    ' Test for this condition and set
    ' ItemVar appropriately.
    '''''''''''''''''''''''''''''''''''''''
    If IsObject(coll(Ndx)) = True Then
        Set ItemVar = coll(Ndx)
    Else
        ItemVar = coll(Ndx)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call the user-supplied CreateDictionaryKeyFromCollectionItem
    ' function to get the Key to be used in the Dictionary.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ItemKey = CreateDictionaryKeyFromCollectionItem(dict:=dict, item:=ItemVar)
    ''''''''''''''''''''''''''''''''
    ' ItemKey must not be spaces or
    ' an empty string.
    ''''''''''''''''''''''''''''''''
    If Trim(ItemKey) = nullstr Then
        Set CollectionToDictionary = Nothing
        Exit Function
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' See if ItemKey already exists in the Dictionary.
    ' If so, return False. You can't have duplicate keys.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If dict.Exists(Key:=ItemKey) = True Then
        Set dict = Nothing
        Set CollectionToDictionary = Nothing
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ItemKey does not exist in Dict, so add ItemVar to
    ' Dict with a key of ItemKey.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dict.Add Key:=ItemKey, item:=ItemVar
Next Ndx
Set CollectionToDictionary = dict

End Function

Private Function CreateDictionaryKeyFromCollectionItem( _
    dict As Object, _
    item As Variant) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateDictionaryKeyFromCollectionItem
' This function is called by CollectionToDictionary to create
' a Key for a Dictionary item that is take from a Collection
' item. The collection item is passed in the Item parameter.
' It is up to you to create a unique key based on the
' Item parameter.
' Dict is the Dictionary for which the result of this function
' will be used as a Key, and Item is the item of the
' Dictionary for which this procedure is creating a Key.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ItemKey As String
''''''''''''''''''''''''''''''''''''''''''
' Your code to set ItemKey to the
' appropriate string value. ItemKey
' must not be all spaces or nullstr.
''''''''''''''''''''''''''''''''''''''''''

CreateDictionaryKeyFromCollectionItem = ItemKey
End Function


Public Function CollectionToRange(coll As VBA.Collection, StartCells As Range) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToRange
' This procedure writes the contents of a Collection Coll to a range of cells starting
' in StartCells. If StartCells is a single cell, the contents of Collection are
' written downward in a single column starting in StartCell. If StartCell is
' two cells, the Collection is written in the same orientation (down a column or
' across a row) as StartCells. If StartCells is more than two cells, ONLY those
' cells will be written to, moving across then down. StartCells must be a single
' area range.
'
' If an item in Coll is an object, it is skipped.
'
' The function returns True if successful or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DestRng As Range
Dim V As Variant
Dim Ndx As Long

'''''''''''''''''''''''''''''''''''''
' Ensure parameters are not Nothing.
'''''''''''''''''''''''''''''''''''''
If (coll Is Nothing) Or (StartCells Is Nothing) Then
    CollectionToRange = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure StartCells is a single area.
'''''''''''''''''''''''''''''''''''''
If StartCells.Areas.Count > 1 Then
    CollectionToRange = False
    Exit Function
End If

If StartCells.Cells.Count = 1 Then
    '''''''''''''''''''''''''''''''''''''
    ' StartCells is one cell. Write out
    ' the collection moving downwards.
    '''''''''''''''''''''''''''''''''''''
    Set DestRng = StartCells
    For Each V In coll
        If IsObject(V) = False Then
            DestRng.value = V
            If DestRng.row < DestRng.Parent.Rows.Count Then
                Set DestRng = DestRng(2, 1)
            Else
                CollectionToRange = False
                Exit Function
            End If
                
        End If
    Next V
    CollectionToRange = True
    Exit Function
End If

If StartCells.Cells.Count = 2 Then
    ''''''''''''''''''''''''''''''''''
    ' Test the orientation of the two
    ' cells in StartCells.
    ''''''''''''''''''''''''''''''''''
    If StartCells.Rows.Count = 1 Then
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' across the row.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In coll
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Column < StartCells.Parent.Columns.Count Then
                    Set DestRng = DestRng(1, 2)
                Else
                    CollectionToRange = False
                    Exit Function
                End If
            End If
        Next V
        CollectionToRange = True
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' down the column.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In coll
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.row < StartCells.Parent.Rows.Count Then
                    Set DestRng = DestRng(2, 1)
                Else
                    CollectionToRange = False
                    Exit Function
                End If
            End If
        Next V
        CollectionToRange = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''''''''''
' Write the collection only into
' Cells of StartCells.
'''''''''''''''''''''''''''''''''''''
For Ndx = 1 To StartCells.Cells.Count
    If Ndx <= coll.Count Then
        V = coll(Ndx)
        If IsObject(V) = False Then
            StartCells.Cells(Ndx).value = V
        End If
    End If
Next Ndx

CollectionToRange = True


End Function


Public Function DictionaryToRange(dict As Object, StartCells As Range) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToRange
' This procedure writes the contents of a Dictionary Dict to a range of cells starting
' in StartCells. If StartCells is a single cell, the contents of Dict are
' written downward in a single column starting in StartCell. If StartCell is
' two cells, the Dictionary is written in the same orientation (down a column or
' across a row) as StartCells. If StartCells is more than two cells, ONLY those
' cells will be written to, moving across then down. StartCells must be a single
' area range.
'
' If an item in Dict is an object, it is skipped.
'
' The function returns True if successful or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DestRng As Range
Dim V As Variant
Dim Ndx As Long

'''''''''''''''''''''''''''''''''''''
' Ensure parameters are not Nothing.
'''''''''''''''''''''''''''''''''''''
If (dict Is Nothing) Or (StartCells Is Nothing) Then
    DictionaryToRange = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure StartCells is a single area.
'''''''''''''''''''''''''''''''''''''
If StartCells.Areas.Count > 1 Then
    DictionaryToRange = False
    Exit Function
End If

If StartCells.Cells.Count = 1 Then
    '''''''''''''''''''''''''''''''''''''
    ' StartCells is one cell. Write out
    ' the collection moving downwards.
    '''''''''''''''''''''''''''''''''''''
    Set DestRng = StartCells
    For Each V In dict.Items
        If IsObject(V) = False Then
            DestRng.value = V
            If DestRng.row < DestRng.Parent.Rows.Count Then
                Set DestRng = DestRng(2, 1)
            Else
                DictionaryToRange = False
                Exit Function
            End If
                
        End If
    Next V
    DictionaryToRange = True
    Exit Function
End If

If StartCells.Cells.Count = 2 Then
    ''''''''''''''''''''''''''''''''''
    ' Test the orientation of the two
    ' cells in StartCells.
    ''''''''''''''''''''''''''''''''''
    If StartCells.Rows.Count = 1 Then
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' across the row.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In dict.Items
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Column < StartCells.Parent.Columns.Count Then
                    Set DestRng = DestRng(1, 2)
                Else
                    DictionaryToRange = False
                    Exit Function
                End If
            End If
        Next V
        DictionaryToRange = True
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''
        ' Write out the Dictionary moving
        ' down the column.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In dict.Items
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.row < StartCells.Parent.Rows.Count Then
                    Set DestRng = DestRng(2, 1)
                Else
                    DictionaryToRange = False
                    Exit Function
                End If
            End If
        Next V
        DictionaryToRange = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''''''''''
' Write the Dictionary only into
' Cells of StartCells.
'''''''''''''''''''''''''''''''''''''
For Ndx = 1 To StartCells.Cells.Count
    If Ndx <= dict.Count Then
        V = dict.Items(Ndx - 1)
        If IsObject(V) = False Then
            StartCells.Cells(Ndx).value = V
        End If
    End If
Next Ndx

DictionaryToRange = True


End Function

Public Function DictionaryToArray(dict As Object, Arr As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToArray
' This creates a 0-based, 2-dimensional array Arr from a Dictionary object.  Each
' row of the array is one element of the Dictionary. The first column of the array is the
' Key of the dictionary item, and the second column is the Key of the item in the
' dictionary. Arr MUST be an dynamic array of Variants, e.g.,
'       Dim Arr() As Variant
' The VarType of Arr is tested, and if it does not equal 8204 (vbArray + vbVariant) an
' error occurs.
'
' The existing content of Arr is destroyed. The function returns True if successsful
' or False if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''
' Ensure that Arr is an array
' of Variants.
'''''''''''''''''''''''''''''
If VarType(Arr) <> (vbArray + vbVariant) Then
    DictionaryToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Arr is a dynamic array.
''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=Arr) = False Then
    DictionaryToArray = False
    Exit Function
End If
   
'''''''''''''''''''''''''''''
' Ensure Dict is not nothing.
'''''''''''''''''''''''''''''
If dict Is Nothing Then
    DictionaryToArray = False
    Exit Function
End If
    
'''''''''''''''''''''''''''
' Ensure that Dict contains
' at least one entry.
'''''''''''''''''''''''''''
If dict.Count = 0 Then
    DictionaryToArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''
' Redim the Arr variable.
'''''''''''''''''''''''''''''
ReDim Arr(0 To dict.Count - 1, 0 To 1)

For Ndx = 0 To dict.Count - 1
    Arr(Ndx, 0) = dict.keys(Ndx)
    '''''''''''''''''''''''''''''''''''''''''
    ' Test to see if the item in the Dict is
    ' an object. If so, use Set.
    '''''''''''''''''''''''''''''''''''''''''
    If IsObject(dict.Items(Ndx)) = True Then
        Set Arr(Ndx, 1) = dict.Items(Ndx)
    Else
        Arr(Ndx, 1) = dict.Items(Ndx)
    End If

Next Ndx

'''''''''''''''''
' Return success.
'''''''''''''''''
DictionaryToArray = True

End Function

Public Function DictionaryToCollection(dict As Object, coll As VBA.Collection, _
    Optional PreserveColl As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToCollection
' This procedure converts an existing Dictionary to a new Collection object. Keys from
' the Dictionary are used as the keys for the Collection. This function returns True
' if successful, or False if an error occurred. The contents of Dict are not modified.
' If PreserveColl is omitted or False, the existing contents of the Coll collection are
' destroyed. If PreserveColl is True, the existing contents of Coll are preserved.
' If PreserveColl is true, then the possibility exists that we will run into duplicate
' key values for the Collection. If StopOnDuplicateKey is omitted or false, this error
' is ignored, but the item from the Dict Dictionary will not be added to Coll Collection.
' If StopOnDuplicateKey is True, the procedure will terminate, and not all of the items in
' the Dict Dictionary will have copied to the Coll Collection. The Coll Collection will
' be in an indeterminant state.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim ItemVar As Variant
Dim KeyVal As String

''''''''''''''''''''''''''''''''
' Ensure Dict is not Nothing
''''''''''''''''''''''''''''''''
If dict Is Nothing Then
    DictionaryToCollection = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' If PreseveColl is omitted or
' False, destroy the existing
' Coll Collection.
'''''''''''''''''''''''''''''''''
If PreserveColl = False Then
    Set coll = Nothing
    Set coll = New VBA.Collection
End If

'''''''''''''''''''''''''''''''''
' Loop through the Dictionary
' and transfer the data to
' the Collection.
'''''''''''''''''''''''''''''''''
On Error Resume Next
For Ndx = 0 To dict.Count - 1
    If IsObject(dict.Items(Ndx)) = True Then
        Set ItemVar = dict.Items(Ndx)
    Else
        ItemVar = dict.Items(Ndx)
    End If
    KeyVal = dict.keys(Ndx)
    err.Clear
    coll.Add item:=ItemVar, Key:=KeyVal
    If err.Number <> 0 Then
        If StopOnDuplicateKey = True Then
            DictionaryToCollection = False
            Exit Function
        End If
    End If
Next Ndx
DictionaryToCollection = True
End Function

Public Function KeyExistsInCollection(coll As VBA.Collection, KeyName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' KeyExistsInCollection
' This function determines if the key KeyName exists in the collection Coll. The
' function returns True if an item with the specified key exists, or False if
' the key does not exist. If Coll is Nothing, the result is False.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V As Variant
    
    If coll Is Nothing Then
        KeyExistsInCollection = False
        Exit Function
    End If
    
    On Error Resume Next
    V = coll(KeyName)
    Select Case err.Number
        Case 0
            KeyExistsInCollection = True
        Case 5, 438
            '''''''''''''''''''''''''''''''''''''
            ' We'll get one of these error if
            ' Coll(KeyName) is an object variable.
            ' SET V to the item and retest the
            ' error code.
            ''''''''''''''''''''''''''''''''''''''
            err.Clear
            Set V = coll(KeyName)
            Select Case err.Number
                Case 0
                    KeyExistsInCollection = True
                Case Else
                    KeyExistsInCollection = False
            End Select
        Case Else
            '''''''''''''''''''''''''''''
            ' Error. Key does not exist.
            '''''''''''''''''''''''''''''
            KeyExistsInCollection = False
    End Select

End Function




Public Sub SortCollection(ByRef coll As VBA.Collection, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortCollection
' This sorts a collection by its items. It does not preserve
' the keys associated with the item. This limitation is due
' to the fact that Key is a write-only property. If you need
' sort by or preserve Keys, you should be using a Dictionary
' object rather than a Collection object. You can convert
' a Collection to a Dictionary with the function
' CollectionToDictionary. This procedure requires that you
' provide a funtion called CreateDictionaryKeyFromCollectionItem
' that creates a Dictionary Key from each Item in the
' Collection.
'
' By default, string comparison are case-INSENSITIVE (e.g.,
' "a" = "A"). To sort case-SENSITIVE (e.g., "a" <> "A"), set
' the CompareMode parameter to vbBinaryCompare.
' By default, the items in Coll are sorted in ascending order.
' You can sort in descending order by setting the Descending
' parameter to True.
'
' The items in the collection must be simple data types.
' Objects, Arrays, and UserDefinedTypes are not allowed.
'
' Note: This procedure requires the
' QSortInPlace function, which is described and available for
' download at www.cpearson.com/excel/qsort.htm .
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Arr() As Variant
Dim Ndx As Long

'''''''''''''''''''''''''''''''''''''
' Ensure that Coll is not Nothing.
'''''''''''''''''''''''''''''''''''''
If coll Is Nothing Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''
' Ensure CompareMode is valid value.
''''''''''''''''''''''''''''''''''''''
Select Case CompareMode
    Case vbTextCompare, vbBinaryCompare
    Case Else
        Exit Sub
End Select


''''''''''''''''''''''''''''''''''''''
' If the number of elements in Coll
' is 0 or 1, no sorting is required.
' Get out.
''''''''''''''''''''''''''''''''''''''
If coll.Count <= 1 Then
    Exit Sub
End If
ReDim Arr(1 To coll.Count)
For Ndx = 1 To coll.Count
    If IsObject(Arr(Ndx)) = True Or IsArray(Arr(Ndx)) = True Then
        Debug.Print "The items of the Collection cannot be arrays or objects."
        Exit Sub
    End If
    Arr(Ndx) = coll(Ndx)
Next Ndx
''''''''''''''''''''''''''''''''''''''''''
' Sort the elements in the array. The
' QSortInPlace function is described on
' and downloadable from:
' http://www.cpearson.com/excel/qsort.htm
''''''''''''''''''''''''''''''''''''''''''
QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, _
    Descending:=Descending, CompareMode:=vbTextCompare
''''''''''''''''''''''''''''''''''''''''''
' Now reset Coll to a new, empty colletion.
''''''''''''''''''''''''''''''''''''''''''
Set coll = Nothing
Set coll = New VBA.Collection
''''''''''''''''''''''''''''''''''''''''''
' Load the array back into the new
' collection.
'''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(Arr) To UBound(Arr)
    coll.Add item:=Arr(Ndx)
Next Ndx

End Sub

Function RangeToDictionary(KeyRange As Range, ItemRange As Range, dict As Object, _
    Optional RangeAsObject As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = True, _
    Optional ReplaceOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RangeToDictionary
' This funciton loads an existing Dictionary Dict with the keys and value from
' worksheet ranges.
' The KeyRange and ItemRange must be the same size. Each element in KeyRange
' is the Key value for the corresponding item in ItemRange.
'
' If RangeAsObject is omitted of False, the Items added to the Dictionary are
' the values in the cells of ItemRange. If RangeAsObject is True, the cells
' are added as objects to the Dictionary.
'
' If a duplicate key is encountered when adding an item to Dict, the code
' will do one of the following:
'   If StopOnDuplicateKey is omitted or True, the funcion stops processing
'   and returns False. Items added to the Dictionary before the duplicate key
'   was encountered remain in the Dictionary.
'
'   If StopOnDuplicateKey is False, then if ReplaceOnDuplicateKey is False,
'   the Item that caused the duplicate key error is not added to the Dictionary
'   but processing continues with the rest of the items in the range. If
'   ReplaceOnDuplicateKey if True, the existing item in the Dictionary is
'   deleted and replaced with the new item.
'
' If Dict is Nothing, it will be created as a new Dictionary.
'
' The function returns True if all items were added to the dictionary, or False
' if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim KRng As Range
Dim KeyExists As Boolean
Dim ItemNdx As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the KeyRange and ItemRange variables are not
' Nothing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange Is Nothing) Or (ItemRange Is Nothing) Then
    RangeToDictionary = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure KeyRange and ItemRange as the same size.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange.Rows.Count <> ItemRange.Rows.Count) Or _
    (KeyRange.Columns.Count <> ItemRange.Columns.Count) Then
    RangeToDictionary = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure both KeyRange and ItemRange are single area
' ranges.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange.Areas.Count > 1) Or (ItemRange.Areas.Count > 1) Then
    RangeToDictionary = False
    Exit Function
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If Dict is Nothing, create a new dictionary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If dict Is Nothing Then
    Set dict = CreateObject("Scripting.Dictionary")
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through KeyRange, testing whether the Key exists
' and adding items to the Dictionary.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each KRng In KeyRange.Cells
    ItemNdx = ItemNdx + 1
    KeyExists = dict.Exists(Key:=KRng.text)
    If KeyExists = True Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The key already exists in the Dictionary.
        ' Determine what to do.
        '''''''''''''''''''''''''''''''''''''''''''
        If StopOnDuplicateKey = True Then
            RangeToDictionary = False
            Exit Function
        Else
            ''''''''''''''''''''''''''''''''''''''
            ' Do nothing here. Test the value of
            ' ReplaceOnDuplicateKey below.
            ''''''''''''''''''''''''''''''''''''''
        End If
        '''''''''''''''''''''''''''''''''''''''''
        ' If ReplaceOnDuplicateKey is True then
        ' remove the existing entry. Otherwise,
        ' exit the function.
        '''''''''''''''''''''''''''''''''''''''''
        If ReplaceOnDuplicateKey = True Then
            dict.Remove Key:=KRng.text
            KeyExists = False
        Else
            If StopOnDuplicateKey = True Then
                RangeToDictionary = False
                Exit Function
            End If
        End If
    End If
    If KeyExists = False Then
        If RangeAsObject = True Then
            dict.Add Key:=KRng.text, item:=ItemRange.Cells(ItemNdx)
        Else
            dict.Add Key:=KRng.text, item:=ItemRange.Cells(ItemNdx).text
        End If
    End If
Next KRng

'''''''''''''''''
' Return Success.
'''''''''''''''''
RangeToDictionary = True

End Function

Function RangeToCollection(KeyRange As Range, ItemRange As Range, coll As VBA.Collection, _
    Optional RangeAsObject As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = True, _
    Optional ReplaceOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RangeToCollection
' This function load an existing Collection Coll with items from worksheet
' ranges.
'
' The KeyRange and ItemRange must be the same size. Each element in KeyRange
' is the Key value for the corresponding item in ItemRange.
'
' KeyRange may be Nothing. In this case, the items in ItemRange are added to
' the Collection Coll without keys.
'
' If RangeAsObject is omitted of False, the Items added to the Collection are
' the values in the cells of ItemRange. If RangeAsObject is True, the cells
' are added as objects to the Collection.
'
' If a duplicate key is encountered when adding an item to Coll, the code
' will do one of the following:
'   If StopOnDuplicateKey is omitted or True, the funcion stops processing
'   and returns False. Items added to the Collection before the duplicate key
'   was encountered remain in the Collection.
'
'   If StopOnDuplicateKey is False, then if ReplaceOnDuplicateKey is False,
'   the Item that caused the duplicate key error is not added to the Collection
'   but processing continues with the rest of the items in the range. If
'   ReplaceOnDuplicateKey if True, the existing item in the Collection is
'   deleted and replaced with the new item.
'
' If Coll is Nothing, it will be created as a new Collection.
'
' The function returns True if all items were added to the Collection, or False
' if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim IRng As Range
Dim KeyExists As Boolean
Dim KeyNdx As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the KeyRange and ItemRange variables are not
' Nothing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (ItemRange Is Nothing) Then
    RangeToCollection = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure KeyRange and ItemRange as the same size.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not KeyRange Is Nothing Then
    If (KeyRange.Rows.Count <> ItemRange.Rows.Count) Or _
        (KeyRange.Columns.Count <> ItemRange.Columns.Count) Then
        RangeToCollection = False
        Exit Function
    End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure both KeyRange and ItemRange are single area
' ranges.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ItemRange.Areas.Count > 1 Then
    RangeToCollection = False
    Exit Function
End If

If Not KeyRange Is Nothing Then
    If KeyRange.Areas.Count > 1 Then
        RangeToCollection = False
        Exit Function
    End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If Coll is Nothing, create a new Collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If coll Is Nothing Then
    Set coll = New VBA.Collection
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through ItemRange, testing whether the Key exists
' and adding items to the Collection.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each IRng In ItemRange.Cells
    KeyNdx = KeyNdx + 1
    If KeyRange Is Nothing Then
        KeyExists = False
    Else
        KeyExists = KeyExistsInCollection(coll:=coll, KeyName:=KeyRange.Cells(KeyNdx))
    End If
    
    If KeyExists = True Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The key already exists in the Collection.
        ' Determine what to do.
        '''''''''''''''''''''''''''''''''''''''''''
        If StopOnDuplicateKey = True Then
            RangeToCollection = False
            Exit Function
        Else
            ''''''''''''''''''''''''''''''''''''''
            ' Do nothing here. Test the value of
            ' ReplaceOnDuplicateKey below.
            ''''''''''''''''''''''''''''''''''''''
        End If
        '''''''''''''''''''''''''''''''''''''''''
        ' If ReplaceOnDuplicateKey is True then
        ' remove the existing entry. Otherwise,
        ' exit the function.
        '''''''''''''''''''''''''''''''''''''''''
        If ReplaceOnDuplicateKey = True Then
            coll.Remove KeyRange.Cells(KeyNdx)
            KeyExists = False
        Else
            If StopOnDuplicateKey = True Then
                RangeToCollection = False
                Exit Function
            End If
        End If
    End If
    If KeyExists = False Then
        '''''''''''''''''''''''''''''''
        ' Check KeyRange  to see if
        ' we're adding with Keys.
        '''''''''''''''''''''''''''''''
        If Not KeyRange Is Nothing Then
            '''''''''''''''''''''''''
            ' Add with key.
            '''''''''''''''''''''''''
            If RangeAsObject = True Then
                coll.Add item:=IRng, Key:=KeyRange.Cells(KeyNdx)
            Else
                coll.Add item:=IRng.text, Key:=KeyRange.Cells(KeyNdx)
            End If
        Else
            '''''''''''''''''''''
            ' Add without key.
            If RangeAsObject = True Then
                coll.Add item:=IRng
            Else
                coll.Add item:=IRng.text
            End If
            '''''''''''''''''''''
            
        End If
    End If
Next IRng

'''''''''''''''''
' Return Success.
'''''''''''''''''
RangeToCollection = True

End Function


Public Sub SortDictionary(dict As Object, _
    SortByKey As Boolean, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortDictionary
' This sorts a Dictionary object. If SortByKey is False, the
' the sort is done based on the Items of the Dictionary, and
' these items must be simple data types. They may not be
' Object, Arrays, or User-Defined Types. If SortByKey is True,
' the Dictionary is sorted by Key value, and the Items in the
' Dictionary may be Object as well as simple variables.
'
' If sort by key is True, all element of the Dictionary
' must have a non-blank Key value. If Key is nullstr
' the procedure will terminate.
'
' By defualt, sorting is done in Ascending order. You can
' sort by Descending order by setting the Descending parameter
' to True.
'
' By default, text comparisons are done case-INSENSITIVE (e.g.,
' "a" = "A"). To use case-SENSITIVE comparisons (e.g., "a" <> "A")
' set CompareMode to vbBinaryCompare.
'
' Note: This procedure requires the
' QSortInPlace function, which is described and available for
' download at www.cpearson.com/excel/qsort.htm .
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim KeyValue As String
Dim ItemValue As Variant
Dim Arr() As Variant
Dim KeyArr() As String
Dim VTypes() As VbVarType


Dim V As Variant
Dim SplitArr As Variant

Dim TempDict As Object
'''''''''''''''''''''''''''''
' Ensure Dict is not Nothing.
'''''''''''''''''''''''''''''
If dict Is Nothing Then
    Exit Sub
End If
''''''''''''''''''''''''''''
' If the number of elements
' in Dict is 0 or 1, no
' sorting is required.
''''''''''''''''''''''''''''
If (dict.Count = 0) Or (dict.Count = 1) Then
    Exit Sub
End If

''''''''''''''''''''''''''''
' Create a new TempDict.
''''''''''''''''''''''''''''
Set TempDict = CreateObject("Scripting.Dictionary")

If SortByKey = True Then
    ''''''''''''''''''''''''''''''''''''''''
    ' We're sorting by key. Redim the Arr
    ' to the number of elements in the
    ' Dict object, and load that array
    ' with the key names.
    ''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To dict.Count - 1)

    Dim dict_key As Variant

    Ndx = 0
    For Each dict_key In dict
        Arr(Ndx) = dict_key
        Ndx = Ndx + 1
    Next dict_key
    
    ''''''''''''''''''''''''''''''''''''''
    ' Sort the key names.
    ''''''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=CompareMode
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Load TempDict. The key value come from
    ' our sorted array of keys Arr, and the
    ' Item comes from the original Dict object.
    ''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = 0 To dict.Count - 1
        KeyValue = Arr(Ndx)
        TempDict.Add Key:=KeyValue, item:=dict.item(KeyValue)
    Next Ndx
    '''''''''''''''''''''''''''''''''
    ' Set the passed in Dict object
    ' to our TempDict object.
    '''''''''''''''''''''''''''''''''
    Set dict = TempDict
    ''''''''''''''''''''''''''''''''
    ' This is the end of processing.
    ''''''''''''''''''''''''''''''''
Else
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Here, we're sorting by items. The Items must
    ' be simple data types. They may NOT be Objects,
    ' arrays, or UserDefineTypes.
    ' First, ReDim Arr and VTypes to the number
    ' of elements in the Dict object. Arr will
    ' hold a string containing
    '   Item & vbNullChar & Key
    ' This keeps the association between the
    ' item and its key.
    '''''''''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To dict.Count - 1)
    ReDim VTypes(0 To dict.Count - 1)

    For Ndx = 0 To dict.Count - 1
        If (IsObject(dict.Items(Ndx)) = True) Or _
            (IsArray(dict.Items(Ndx)) = True) Or _
            VarType(dict.Items(Ndx)) = vbUserDefinedType Then
            Debug.Print "***** ITEM IN DICTIONARY WAS OBJECT OR ARRAY OR UDT"
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Here, we create a string containing
        '       Item & vbNullChar & Key
        ' This preserves the associate between an item and its
        ' key. Store the VarType of the Item in the VTypes
        ' array. We'll use these values later to convert
        ' back to the proper data type for Item.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Arr(Ndx) = dict.Items(Ndx) & vbNullChar & dict.keys(Ndx)
            VTypes(Ndx) = VarType(dict.Items(Ndx))
            
    Next Ndx
    ''''''''''''''''''''''''''''''''''
    ' Sort the array that contains the
    ' items of the Dictionary along
    ' with their associated keys
    ''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=vbTextCompare
    
    For Ndx = LBound(Arr) To UBound(Arr)
        '''''''''''''''''''''''''''''''''''''
        ' Loop trhogh the array of sorted
        ' Items, Split based on vbNullChar
        ' to get the Key from the element
        ' of the array Arr.
        SplitArr = Split(Arr(Ndx), vbNullChar)
        ''''''''''''''''''''''''''''''''''''''''''
        ' It may have been possible that item in
        ' the dictionary contains a vbNullChar.
        ' Therefore, use UBound to get the
        ' key value, which will necessarily
        ' be the last item of SplitArr.
        ' Then Redim Preserve SplitArr
        ' to UBound - 1 to get rid of the
        ' Key element, and use Join
        ' to reassemble to original value
        ' of the Item.
        '''''''''''''''''''''''''''''''''''''''''
        KeyValue = SplitArr(UBound(SplitArr))
        ReDim Preserve SplitArr(LBound(SplitArr) To UBound(SplitArr) - 1)
        ItemValue = Join(SplitArr, vbNullChar)
        '''''''''''''''''''''''''''''''''''''''
        ' Join will set ItemValue to a string
        ' regardless of what the original
        ' data type was. Test the VTypes(Ndx)
        ' value to convert ItemValue back to
        ' the proper data type.
        '''''''''''''''''''''''''''''''''''''''
        Select Case VTypes(Ndx)
            Case vbBoolean
                ItemValue = CBool(ItemValue)
            Case vbByte
                ItemValue = CByte(ItemValue)
            Case vbCurrency
                ItemValue = CCur(ItemValue)
            Case vbDate
                ItemValue = CDate(ItemValue)
            Case vbDecimal
                ItemValue = CDec(ItemValue)
            Case vbDouble
                ItemValue = CDbl(ItemValue)
            Case vbInteger
                ItemValue = CInt(ItemValue)
            Case vbLong
                ItemValue = CLng(ItemValue)
            Case vbSingle
                ItemValue = CLng(ItemValue)
            Case vbString
                ItemValue = CStr(ItemValue)
            Case Else
                ItemValue = ItemValue
        End Select
        ''''''''''''''''''''''''''''''''''''''
        ' Finally, add the Item and Key to
        ' our TempDict dictionary.
        ''''''''''''''''''''''''''''''''''''''
        TempDict.Add Key:=KeyValue, item:=ItemValue
    Next Ndx
End If


'''''''''''''''''''''''''''''''''
' Set the passed in Dict object
' to our TempDict object.
'''''''''''''''''''''''''''''''''
Set dict = TempDict

End Sub

Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortInPlace
'
' This function sorts the array InputArray in place -- this is, the original array in the
' calling procedure is sorted. It will work with either string data or numeric data.
' It need not sort the entire array. You can sort only part of the array by setting the LB and
' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
' UB to UBound(InputArray).
'
' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'
' The function returns TRUE if the array was successfully sorted or FALSE if an error
' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'
''''''''''''''''''''''''''''''''''''''
' MODIFYING THIS CODE:
''''''''''''''''''''''''''''''''''''''
' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
' variable. E.g.,
'       If SomethingThatCausesAnExit Then
'           RecursionLevel = RecursionLevel - 1
'           Exit Function
'       End If
'''''''''''''''''''''''''''''''''''''''
'
' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
' able to reference the InputArray in the calling procedure and the array will
' not be sorted.
'
' This function uses the following procedures. These are declared as Private procedures
' at the end of this module:
'       IsArrayAllocated
'       IsSimpleDataType
'       IsSimpleNumericType
'       QSortCompare
'       NumberOfArrayDimensions
'       ReverseArrayInPlace
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Buffer As Variant
Dim CurLow As Long
Dim CurHigh As Long
Dim CurMidpoint As Long
Dim Ndx As Long
Dim pCompareMode As VbCompareMethod

'''''''''''''''''''''''''
' Set the default result.
'''''''''''''''''''''''''
QSortInPlace = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level
' of recursion  (the function calling itself).
' RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure
' or recursively by itself. The variable is decremented
' when the procedure exits. We do the input parameter
' validation only when RecursionLevel is 1 (when
' the function is called by another function, not
' when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Static RecursionLevel As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Keep track of the recursion level -- that is, how many
' times the procedure has called itself.
' Carry out the validation routines only when this
' procedure is first called. Don't run the
' validations on a recursive call to the
' procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RecursionLevel = RecursionLevel + 1

If RecursionLevel = 1 Then
    ''''''''''''''''''''''''''''''''''
    ' Ensure InputArray is an array.
    ''''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' InputArray is not an array. Exit with a False result.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test LB and UB. If < 0 then set to LBound and UBound
    ' of the InputArray.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LB < 0 Then
        LB = LBound(InputArray)
    End If
    If UB < 0 Then
        UB = UBound(InputArray)
    End If
    
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            ''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is an empty, unallocated array."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case 1
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
        Case Else
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is multi-dimensional." & _
                      "QSortInPlace works only on single-dimensional arrays."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that InputArray is an array of simple data
    ' types, not other arrays or objects. This tests
    ' the data type of only the first element of
    ' InputArray. If InputArray is an array of Variants,
    ' subsequent data types may not be simple data types
    ' (e.g., they may be objects or other arrays), and
    ' this may cause QSortInPlace to fail on the StrComp
    ' operation.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "InputArray is not an array of simple data types."
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure that the LB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case LB
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UB
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure the UB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UB
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LB
            If NoAlerts = False Then
                MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if UB = LB, we have nothing to sort, so get out.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UB = LB Then
        QSortInPlace = True
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If

End If ' RecursionLevel = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that CompareMode is either vbBinaryCompare  or
' vbTextCompare. If it is neither, default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
    pCompareMode = CompareMode
Else
    pCompareMode = vbTextCompare
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Begin the actual sorting process.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CurLow = LB
CurHigh = UB

If LB = 0 Then
    CurMidpoint = ((LB + UB) \ 2) + 1
Else
    CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here
End If
Temp = InputArray(CurMidpoint)

Do While (CurLow <= CurHigh)
    
    Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, CompareMode:=pCompareMode) < 0
        CurLow = CurLow + 1
        If CurLow = UB Then
            Exit Do
        End If
    Loop
    
    Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
        CurHigh = CurHigh - 1
        If CurHigh = LB Then
           Exit Do
        End If
    Loop

    If (CurLow <= CurHigh) Then
        Buffer = InputArray(CurLow)
        InputArray(CurLow) = InputArray(CurHigh)
        InputArray(CurHigh) = Buffer
        CurLow = CurLow + 1
        CurHigh = CurHigh - 1
    End If
Loop

If LB < CurHigh Then
    QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

If CurLow < UB Then
    QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

'''''''''''''''''''''''''''''''''''''
' If Descending is True, reverse the
' order of the array, but only if the
' recursion level is 1.
'''''''''''''''''''''''''''''''''''''
If Descending = True Then
    If RecursionLevel = 1 Then
        ReverseArrayInPlace2 InputArray, LB, UB
    End If
End If

RecursionLevel = RecursionLevel - 1
QSortInPlace = True
End Function

Public Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortCompare
' This function is used in QSortInPlace to compare two elements. If
' V1 AND V2 are both numeric data types (integer, long, single, double)
' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
' that contain numeric data, they are converted to Doubles and compared.
' If either V1 or V2 is a string and does NOT contain numeric data, both
' V1 and V2 are converted to Strings and compared with StrComp.
'
' The result is -1 if V1 < V2,
'                0 if V1 = V2
'                1 if V1 > V2
' For text comparisons, case sensitivity is controlled by CompareMode.
' If this is vbBinaryCompare, the result is case SENSITIVE. If this
' is omitted or any other value, the result is case INSENSITIVE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim D1 As Double
Dim D2 As Double
Dim S1 As String
Dim S2 As String

Dim Compare As VbCompareMethod
''''''''''''''''''''''''''''''''''''''''''''''''
' Test CompareMode. Any value other than
' vbBinaryCompare will default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
    Compare = CompareMode
Else
    Compare = vbTextCompare
End If
'''''''''''''''''''''''''''''''''''''''''''''''
' If either V1 or V2 is either an array or
' an Object, raise a error 13 - Type Mismatch.
'''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V1) = True Or IsArray(V2) = True Then
    err.Raise 13
    Exit Function
End If
If IsObject(V1) = True Or IsObject(V2) = True Then
    err.Raise 13
    Exit Function
End If

If IsSimpleNumericType(V1) = True Then
    If IsSimpleNumericType(V2) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' If BOTH V1 and V2 are numeric data
        ' types, then convert to Doubles and
        ' do an arithmetic compare and
        ' return the result.
        '''''''''''''''''''''''''''''''''''''
        D1 = CDbl(V1)
        D2 = CDbl(V2)
        If D1 = D2 Then
            QSortCompare = 0
            Exit Function
        End If
        If D1 < D2 Then
            QSortCompare = -1
            Exit Function
        End If
        If D1 > D2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''
' Either V1 or V2 was not numeric data type.
' Test whether BOTH V1 AND V2 are numeric
' strings. If BOTH are numeric, convert to
' Doubles and do a arithmetic comparison.
''''''''''''''''''''''''''''''''''''''''''''
If IsNumeric(V1) = True And IsNumeric(V2) = True Then
    D1 = CDbl(V1)
    D2 = CDbl(V2)
    If D1 = D2 Then
        QSortCompare = 0
        Exit Function
    End If
    If D1 < D2 Then
        QSortCompare = -1
        Exit Function
    End If
    If D1 > D2 Then
        QSortCompare = 1
        Exit Function
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Either or both V1 and V2 was not numeric
' string. In this case, convert to Strings
' and use StrComp to compare.
''''''''''''''''''''''''''''''''''''''''''''''
S1 = CStr(V1)
S2 = CStr(V2)
QSortCompare = StrComp(S1, S2, Compare)

End Function

Public Function NumberOfArrayDimensions(Arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function


Public Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is sorted. An error will occur if InputArray is not an array,
 'if it is an empty, unallocated array, or if the number of dimensions is not 1.
'
' NOTE: Before calling the ReverseArrayInPlace procedure, consider if your needs can
' be met by simply reading the existing array in reverse order (Step -1). If so, you can save
' the overhead added to your application by calling this function.
'
' The function returns TRUE if the array was successfully reversed, or FALSE if
' an error occurred.
'
' If an error occurred, a message box is displayed indicating the error. To suppress
' the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long
Dim OrigN As Long
Dim NewN As Long
Dim NewArr() As Variant

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' Ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
   If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''''
        ' Zero dimensions indicates an unallocated
        ' dynamic array.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function

End Select

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure that we have only simple data types,
' not an array of objects or arrays.
'''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
    If NoAlerts = False Then
        MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "ReverseArrayInPlace can reverse only arrays of simple data types."
        Exit Function
    End If
End If

ReDim NewArr(LBound(InputArray) To UBound(InputArray))
NewN = UBound(NewArr)
For OrigN = LBound(InputArray) To UBound(InputArray)
    NewArr(NewN) = InputArray(OrigN)
    NewN = NewN - 1
Next OrigN

For NewN = LBound(NewArr) To UBound(NewArr)
    InputArray(NewN) = NewArr(NewN)
Next NewN

ReverseArrayInPlace = True
End Function


Public Function ReverseArrayInPlace2(InputArray As Variant, _
    Optional LB As Long = -1, Optional UB As Long = -1, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace2
' This reverses the order of elements in InputArray. To reverse the entire array, omit or
' set to less than 0 the LB and UB parameters. To reverse only part of tbe array, set LB and/or
' UB to the LBound and UBound of the sub array to be reversed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim N As Long
Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long
Dim OrigN As Long
Dim NewN As Long
Dim NewArr() As Variant

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
ReverseArrayInPlace2 = False

'''''''''''''''''''''''''''''''''
' Ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''''
        ' Zero dimensions indicates an unallocated
        ' dynamic array.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function

End Select

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure that we have only simple data types,
' not an array of objects or arrays.
'''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
    If NoAlerts = False Then
        MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "ReverseArrayInPlace can reverse only arrays of simple data types."
        Exit Function
    End If
End If

If LB < 0 Then
    LB = LBound(InputArray)
End If
If UB < 0 Then
    UB = UBound(InputArray)
End If

For N = LB To (LB + ((UB - LB - 1) \ 2))
    Temp = InputArray(N)
    InputArray(N) = InputArray(UB - (N - LB))
    InputArray(UB - (N - LB)) = Temp
Next N

ReverseArrayInPlace2 = True
End Function


Public Function IsSimpleNumericType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleNumericType
' This returns TRUE if V is one of the following data types:
'        vbBoolean
'        vbByte
'        vbCurrency
'        vbDate
'        vbDecimal
'        vbDouble
'        vbInteger
'        vbLong
'        vbSingle
'        vbVariant if it contains a numeric value
' It returns FALSE for any other data type, including any array
' or vbEmpty.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(V) = True Then
    Select Case VarType(V)
        Case vbBoolean, _
                vbByte, _
                vbCurrency, _
                vbDate, _
                vbDecimal, _
                vbDouble, _
                vbInteger, _
                vbLong, _
                vbSingle
            IsSimpleNumericType = True
        Case vbVariant
            If IsNumeric(V) = True Then
                IsSimpleNumericType = True
            Else
                IsSimpleNumericType = False
            End If
        Case Else
            IsSimpleNumericType = False
    End Select
Else
    IsSimpleNumericType = False
End If
End Function

Public Function IsSimpleDataType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleDataType
' This function returns TRUE if V is one of the following
' variable types (as returned by the VarType function:
'    vbBoolean
'    vbByte
'    vbCurrency
'    vbDate
'    vbDecimal
'    vbDouble
'    vbEmpty
'    vbError
'    vbInteger
'    vbLong
'    vbNull
'    vbSingle
'    vbString
'    vbVariant
'
' It returns FALSE if V is any one of the following variable
' types:
'    vbArray
'    vbDataObject
'    vbObject
'    vbUserDefinedType
'    or if it is an array of any type.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test if V is an array. We can't just use VarType(V) = vbArray
' because the VarType of an array is vbArray + VarType(type
' of array element). E.g, the VarType of an Array of Longs is
' 8195 = vbArray + vbLong.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We must also explicitly check whether V is an object, rather
' relying on VarType(V) to equal vbObject. The reason is that
' if V is an object and that object has a default proprety, VarType
' returns the data type of the default property. For example, if
' V is an Excel.Range object pointing to cell A1, and A1 contains
' 12345, VarType(V) would return vbDouble, the since Value is
' the default property of an Excel.Range object and the default
' numeric type of Value in Excel is Double. Thus, in order to
' prevent this type of behavior with default properties, we test
' IsObject(V) to see if V is an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''
' Test the value returned by VarType.
'''''''''''''''''''''''''''''''''''''
Select Case VarType(V)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType
        '''''''''''''''''''''''
        ' not simple data types
        '''''''''''''''''''''''
        IsSimpleDataType = False
    Case Else
        ''''''''''''''''''''''''''''''''''''
        ' otherwise it is a simple data type
        ''''''''''''''''''''''''''''''''''''
        IsSimpleDataType = True
End Select

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''
' If Arr is not an array, return FALSE and get out.
'''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Try to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occured.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
N = UBound(Arr, 1)
If err.Number = 0 Then
    '''''''''''''''''''''''''''''''''''''
    ' No error. Array has been allocated.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = True
Else
    '''''''''''''''''''''''''''''''''''''
    ' Error. Unallocated array.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = False
End If

End Function

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
