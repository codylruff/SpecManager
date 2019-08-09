VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KrishDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : KrishDll
' Author    : KRISH J
' Date      : 01/06/2018
' Purpose   :
' Returns   :
' 01-06-2018: ShowDialogRich chagned to MsgBox for easier migration from old msgbox
' 30-07-2018: AreYouSure & ShowDialog can be colour formatted.
' 31-07-2018: FTP Delete File added
' 01-08-2018: ImportJSON function added to create local table from a json string.
' 19/06/2019: Form related function and regex functions added
'---------------------------------------------------------------------------------------
Option Explicit

Private dllObject As Object

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function KRISH_VBA_TOOLS Lib "VBA_TOOLS.DLL" () As Object
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal Hwnd As LongPtr, ByRef lpRect As Rect) As LongPtr
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long
    Private Declare Function KRISH_VBA_TOOLS Lib "VBA_TOOLS.DLL" () As Object
    Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As LongPtr, ByRef lpRect As Rect) As LongPtr
#End If

Public Function DLL() As Object
    On Error Resume Next
    If dllObject Is Nothing Then
        'load from the dll
        LoadLibrary (ThisWorkbook.path & "\libs\vba-tools\VBA_TOOLS.dll")
        Set dllObject = KRISH_VBA_TOOLS()
        
        On Error Resume Next
        'Send this applicaiton full name to the dll for future references. i.e. call backs
        dllObject.SetAccessApplicationPath ActiveWorkbook.FullName
    End If
    
    Set DLL = dllObject
End Function


Public Sub Class_Initialize()
    Call DLL
End Sub

Public Function Toast(iMessage As String, Optional iType As VbMsgBoxStyle = vbInformation, Optional iDURATION As Long = 2000, _
    Optional iHWND As Long = 0, Optional iBG As String)
'---------------------------------------------------------------------------------------
' Procedure : Me.toast
' Author    : KRISH
' Date      : 08/02/2018
' Purpose   : SHows a toastr message
' Returns   : Non blocking Null
'---------------------------------------------------------------------------------------
'
    
On Error GoTo LABEL_EXIT_ROUTINE:
    Dim iX As Long, iY As Long
    
    'check if vba_tools object is available
    Dim iFONT_COLOR As String:    iFONT_COLOR = "#FFFFFF"
    
    If DLL Is Nothing Then GoTo LABEL_EXIT_ROUTINE Else GoTo LABEL_TOAST
    Exit Function
    
LABEL_EXIT_ROUTINE:
    MsgBox iMessage
    On Error GoTo 0
    Exit Function

LABEL_TOAST:
    
    Dim iTypeS As String
    Select Case iType
        Case vbCritical To vbCritical + 6, vbError To vbError + 6
            iTypeS = DANGER_COLOR_DARK
        Case vbExclamation To vbExclamation + 6
            iTypeS = WARNING_COLOR_DARK
            iFONT_COLOR = WHITE
        Case vbInformation To vbInformation + 6
            iTypeS = SUCCESS_COLOR
        Case Else
            iTypeS = INFO_COLOR
    End Select
        
    If Not (IsBlank(iBG)) Then
        iTypeS = iBG
    End If

    
    'if parent object is provided show the toast on top of the parent.
    Dim mRect As Rect
    If Not iHWND <= 0 Then
        GetWindowRect iHWND, mRect
    End If

    iX = mRect.Right - 10
    iY = mRect.Bottom - 10

    
    On Error Resume Next
    DLL.notification.FN_SHOW_TOAST iMessage, iDURATION, iTypeS, 1000, iFONT_COLOR, iX, iY, 0, 0
    
End Function

Public Function ShowWarning(message As String, Optional Duration As Long = 4000, Optional handle As Long)
    Toast message, vbExclamation, Duration, handle
End Function

Public Function ShowSuccess(message As String, Optional Duration As Long = 4000, Optional handle As Long)
    Toast message, vbInformation, Duration, handle
End Function

Public Function ShowError(message As String, Optional Duration As Long = 4000, Optional handle As Long)
    Toast message, vbExclamation, Duration, handle
End Function

Public Function AreYouSure(Optional ISender As Object) As Integer
'---------------------------------------------------------------------------------------
' Procedure : AreYouSure
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   : Just an are you sure form
' Returns   :
'---------------------------------------------------------------------------------------
'
    Dim Hwnd As Long
    If Not ISender Is Nothing Then Hwnd = ISender.Hwnd
    
    AreYouSure = (DLL.Dialog.AreYouSure(Hwnd) = vbYes)
End Function

Public Function AreYouSureE(Optional Sender, Optional ThemeBg As String, Optional themeForeColour As String, _
                Optional PositiveColour As String, Optional NegativeColour As String) As Integer
'---------------------------------------------------------------------------------------
' Procedure : AreYouSure
' Author    : KRISH
' Date      : 27/07/2018
' Purpose   : Just an are you sure form
' Returns   :
'---------------------------------------------------------------------------------------
'
    Dim Hwnd As Long
    If ((Not IsBlank(Sender))) Then Hwnd = Sender.Hwnd
    
    AreYouSureE = (DLL.Dialog.AreYouSureE(Hwnd, ThemeBg, themeForeColour, PositiveColour, NegativeColour) = vbYes)
End Function

Public Function ShowDialog(message As String, Optional dialogType As VbMsgBoxStyle = vbInformation, _
                Optional title As String, Optional ButtonTextForYes As String = "", Optional ButtonTextForNo As String = "", _
                Optional ThemeBg As String = SAATI_BLUE, Optional themeForeColour As String = "", Optional BodyBg As String = "") As Integer

    ShowDialog = DLL.Dialog.ShowDialog(title, message, ButtonTextForYes, ButtonTextForNo, dialogType, ThemeBg, themeForeColour, BodyBg)
End Function

Public Function ShowDialogRich(message As String, Optional Style As VbMsgBoxStyle, Optional title As String, _
                Optional Duration = 0, Optional ButtonTextForYes As String = "Yes", Optional ButtonTextForNo As String = "No", _
                Optional TitleForeColour As String = WHITE, Optional ThemeBg As String = SAATI_BLUE, Optional IsModal = True) As Integer
'---------------------------------------------------------------------------------------
' Procedure : DialogRich
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   : A vba Wrapper to directly replace msgbox function. Uses ShowDialogJson from the dll
' Returns   :
'---------------------------------------------------------------------------------------
'
    'If IsBlank(Message) Then Exit Function
    If Style = 0 Then Style = vbInformation
    
    Dim T As New TDialog
    
    T.title = title
    T.message = message
    T.dialogType = Style
    T.TitleForeColour = TitleForeColour
    T.ThemeBg = ThemeBg
    T.Duration = Duration
    T.IsModal = IsModal
    
    'Debug.Print Style
    If ((Style And vbInformation) = vbInformation) And Style = vbInformation Or ((Style And vbQuestion) = vbQuestion) And Style = vbQuestion Then
        T.ThemeBg = "#A6CC56"
    End If
    
    If (((Style And vbInformation) = vbInformation) And Style = vbInformation) Or (((Style And vbQuestion) = vbQuestion) And Style = vbQuestion) Or (((Style And vbExclamation) = vbExclamation) And Style = vbExclamation) Or (((Style And vbCritical) = vbCritical) And Style = vbCritical) Then
        ' // no buttons selected show vbokonly?
        T.ClearButtons
        T.AddButton "OK", vbOK, True
    End If
    
    'Debug.Print t.ToJSON
    ShowDialogRich = DLL.Dialog.ShowDialogJson(T.ToJSON)
End Function

Public Function ShowDialogJson(JsonConfig As String)
'---------------------------------------------------------------------------------------
' Procedure : ShowDialogJson
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   : uses json config string and directly calls the ShowDialogJson function from dll
' Returns   :
'---------------------------------------------------------------------------------------
'
    ShowDialogJson = DLL.Dialog.ShowDialogJson(JsonConfig)
End Function

Public Function ShowProgressBar(TopValue As Long, Optional title As String, Optional InnerMessage As String = "", Optional AutoClose As Boolean = False, Optional ThemeBg As String = "#A6CC56", Optional TitleForeColour As String = "#FFFFFF", Optional BodyBg As String = "#FFFFFF", Optional pBarColour As String = "#00C851")
'---------------------------------------------------------------------------------------
' Procedure : ShowProgressBar
' Author    : KRISH
' Date      : 08/05/2018
' Purpose   : Shows a colourful progressbar on top of your vba application
' Returns   : Handle of the progressbar form. USe the handle to set new values
'---------------------------------------------------------------------------------------
'
    If TopValue = 0 Then
        Me.Toast "Sorry Maximum value must be greater than 0", vbCritical, 0
    End If
    
    If title = "" Then title = "Please wait!"
    If InnerMessage = "" Then InnerMessage = "Please wait. We are preparing your task"
    
    Dim handle As Long
    handle = DLL.Dialog.OpenProgressBar(title, InnerMessage, TopValue, AutoClose, ThemeBg, TitleForeColour, BodyBg, pBarColour)
    
    ShowProgressBar = handle
End Function

Public Function SetProgressBar(ProgressBarId As Long, CurrentValue As Long, Optional InnerMessage As String = "", Optional NewTopValue As Long = 0, Optional AutoClose As Boolean = False, Optional ThemeBg = "", Optional TitleForeColour = "", Optional BodyBg = "", Optional pBarColour As String = "")
'---------------------------------------------------------------------------------------
' Procedure : SetProgressBar
' Author    : KRISH
' Date      : 08/05/2018
' Purpose   : Sets new progress step for a progressbar. Use the progressbar id received from showProgressbar function
' Returns   : Returns progressbarid
'---------------------------------------------------------------------------------------
    
    If CurrentValue = 0 Then Exit Function
    
    If ProgressBarId = 0 Then
        Me.Toast "Invalid progressbar id. Current step: " & CurrentValue, vbCritical, 0
    End If
    
    Dim handle As Long
    handle = DLL.Dialog.SetProgressBar(ProgressBarId, CurrentValue, InnerMessage, NewTopValue, AutoClose, ThemeBg, TitleForeColour, BodyBg, pBarColour)
    SetProgressBar = handle
End Function

Public Function CloseProgressBar(ProgressBarId As Long)
'---------------------------------------------------------------------------------------
' Procedure : CloseProgressBar
' Author    : KRISH
' Date      : 21/05/2018
' Purpose   : Closes a progressbar by it's handle
' Returns   :
'---------------------------------------------------------------------------------------

    If ProgressBarId = 0 Then Exit Function
    On Error Resume Next
    DLL.Dialog.CloseProgressBar ProgressBarId
End Function

Public Function SaveClipboardToImage(Optional path As String = "", Optional fileName As String = "", Optional FileType As String = "jpeg")
'---------------------------------------------------------------------------------------
' Procedure : SaveClipboardToImage
' Author    : KRISH
' Date      : 08/05/2018
' Purpose   : Saves the image from clipboard (if contains) to local file system. Returns empty string if no image found in clipbord
' Returns   : the path of saved file
'---------------------------------------------------------------------------------------

    SaveClipboardToImage = DLL.Graphics.SaveClipboardToImage(path, fileName, FileType)
End Function

Public Function DownloadAFile(url As String, Optional Destination As String = "", Optional Overwrite As Boolean = True, Optional ShowProgress = True)
'---------------------------------------------------------------------------------------
' Procedure : DownloadAFile
' Author    : KRISH
' Date      : 08/05/2018
' Purpose   : Will download a file from the Internet to the local file path. Returns the downloaded Filepath or empty string
' Returns   :
'---------------------------------------------------------------------------------------
    
    DownloadAFile = DLL.Dialog.DownloadAFile(url, Destination, Overwrite, ShowProgress)
End Function


' StringClass
Public Function TrimEnd(ByRef text As String, trimChar As String) As String
    text = DLL.String.TrimEnd(text, trimChar)
    TrimEnd = text
End Function

Public Function TrimStart(ByRef text As String, trimChar As String) As String
    text = DLL.String.TrimStart(text, trimChar)
    TrimStart = text
End Function

Public Function PadLeft(ByRef InputString As String, length As Integer, Optional PadWithChar As String = "")
'For question: https://stackoverflow.com/questions/12060347/any-method-equivalent-to-padleft-padright
    InputString = DLL.String.PadLeft(InputString, length, Left(PadWithChar, 1))
    PadLeft = InputString
End Function

Public Function PadRight(ByRef InputString As String, length As Integer, Optional PadWithChar As String = "")
    InputString = DLL.String.PadRight(InputString, length, Left(PadWithChar, 1))
    PadRight = PadRight
End Function

Public Function StartsWith(text As String, startsWith_ As String) As Boolean
    StartsWith = DLL.String.StartsWith(text, startsWith_)
End Function
Public Function EndsWith(text As String, startsWith_ As String) As Boolean
    EndsWith = DLL.String.EndsWith(text, startsWith_)
End Function

Public Function ToCamelCase(ByRef text As String, Optional replaceInvalidString As Boolean = True, Optional invalidStringArray = "") As String
'        /// <summary>
'        /// Converts a long string to camelCasedString. If replaceInvalidString is set true, _ - . are removed
'        /// </summary>
'        /// <param name="input"></param>
'        /// <returns></returns>
    text = DLL.String.ToCamelCase(text, replaceInvalidString, invalidStringArray)
    ToCamelCase = text
End Function

'GraphicsClass
Public Function ImageControlGetImage(ImagePath As String, Optional ShowError As Boolean = True)
'---------------------------------------------------------------------------------------
' Procedure : GetImage
' Author    : KRISH
' Date      : 11/05/2018
' Purpose   : Retireves an image object from any web url. If the url is a local file path, just set as picture=filepath
' Returns   :
'---------------------------------------------------------------------------------------
'
    ImageControlGetImage = DLL.Graphics.PictureFromUrl(url:=ImagePath, ShowError:=True)
End Function

Public Function PicturePathFromUrl(url As String, Optional ShowError As Boolean = True, Optional SavePath As String = "", Optional useCache As Boolean = True)
'---------------------------------------------------------------------------------------
' Procedure : PicturePathFromUrl
' Author    : KRISH
' Purpose   : Downloads an image from the given url, saves it locallay and returns the local path
' Returns   : local path
'---------------------------------------------------------------------------------------
'
    PicturePathFromUrl = DLL.Graphics.PictureFromUrlP(url, ShowError, SavePath, useCache)
End Function

Public Function FN_IMAGE_LOAD(byteArr() As Byte, Optional SavePath As String) As String
'---------------------------------------------------------------------------------------
' Procedure : FN_IMAGE_LOAD
' Author    : KRISH
' Date      : 30/10/2017
' Purpose   : Saves byteArray into an image and returns the path.
'---------------------------------------------------------------------------------------

    'check if vba_tools object is available
   On Error GoTo FN_IMAGE_LOAD_Error

    
    If (IsNull(byteArr)) Or IsBlank(byteArr) Then
        FN_IMAGE_LOAD = ""
        Exit Function
    Else
        If IsMissing(SavePath) Or IsBlank(SavePath) Then SavePath = Environ$("temp")
        FN_IMAGE_LOAD = DLL.Graphics.ByteToImage(byteArr, SavePath)
    End If

   On Error GoTo 0
   Exit Function

FN_IMAGE_LOAD_Error:
    App.gMsg = err.Description & ": in procedure KrishDll of Module KrishDll"
    FN_IMAGE_LOAD = ""

End Function

Public Function OpenFileReturnsArray(title As String, Filters As Variant, Optional AllowMultipleFiles As Boolean = False)
 '---------------------------------------------------------------------------------------
' Procedure : OpenFile
' Author    : KRISH
' Date      : 21/05/2018
' Purpose   :
' Returns   : String array
'---------------------------------------------------------------------------------------
'
    On Error Resume Next
    Dim FilePaths() As String
    Dim FiltersS() As String
    Dim item As Variant
    
    If (IsNull(Filters)) Then
        FnArrayAddItem FiltersS, "All files |*.*"
    Else
        For Each item In Filters
            If item <> "" Then
                FnArrayAddItem FiltersS, CStr(item)
            End If
        Next item
        
    End If
    
    FilePaths = DLL.Dialog.ShowDialogForFileArray(title, AllowMultipleFiles, Filters)

    OpenFileReturnsArray = FilePaths
End Function

Public Function OpenFileInlineFilterArray(title As String, Optional Filters As Variant, Optional AllowMultipleFiles As Boolean = False)
 '---------------------------------------------------------------------------------------
' Procedure : OpenFile
' Author    : KRISH
' Purpose   : File open dialog with inline array filter.
' Returns   : String array
' Usage     : App.gDll.OpenFileInlineFilterArray("Open Png files",Array("Png Pictures only |*.png"), false)
'---------------------------------------------------------------------------------------
'
    On Error Resume Next
    Dim FilePaths() As String
    Dim FiltersS() As String
    Dim item As Variant
    
    If (IsNull(Filters) Or IsMissing(Filters)) Then
        FnArrayAddItem FiltersS, "All files |*.*"
    Else
        For Each item In Filters
            If item <> "" Then
                FnArrayAddItem FiltersS, CStr(item)
            End If
        Next item
        
    End If
    FilePaths = DLL.Dialog.ShowDialogForFileArray(title, AllowMultipleFiles, FiltersS)

    OpenFileInlineFilterArray = FilePaths
End Function

Public Function OpenFile(title As String, Optional FilterUsesInlineArray As Object, Optional AllowMultipleFiles As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : OpenFile
' Author    : KRISH J
' Date      : 05/02/2019
' Purpose   : allows to open for specific files. Uses inline Array() for filter
' Returns   : FilePath of the selected file, or jsonArray if multiFiles are requested
' Usage     : App.gDll.OpenFile("Open Png files",Array("Png Pictures only |*.png"), false)
'---------------------------------------------------------------------------------------
'

    On Error Resume Next
    If IsBlank(FilterUsesInlineArray) Then
        OpenFile = DLL.Dialog.ShowDialogForFile(title, AllowMultipleFiles, Array("Everything |*.*"))
    Else
        OpenFile = DLL.Dialog.ShowDialogForFile(title, AllowMultipleFiles, FilterUsesInlineArray)
    End If
    
End Function

Public Function OpenFileA(title As String, FilterUsesStringArray() As String, Optional AllowMultipleFiles As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : OpenFile
' Author    : KRISH J
' Date      : 05/02/2019
' Purpose   : allows to open for specific files. Uses a string array for filters
' Returns   : FilePath of the selected file, or jsonArray if multiFiles are requested
' Usage     : App.gDll.OpenFile("Open Png files",Array("Png Pictures only |*.png"), false)
'---------------------------------------------------------------------------------------
'

    On Error Resume Next
    If IsBlank(FilterUsesStringArray) Then
        ReDim FilterUsesStringArray(1)
        FilterUsesStringArray(0) = "Everything |*.*"
    End If
    
    OpenFileA = DLL.Dialog.ShowDialogForFileA(title, AllowMultipleFiles, FilterUsesStringArray)
    
End Function

' ------------------------------- --- ----------------------------------------------------
' ------------------------------- JSONClass ----------------------------------------------------
' ------------------------------- --- ----------------------------------------------------

Public Function JsonToString(jsonObject As Variant, Optional Formatting As JsonFormatting = JsonFormatting.Indented) As String
'---------------------------------------------------------------------------------------
' Procedure : JsonConvertToJson
' Author    : KRISH J
' Date      : 10/09/2018
' Purpose   : Converts json object to string
' Returns   : String representation of the Newtonsoft.Json.Linq.JObject
'---------------------------------------------------------------------------------------
'
    JsonToString = DLL.json.JsonString(jsonObject, Formatting)
End Function

Public Function JsonNewObject()
'---------------------------------------------------------------------------------------
' Procedure : JsonNewObject
' Author    : KRISH J
' Date      : 10/09/2018
' Purpose   :
' Returns   : An empty JObject
'---------------------------------------------------------------------------------------
'
    Set JsonNewObject = JsonParse("{}")
End Function

Public Function JsonParse(JsonString As String)
'---------------------------------------------------------------------------------------
' Procedure : JsonParse
' Author    : KRISH
' Date      : 22/05/2018
' Purpose   : Will retun Newtonsoft.Json.Linq.JObject from .NET which can be used by the jsonGetValue, jsonGetObject functions
' Returns   :
'---------------------------------------------------------------------------------------
'

    Set JsonParse = DLL.json.JsonToObject(JsonString)
End Function

Public Function JsonGetValue(parsedJsonObject As Variant, propertyName As String)
'---------------------------------------------------------------------------------------
' Procedure : JsonGetValue
' Author    : KRISH
' Date      : 22/05/2018
' Purpose   : returns string value of the json property name
' Returns   :
'---------------------------------------------------------------------------------------
'
    JsonGetValue = DLL.json.JsonGetValue(parsedJsonObject, propertyName)
End Function

Public Function JsonGetObject(parsedJsonObject As Variant, propertyName As String)
'---------------------------------------------------------------------------------------
' Procedure : JsonGetObject
' Author    : KRISH
' Date      : 22/05/2018
' Purpose   : Will retun Newtonsoft.Json.Linq.JObject from .net. Use this to extract objects i.e. array objects from a root json object.
' Returns   :
'---------------------------------------------------------------------------------------
'

    Set JsonGetObject = DLL.json.JsonGetObject(parsedJsonObject, propertyName)
End Function

Public Function JsonSetObject(ByRef parsedJsonObject As Variant, propertyName As String, propertyValue As Variant)
'---------------------------------------------------------------------------------------
' Procedure : JsonSetObject
' Author    : KRISH J
' Date      : 10/09/2018
' Purpose   : Adds new JProperty to an existing JObject
' Returns   : JObject
'---------------------------------------------------------------------------------------
'
    Set parsedJsonObject = DLL.json.JsonSetObject(parsedJsonObject, propertyName, propertyValue)
    Set JsonSetObject = parsedJsonObject
End Function

Public Function JsonIsArray(ByRef parsedJsonObject As Variant) As Boolean
'/// returns true if the object is a json array
    JsonIsArray = DLL.json.JsonIsArray(parsedJsonObject)
End Function

Public Function JsonGetArrayItem(ByRef parsedJsonObject As Variant, index As Integer) As Boolean
'Returns Json Object from Json array object

    JsonGetArrayItem = DLL.json.JsonIsArray(JsonGetArrayItem, index)
End Function

Public Function JsonGetCount(ByRef jsonObject As Variant) As Long
'Returns number of tokens in a json node. will not recursivly check sub nodes

    JsonGetCount = DLL.json.JsonGetCount(jsonObject)
End Function

Public Function JsonGetChildrenCount(ByRef jsonObject As Variant) As Long
'Returns number of tokens in a json node. will not recursivly check sub nodes

    JsonGetChildrenCount = DLL.json.JsonChildrenCount(jsonObject)
End Function
        
Public Function JsonGetNextItem(ByRef jsonObject As Variant)
'Returns the next sibiling node from this node

    JsonGetNextItem = DLL.json.JsonGetNextItem(jsonObject)
End Function

Public Function JsonGetItem(ByRef jsonObject As Variant, index As Integer) As Variant
'Returns the jobject found at the given index from the root node

    Set JsonGetItem = DLL.json.JsonGetItem(jsonObject, index)
End Function


' ------------------------------- END JSON ----------------------------------------------------



Public Function sFormat(ByRef iString, ParamArray args() As Variant) As String
'---------------------------------------------------------------------------------------
' Procedure : sFormat
' Author    : KRISH
' Date      : 08/02/2018
' Purpose   : Uses .Net string format function using the dll. Arg(0) is the first entry
' Returns   : converted strin
'---------------------------------------------------------------------------------------
'

    On Error Resume Next
    Dim i       As Long
    Dim mParams As String
    Dim mArgLen As Long
    
    sFormat = iString
    mArgLen = UBound(args)
    
    If mArgLen < 0 Then
        Exit Function
    End If
    
On Error GoTo EXIT_ROUTINE:

    Select Case mArgLen
        Case 0
            sFormat = DLL.String.fn_string_format(iString, args(0))
        Case 1
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1))
        Case 2
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2))
        Case 3
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3))
        Case 4
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4))
        Case 5
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 6
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 7
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 8
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 9
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case 10
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
        Case 11
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
        Case 12
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
        Case 13
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13))
        Case 14
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14))
        Case 15
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15))
        Case 16
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16))
        Case 17
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17))
        Case 18
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18))
        Case 19
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19))
        Case 20
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20))
        Case 21
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21))
        Case 22
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22))
        Case 23
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23))
        Case 24
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24))
        Case 25
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25))
        Case 26
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26))
        Case 27
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27))
        Case 28
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28))
        Case 29
            sFormat = DLL.String.fn_string_format(iString, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29))
    Case Else
        err.Raise "0", "DLL", "Too many parameters"
        Exit Function
    End Select
    
    
    Exit Function
    
    
EXIT_ROUTINE:
    On Error Resume Next
    For i = 0 To mArgLen
        iString = VBA.Replace(iString, "{" & i & "}", args(i), , , vbDatabaseCompare)
    Next i

    sFormat = iString
End Function


Public Function FTPUploadFile(Server As String, Port As Long, Username As String, Password As String, localFile As String, remoteLocation As String, Optional TLSHostFingerprint = "", Optional Protocol As ProtocolEnum = ProtocolEnum.ftp, Optional FtpSecure As FtpSecureEnum = FtpSecureEnum.Explicit, Optional giveUpSecurityAndAcceptAnyTlsHostCertificate As Boolean = False)
'---------------------------------------------------------------------------------------
' Procedure : FTPUploadFile
' Author    : KRISH J
' Date      : 04/09/2018
' Purpose   : Uplods a file to remote server.
' Returns   : True or false or server message.
' Usage     : FTPUploadFile(ServerName, 21, UserName,Password,"C:\somefile.png", "/RemoteFoder/",SSHFingerprint, ProtocolType, FtpSecureType, yes/no whether to use sshfingerprint.
' giveUpSecurityAndAcceptAnyTlsHostCertificate = yes means connection will accept any server sshfingerprint. use this for debug purpose only
' Read more in https://winscp.net/forum/viewtopic.php?t=15428
'---------------------------------------------------------------------------------------
'
    'USE giveUpSecurityAndAcceptAnyTlsHostCertificate only for debuggin purposes
    
    Dim result As String
    result = DLL.ftp.FTPS_UPLOAD(Server, Port, Username, Password, localFile, remoteLocation, TLSHostFingerprint, Protocol, FtpSecure, giveUpSecurityAndAcceptAnyTlsHostCertificate)
    
    If result = "true" Then
    
        'File has been uploaded to the server.
        'Construct your own method to convert the fileName to remote filename. i.e. http://yourdomain/yourfolder/filename.ext
        result = System.FN_FILE_GET_NAME(localFile) ' Returning local file path
        Me.Toast "File uploaded to the server"
        
    Else
        Me.ShowDialog sFormat("Unable to upload the file: {0}", result), vbExclamation, "File upload error"
        result = ""
    End If
    
    FTPUploadFile = result
    
    
End Function

Public Function FTPDeleteFile(Server As String, Port As Long, Username As String, Password As String, RemoteFile As String, Optional TLSHostFingerprint = "", Optional Protocol As ProtocolEnum = ProtocolEnum.ftp, Optional FtpSecure As FtpSecureEnum = FtpSecureEnum.Explicit, Optional giveUpSecurityAndAcceptAnyTlsHostCertificate As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : FTPDeleteFile
' Author    : KRISH J
' Date      : 31/07/2018
' Purpose   : Deletes a file from remote server
' Returns   : string, true or failed or error message
'---------------------------------------------------------------------------------------
'

    If IsBlank(RemoteFile) Then Exit Function
    
    RemoteFile = VBA.Replace(RemoteFile, "", "", , , vbDatabaseCompare) ' Do some url manipulatios if you require. I.e remove the web hyperlink part
    If TLSHostFingerprint = "" Then TLSHostFingerprint = "" ' place your sshfingerprint
    
    Dim result As String
    result = DLL.ftp.FTPDeleteFile(Server, Port, Username, Password, RemoteFile, TLSHostFingerprint, Protocol, FtpSecure, giveUpSecurityAndAcceptAnyTlsHostCertificate)
    If result = "true" Then
        
        Me.Toast "File removed from server"
    Else
        Me.ShowDialog sFormat("Unable to delete the file: {0}", result), vbExclamation, "File Delete", "OK", ""
    End If
    
    FTPDeleteFile = result
    
End Function


Public Function ImportJSON(JsonSource As String, TableName As String, ImportOptions As AcImportXMLOption, Optional Recreate As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : ImportJSON
' Author    : KRISH J
' Date      : 01/08/2018
' Purpose   : Imports Json string to a table.
' Returns   : string; True or false + error message
'---------------------------------------------------------------------------------------
'

    ImportJSON = DLL.dbimportjson(JsonSource, ActiveWorkbook.FullName, TableName, ImportOptions, Recreate)
    
End Function


Public Function ExportToJSON(TableOrQueryName As String, SaveAs As String, Optional Overwrite As Boolean = False, Optional isRawSql As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : ImportJSON
' Author    : KRISH J
' Date      : 01/08/2018
' Purpose   : Imports Json string to a table.
' Returns   : string; True or false + error message
'---------------------------------------------------------------------------------------
'
''
    ExportToJSON = DLL.DBExportToJSON(ActiveWorkbook.FullName, TableOrQueryName, SaveAs, Overwrite, isRawSql)
    
End Function



Public Function CreateBarcode(symbology As symbology, barcodeDataValue As String, multiplier As Double)
'---------------------------------------------------------------------------------------
' Procedure : CreateBarcode
' Author    : KRISH J
' Date      : 03/08/2018
' Purpose   : Generates Barcodes and returns a barcode picture as PictureData. Can be used to Image.picturedata =
' Returns   : PictureData
'---------------------------------------------------------------------------------------
'
'    /// <summary>
'    ///
'    /// </summary>
'    /// <param name="symbology">Type of the barcode</param>
'    /// <param name="barcodeData">Data value for barcode</param>
'    /// <param name="width">Width of the graphics / Image</param>
'    /// <param name="height">height of the grapics/ Image</param>
'    /// <param name="multiplier">Multiply the size by this value.</param>
'    /// <returns>Picture Data</returns>
'    CreateBarcode(Symbology symbology, string barcodeData, int width, int height, float multiplier )

    CreateBarcode = DLL.Barcode.CreateBarcode(symbology, barcodeDataValue, 300, 300, multiplier)
End Function


Public Function ShowDropDown(title As String, message As String, dataSource As String, boundColumn As Integer, ColumnWidths As Variant, Optional isRawSql As Boolean = False, Optional posX As Long = 0, Optional posY As Long = 0, Optional themeColour As String = "#00C851", Optional themeForeColour = "#FFFFFF", Optional doneButtonCaption As String = "Done") As String
'---------------------------------------------------------------------------------------
' Procedure : ShowDropDown
' Author    : KRISH J
' Date      : 08/08/2018
' Purpose   : Shows a dropdown box for input. This is an InputBox extension.
' Returns   : String value of the BoundColumn
'---------------------------------------------------------------------------------------
'

'        /// <summary>
'        /// Shows a Dialogbox with a dropdown for selection. Returns a string value
'        /// </summary>
'        /// <param name="title">Title for the inputbox</param>
'        /// <param name="message">Inner message for the input box</param>
'        /// <param name="dbSource">Database path</param>
'        /// <param name="tableSource">Table name or SQL. If SQL is used, use isRawSql=true</param>
'        /// <param name="boundColumn">Column Index to get the value from</param>
'        /// <param name="columnWidths">an Array of integers</param>
'        /// <param name="isRawSql">Specifies whether the tablesource is a plain SQL command </param>
'        /// <param name="posX"></param>
'        /// <param name="posY"></param>
'        /// <param name="themeColour"></param>
'        /// <param name="themeForeColour"></param>
'        /// <param name="doneButtonCaption">Button Text</param>
'        /// <returns>String value</returns>
    
    'ShowDropDown = DLL.ShowDropDownBox(title:=title, message:=message, dbSource:=activeworkbook.FullName, tableSource:=dataSource, boundColumn:=boundColumn, ColumnWidths:=ColumnWidths, isRawSql:=isRawSql, posX:=posX, posY:=posY, themeColour:=themeColour, themeForeColour:=themeForeColour)
    
    ShowDropDown = DLL.Dialog.ShowDropDownBox(title, message, ActiveWorkbook.FullName, dataSource, boundColumn, ColumnWidths, isRawSql, posX, posY, themeColour, themeForeColour, doneButtonCaption)
    
End Function


Public Function ShowContexMenu(MenuITems As Variant, Optional Hwnd As Long, Optional IconBarWidth As Long = 50, Optional MenuWidth As Long = 200, Optional RowHeight As Long = 44, Optional BgColour As String = "#FFFFFF", Optional SelectionBgColour As String = "#00C851", Optional ForeColour As String = "#9e9e9e", Optional SelectionForeColour As String = "#FFFFFF", Optional IconBarBgColour As String = "#2E2E2E", Optional IconBarSelectionBgColour As String = "#2E2E2E", Optional X As Long = 0, Optional Y As Long = 0, Optional canMove As Boolean = True, Optional DefaultReturnValue As String = "")
'---------------------------------------------------------------------------------------
' Procedure : ShowContexMenu
' Author    : KRISH J
' Date      : 13/08/2018
' Purpose   : Shows a moveable context menu at the cursor click position
' Returns   : String data value
'---------------------------------------------------------------------------------------
'
    
'       /// <summary>
'        /// Shows a moveable context menu at a given x,y or last mouse position.
'        /// </summary>
'        /// <param name="menuItems">Array of strings/ Menu items</param>
'        /// <param name="hWnd">Form handler</param>
'        /// <param name="iconBarWidth">Width of the icon bar (left side menu)</param>
'        /// <param name="menuWidth">Max width of the menu item (-1 = autosize) </param>
'        /// <param name="rowHeight">Rowheight for each menu item = default 44</param>
'        /// <param name="bgColour">Background color for each menu items</param>
'        /// <param name="selectionBgColour">Hover background colour for menu items</param>
'        /// <param name="foreColour">Font colour for menu item</param>
'        /// <param name="selectionForeColour">Font colour when selecting menu item</param>
'        /// <param name="iconBarBgColour">Iconbar default background colour</param>
'        /// <param name="iconBarSelectionColour">Iconbar hover background colour</param>
'        /// <param name="X">X/Left Coordinate of the menu</param>
'        /// <param name="Y">Y/Top Coordinate of the menu</param>
'        /// <param name="canMove">Can this mouse be moved?</param>
'        /// <param name="defaultSelected">Menu Item to return if none selected</param>
'        /// <returns></returns>
        
    ShowContexMenu = DLL.ShowContextMenu(MenuITems, Hwnd, IconBarWidth, MenuWidth, RowHeight, BgColour, SelectionBgColour, ForeColour, SelectionForeColour, IconBarBgColour, IconBarSelectionBgColour, X, Y, canMove, DefaultReturnValue)
End Function


Public Function ShowContextMenuA(MenuITems As Variant, _
    Optional Hwnd As Long, _
    Optional IconBarWidth As Integer = 50, _
    Optional MenuWidth As Long = 200, _
    Optional RowHeight As Integer = 44, _
    Optional backColour As String = "#ffffff", _
    Optional SelectionBackColour As String = "#00C851", _
    Optional ForeColour As String = "#9e9e9e", _
    Optional SelectionForeColour = "#ffffff", _
    Optional iconbarBackColour As String = "#212121", _
    Optional iconbarSelectionBackColour As String = "#212121", _
    Optional X As Long, _
    Optional Y As Long, _
    Optional canMove As Boolean = True, _
    Optional defaultSelected As String = "")
'---------------------------------------------------------------------------------------
' Procedure : ShowContexMenu
' Author    : KRISH J
' Date      : 13/08/2018
' Purpose   : Shows a moveable and Iconed context menu at the cursor click position.
' Returns   : String data value
'---------------------------------------------------------------------------------------
'
    
'        /// <summary>
'        /// Shows a moveable context menu at a given x,y or last mouse position. Extended version
'        /// </summary>
'        /// <param name="menuItems">Array(PathToIcon, DataValueId, DisplayValue)</param>
'        /// <param name="hWnd">handle to the parent window. If none given, parent window taken where the mouse is found</param>
'        /// <param name="iconBarWidth">Width of the icon bar -1 hides the icon bar</param>
'        /// <param name="menuWidth">Width of the menu column (-1 auto size )</param>
'        /// <param name="rowHeight">Row height</param>
'        /// <param name="bgColour"></param>
'        /// <param name="selectionBgColour"></param>
'        /// <param name="foreColour"></param>
'        /// <param name="selectionForeColour"></param>
'        /// <param name="iconBarBgColour"></param>
'        /// <param name="iconBarSelectionColour"></param>
'        /// <param name="x">X/Left Coordinate of the menu</param>
'        /// <param name="y">Y/Top Coordinate of the menu</param>
'        /// <param name="canMove">Can this mouse be moved?</param>
'        /// <param name="defaultSelected">Menu Item to return if none selected</param>
'        /// <returns></returns>

    ShowContextMenuA = DLL.ShowContextMenuA(MenuITems, Hwnd, IconBarWidth, MenuWidth, RowHeight, backColour, SelectionBackColour, ForeColour, SelectionForeColour, iconbarBackColour, iconbarSelectionBackColour, X, Y, canMove, defaultSelected)
End Function


' RegEx functions
Public Function RegExIsMatch(text As String, Pattern As String) As Boolean
    RegExIsMatch = DLL.regex.IsMatch(text, Pattern)
End Function

Public Function RegExGetFirstMatch(text As String, Pattern As String) As Boolean
    RegExGetFirstMatch = DLL.regex.IsMatch(text, Pattern)
End Function

Public Function RegExReplace(text As String, Pattern As String) As Boolean
    RegExReplace = DLL.regex.IsMatch(text, Pattern)
End Function


'Display function
Public Function GetNumberOfMonitors()
    GetNumberOfMonitors = DLL.Display.GetNumberOfMonitors()
End Function

Public Function GetPrimaryMonitorHandle()
    GetPrimaryMonitorHandle = DLL.Display.GetPrimaryMonitorHandle()
End Function

Public Function GetPrimaryMonitorBounds()
    GetPrimaryMonitorBounds = DLL.Display.GetPrimaryMonitorBounds()
End Function

Public Function GetMonitorBoundsByHandle()
    GetMonitorBoundsByHandle = DLL.Display.GetMonitorBoundsByHandle()
End Function

Public Function GetCursorPosition()
    GetCursorPosition = DLL.Display.GetCursorPosition()
End Function


' Net class

Public Function UrlIsReachable(url As String)
    UrlIsReachable = DLL.Net.UrlIsReachable(url)
End Function

Public Function UrlIsValid(url As String)
    UrlIsValid = DLL.Net.UrlIsValid(url)
End Function

Public Function UrlIsLocalPath(url As String)
    UrlIsLocalPath = DLL.Net.UrlIsLocalPath(url)
End Function

Public Function GetExternalIP()
    GetExternalIP = DLL.Net.GetExternalIP()
End Function

'FormClass
Public Function IsKeyDown(Key As Integer) As Boolean
    'Returns if a key is being pressed.
    'for Alt use : VBA.KeyConstants.vbKeyMenu
    IsKeyDown = DLL.Form.IsKeyDown(Key)
End Function

Public Function FormDisableCloseButton(handle As Long, Optional show As Boolean = False)
'    /// <summary>
'    /// Hides the X button in any form
'    /// </summary>
'    /// <param name="handle">Handle of the window</param>
'    /// <param name="show">True = show, false = hide</param>
    DLL.Form.DisableCloseButton handle, show
End Function

Public Function FnGetWindowRect(Hwnd As Long, ByRef Rect As Rect)
    GetWindowRect Hwnd, Rect
End Function

Public Function DragMe(handle As Long)
    DLL.Form.DragMe (handle)
End Function

Public Function ChangeAccessFormColor(HexColourCode As String)

    DLL.Form.RepaintAccessMDI Application.hWndAccessApp, HexColourCode
    
End Function

Public Function MakeTransparentForm(handle As Long, AlphaValue As Integer, Optional HexColour As String = "")
'---------------------------------------------------------------------------------------
' Procedure : MakeTransparentForm
' Author    : KRISH J
' Date      : 01/02/2019
' Purpose   : Makes a form transparent at first call.
' Returns   :
' Usage     : AlphaValue = 0-255.
'---------------------------------------------------------------------------------------
'

    DLL.Form.MakeTransparent handle, HexColour, AlphaValue
End Function

Public Function FormRepaintWindow(handle As Long, newColour As String)
    DLL.Form.RepaintWindow handle, newColour
End Function

Public Function CreateInputBox(input_type As InputBoxType, title_text As String, message_text As String) As Variant
    
    CreateInputBox = App.gDll.DLL.Dialog.showinputbox(Type:=input_type, _
                        title:=title_text, message:=message_text, ThemeBg:="", themeForeColour:="")
End Function

