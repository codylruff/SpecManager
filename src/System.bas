Attribute VB_Name = "System"
Option Explicit

Public Const WU_LOGPIXELSX = 88
Public Const WU_LOGPIXELSY = 90

Public Const POINTERSIZE As Long = 4
Public Const ZEROPOINTER As Long = 0

#If VBA7 Then
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Sub CopyString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As LongPtr, ByVal source As LongPtr, ByVal bytes As Long)
    Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef source As Any, ByVal length As Long)
    
    Public Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
    Public Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
    
#Else
    Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare Sub CopyString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As LongPtr, ByVal source As LongPtr, ByVal bytes As Long)
    Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef source As Any, ByVal length As Long)
    
    Public Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
    Public Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
    
#End If

'-----------------------------------------------------------------------------
' True if the argument is Nothing, Null, Empty, Missing or an empty string .
' http://blog.nkadesign.com/2009/access-checking-blank-variables/
'-----------------------------------------------------------------------------
Public Function IsBlank(arg As Variant) As Boolean
    Select Case VarType(arg)
        Case vbEmpty
            IsBlank = True
        Case vbNull
            IsBlank = True
        Case vbString
            IsBlank = (arg = nullstr Or arg = vbNullChar)
        Case vbObject
            IsBlank = (arg Is Nothing)
        Case Else
            IsBlank = IsMissing(arg)
    End Select
End Function


Public Function GetBasePath()
'---------------------------------------------------------------------------------------
' Procedure : GetBasePath
' Author    : KRISH
' Date      : 21/05/2018
' Purpose   : Returns app working path
' Returns   :
'---------------------------------------------------------------------------------------
'

    Dim FN As String
    FN = ActiveWorkbook.path
    If VBA.Right(FN, 1) <> "\" Then FN = FN & "\"
    GetBasePath = FN
End Function

Public Function INC(ByRef iValue As Long) As Long
    iValue = iValue + 1
    INC = iValue
End Function



Public Function FnArrayHasitem(ByRef iArray) As Boolean
    
    On Error Resume Next
    FnArrayHasitem = Not IsBlank(iArray(0))
End Function
Public Function FnArrayAddItem(ByRef iArray() As String, iItem As String)
'---------------------------------------------------------------------------------------
' Procedure : ArrayAddItem
' Author    : KRISH
' Date      : 13/03/2018
' Purpose   : Adds an item to the array
' Returns   :
'---------------------------------------------------------------------------------------
'

    On Error Resume Next
    Dim t As Long
    t = UBound(iArray())
    
    If t <= 0 Then
        t = 1
    Else
        INC t
    End If
    
    
    ReDim Preserve iArray(t)
    iArray(t - 1) = iItem
    
End Function

Public Function FN_FILE_GET_NAME(iFN As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FN_FILE_GET_NAME = FSO.GetFileName(iFN)
    Set FSO = Nothing
End Function

Function FN_FILE_GET_EXTENSION(strFilePath)
    On Error Resume Next
    FN_FILE_GET_EXTENSION = ""
    Dim mFSO As Object
    Set mFSO = CreateObject("Scripting.FileSystemObject")
    FN_FILE_GET_EXTENSION = mFSO.GetExtensionName(strFilePath)
End Function
Public Function FN_FILE_GET_PATH(iFN As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FN_FILE_GET_PATH = Replace(FSO.GetAbsolutePathName(iFN), FSO.GetFileName(iFN), "", , , vbDatabaseCompare)
    Set FSO = Nothing
End Function


Public Function FN_GET_UUID() As String
'---------------------------------------------------------------------------------------
' Procedure : FN_GET_UUID
' Author    : KRISH
' Purpose   :
' Returns   : string formatted guid
'---------------------------------------------------------------------------------------
'

    Dim sGUID   As String
    Dim tGUID   As GUID
    Dim bGuid() As Byte
    Dim lRtn    As Long
    Const clLen As Long = 50
    
   On Error GoTo FN_GET_UUID_Error

    lRtn = CoCreateGuid(tGUID)
    If lRtn = 0 Then
        bGuid = String(clLen, 0)
        lRtn = StringFromGUID2(tGUID, VarPtr(bGuid(0)), clLen)
        If lRtn > 0 Then
            sGUID = Mid$(bGuid, 2, 36)
        End If
        FN_GET_UUID = sGUID
    End If

   On Error GoTo 0
   Exit Function

FN_GET_UUID_Error:
    GUI.gMsg = err.description & ": in procedure FN_GET_UUID of Module mod_system_"
    GUI.Krish.ShowDialogRich "Error " & err.Number & " (" & GUI.gMsg & ")", vbExclamation
End Function

' (Module: Messages.bas)
Public Function SayHi(Name As Variant) As String
  SayHi = "Howdy " & Name & "!"
End Function
