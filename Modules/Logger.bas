Attribute VB_Name = "Logger"
'@exclude.json
Option Explicit
Private FSO As Object ' Declare a FileSystemObject.
Private stream As Object ' Declare a TextStream.
Private buffer As Object
Private log_level As Long
Private folder_path As String
Private file_path As String
Private log_enabled As Boolean

Public Sub SetLogLevel(level As Long)
' Sets the log level to more or less verbose output
    log_level = level
End Sub

Public Sub Log(Text As String)
    If buffer Is Nothing Then Set buffer = CreateObject("Scripting.Dictionary")
    Do Until Not buffer.exists(TimeInMS)
        Application.Wait (Now + TimeValue("0:00:01") / 1000)
    Loop
    buffer.Add Key:=TimeInMS, Item:=Text
    Debug.Print Logger.printf("{0} : {1}", TimeInMS, Text)
End Sub

Public Sub Trace(Text As String)
' Used to signify a transition point in the application log
    Log "------------- " & Text
End Sub

Public Sub ResetLog(Optional log_type As String = "runtime")
    Logger.SaveLog log_type
    Logger.ClearBuffer
End Sub

Public Sub NotImplementedException()
    Logger.Log "Not Implemented Exception!"
End Sub

Public Sub SaveLog(Optional file_name As String = "runtime")
    Dim Key As Variant
    If buffer Is Nothing Then Exit Sub
    folder_path = ThisWorkbook.path & "\logs"
    file_path = folder_path & "\" & file_name & ".log"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(folder_path) Then FSO.CreateFolder folder_path
    Logger.Log "Saving : " & file_name & ".log"
    Set stream = FSO.CreateTextFile(file_path, True)
    For Each Key In buffer
      stream.WriteLine Logger.printf("{0} : {1}", Key, buffer.Item(Key))
    Next Key
    stream.Close
End Sub

Public Sub ClearBuffer()
  Set buffer = Nothing
End Sub

Private Function TimeInMS() As String
    TimeInMS = Strings.Format(Now, "dd-MMM-yyyy HH:nn:ss") & "." & _
               Strings.Right(Strings.Format(Timer, "#0.00"), 2)
End Function

Private Function printf(mask As String, ParamArray tokens()) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

