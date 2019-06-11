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

Public Sub Log(text As String)
    If buffer Is Nothing Then Set buffer = New VBA.Collection
    buffer.Add AddLine(text)
    Debug.Print Logger.printf("{0} : {1}", buffer(buffer.count)(0), buffer(buffer.count)(1))
End Sub

Public Sub Trace(text As String)
' Used to signify a transition point in the application log
    Log "------------- " & text
End Sub

Public Sub ResetLog(Optional log_type As String = "runtime")
    Logger.SaveLog log_type
    Logger.ClearBuffer
End Sub

Private Function AddLine(text As String) As Variant
    Dim line As Variant

    line = Array(TimeInMS, text)
    
    AddLine = line
    
End Function

Public Sub NotImplementedException()
    Logger.Log "Not Implemented Exception!"
End Sub

Public Sub SaveLog(Optional file_name As String = "runtime")
    Dim line As Variant
    Dim i As Long
    If buffer Is Nothing Then Exit Sub
    folder_path = ThisWorkbook.path & "\logs"
    file_path = folder_path & "\" & file_name & ".log"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(folder_path) Then FSO.CreateFolder folder_path
    Logger.Log "Saving : " & file_name & ".log"
    Set stream = FSO.CreateTextFile(file_path, True)
    For i = 1 To buffer.count
      stream.WriteLine Logger.printf("{0} : {1}", buffer(i)(0), buffer(i)(1))
    Next i
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

