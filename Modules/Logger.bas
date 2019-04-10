Attribute VB_Name = "Logger"
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

Public Sub LogEnabled(b As Boolean)
    log_enabled = b
End Sub

Public Sub Log(text As String)
    If log_enabled = True Then
        If buffer Is Nothing Then Set buffer = CreateObject("Scripting.Dictionary")
        Do Until Not buffer.exists(TimeInMS)
            Application.Wait (Now + TimeValue("0:00:01") / 1000)
        Loop
        buffer.Add key:=TimeInMS, Item:=text
        Debug.Print Utils.printf("{0} : {1}", TimeInMS, text)
    End If
End Sub

Public Sub Trace(text As String)
' Used to signify a transition point in the application log
    Log "------------- " & text
End Sub

Public Sub NotImplementedException()
    Logger.Log "Not Implemented Exception!"
End Sub

Public Sub SaveLog(Optional file_name As String = "runtime")
    Dim key As Variant
    If buffer Is Nothing Then Exit Sub
    folder_path = ThisWorkbook.path & "\logs"
    file_path = folder_path & "\" & file_name & ".log"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(folder_path) Then FSO.CreateFolder folder_path
    Logger.Log "Saving : " & file_name & ".log"
    Set stream = FSO.CreateTextFile(file_path, True)
    For Each key In buffer
      stream.WriteLine Utils.printf("{0} : {1}", key, buffer.Item(key))
    Next key
    stream.Close
End Sub

Public Sub ClearBuffer()
  Set buffer = Nothing
End Sub

Private Function TimeInMS() As String
    TimeInMS = Strings.Format(Now, "dd-MMM-yyyy HH:nn:ss") & "." & _
               Strings.Right(Strings.Format(Timer, "#0.00"), 2)
End Function


