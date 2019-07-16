Attribute VB_Name = "Logger"
'@exclude.json
Option Explicit
Private FSO As Object ' Declare a FileSystemObject.
Private stream As Object ' Declare a TextStream.
Private Logs(6) As Log
Private log_level As Long
Private folder_path As String
Private file_path As String
Private log_enabled As Boolean

Public Sub SetLogLevel(level As Long)
' Sets the log level to more or less verbose output
    log_level = level
End Sub

Private Function GetLogId(log_type As LogType) As String
' Convert from log type to log id string
    Select Case log_type
        Case 0
            GetLogId = "runtime"
        Case 1
            GetLogId = "user"
        Case 2
            GetLogId = "test"
        Case 3
            GetLogId = "debug"
        Case 4
            GetLogId = "export"
        Case 5
            GetLogId = "sql"
        Case 6
            GetLogId = "error"
    End Select
End Function

Public Sub Log(text As String, Optional log_type As LogType = 0)
    If Logs(log_type).Buffer Is Nothing Then 
        Logs(log_type).Log_Type = log_type
        Logs(log_type).Id = GetLogId(log_type)
        Set Logs(log_type).Buffer = New VBA.Collection 
    End If
    With Logs(log_type)
        .Buffer.Add AddLine(text)
        If log_type = RuntimeLog Then
            Debug.Print Logger.printf("{0} : {1}", .Buffer(.Buffer.Count)(0), .Buffer(.Buffer.Count)(1))
        End If
    End With
End Sub

Public Sub Error(function_name As String)
    Logger.Log ("Error Returned From --> " & function_name), ErrorLog
End Sub

Public Sub Trace(text As String)
' Used to signify a transition point in the application log
    Log "------------- " & text, RuntimeLog
End Sub

Public Sub ResetLog(Optional log_type As LogType = 0)
    Logger.SaveLog log_type
    Logger.ClearBuffer log_type
End Sub

Private Function AddLine(text As String) As Variant
    Dim line As Variant

    line = Array(TimeInMS, text)
    
    AddLine = line
    
End Function

Public Sub NotImplementedException()
    Logger.Log "Not Implemented Exception!"
End Sub

Public Sub SaveAllLogs()
' Saves all non-empty logs
    Dim items As Variant
    items = Array(RuntimeLog, UserLog, TestLog, DebugLog, ExportLog, SqlLog, ErrorLog)
    Dim item As Variant
    For Each item In items
        SaveLog(item)
    Next
End Sub

Public Sub SaveLog(Optional log_type As LogType = 0)
    Dim line As Variant
    Dim i As Long
    Dim file_name As String
    file_name = Logs(log_type).Id
    If Logs(log_type).Buffer Is Nothing Then Exit Sub
    folder_path = ThisWorkbook.path & "\logs"
    file_path = folder_path & "\" & file_name & ".log"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(folder_path) Then FSO.CreateFolder folder_path
    Logger.Log "Saving : " & file_name & ".log", log_type
    With Logs(log_type)
        Set stream = FSO.CreateTextFile(file_path, True)
        For i = 1 To .Buffer.Count
            stream.WriteLine Logger.printf("{0} : {1}", .Buffer(i)(0), .Buffer(i)(1))
        Next i
        stream.Close
    End With

End Sub

Private Sub ClearBuffer(log_id As LogType)
  Set Logs(log_id).Buffer = New VBA.Collection
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

