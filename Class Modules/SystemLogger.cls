VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event LogChanged(log_text As String)

'@exclude.json
Private FSO As Object ' Declare a FileSystemObject.
Private stream As Object ' Declare a TextStream.
Private Logs(6) As Log
Private log_level As Long
Private folder_path As String
Private file_path As String
Private log_enabled As Boolean
Private pImmediateLog As LogType

Private Sub Class_Initialize()
' Constructor
    ImmediateLog = RuntimeLog
End Sub

Public Property Get ImmediateLog() As LogType
    ImmediateLog = pImmediateLog
End Property
Private Property Let ImmediateLog(val As LogType)
    pImmediateLog = val
End Property

Public Sub SetLogLevel(level As Long)
' Sets the log level to more or less verbose output
    log_level = level
End Sub

Public Sub SetImmediateLog(log_type As LogType)
    ImmediateLog = log_type
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
        Logs(log_type).log_type = log_type
        Logs(log_type).Id = GetLogId(log_type)
        Set Logs(log_type).Buffer = New VBA.Collection
    End If
    With Logs(log_type)
        .Buffer.Add AddLine(text)
        If log_type = ImmediateLog Then
            Debug.Print printf("{0} : {1}", CStr(.Buffer(.Buffer.Count)(0)), CStr(.Buffer(.Buffer.Count)(1)))
        End If
        RaiseEvent LogChanged(text)
    End With
End Sub

Public Sub Error(function_name As String)
    Me.Log ("Error Returned From --> " & function_name), ErrorLog
End Sub

Public Sub Trace(text As String)
' Used to signify a transition point in the application log
    Me.Log "------------- " & text, RuntimeLog
End Sub

Public Sub ResetLog(Optional log_type As LogType = 0)
    ClearBuffer log_type
End Sub

Private Function AddLine(text As String) As Variant
    Dim line As Variant

    line = Array(TimeInMS, text)
    
    AddLine = line
    
End Function

Public Sub NotImplementedException()
    Me.Log "Not Implemented Exception!"
End Sub

Public Sub SaveAllLogs()
' Saves all non-empty logs
    Dim items As Variant
    items = Array(RuntimeLog, UserLog, TestLog, DebugLog, ExportLog, SqlLog, ErrorLog)
    Dim item As Variant
    For Each item In items
        SaveLog (item)
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
    Me.Log "Saving : " & file_name & ".log", log_type
    With Logs(log_type)
        Set stream = FSO.CreateTextFile(file_path, True)
        For i = 1 To .Buffer.Count
            stream.WriteLine printf("{0} : {1}", CStr(.Buffer(i)(0)), CStr(.Buffer(i)(1)))
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
        mask = Replace$(mask, "{" & i & "}", CStr(tokens(i)))
    Next
    printf = mask
End Function



