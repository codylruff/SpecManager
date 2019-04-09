Attribute VB_Name = "Updater"
Option Explicit

Public app_version          As String
Public update_available     As Boolean
Private update_dir          As String
Private modules_dir         As String
Private classes_dir         As String
Private forms_dir           As String

Function CheckForUpdates(current_version As String) As Long
    ' Compare current app_version to the global app_version on the network global.json file.
    Dim global_version As String
    Dim local_version As String
    local_version = current_version
    global_version = GetGlobalVersion
    If global_version <> local_version Then
        CheckForUpdates = APP_UPDATE_AVAILABLE
    Else
        CheckForUpdates = APP_UP_TO_DATE
    End If
End Function

Function GetGlobalVersion() As String
' Retrieve the global app version from the network version.json file
    Dim FSO As Object
    Dim JsonTS As Object
    Dim jsonText As String
    Dim Parsed As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' Read .json file
    Set JsonTS = FSO.OpenTextFile(PublicDir & "\version.json", 1)
    jsonText = JsonTS.ReadAll
    JsonTS.Close
    ' Parse json to Dictionary
    Set Parsed = JsonConverter.ParseJson(jsonText)
    GetGlobalVersion = Parsed.Item("app_version")
End Function

Sub InitializeUpdater()
    Logger.Log "Loading update directories . . . "
    update_dir = PublicDir & "\updates"
    modules_dir = update_dir & "\Modules"
    classes_dir = update_dir & "\Class Modules"
    forms_dir = update_dir & "\User Forms"
    Logger.Log "Ready to update."
    RemovePreviousVersion
End Sub

Sub RemovePreviousVersion()
' Removes file associated with the previous version of the app
    On Error Resume Next
    Dim element A Object
    For Each element In ActiveWorkbook.VBProject.VBComponents
        If element.Name <> "Updater.bas" Then ' *Excludes the updater itself
            ActiveWorkbook.VBProject.VBComponents.Remove element
            Debug.Print element.Name
        End If
    Next element
End Sub

Sub ApplyUpdate()
' Imports files associated with the new version from the network drive

End Sub

