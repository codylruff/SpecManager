Attribute VB_Name = "Updater"
'@exclude.json
' ----------------------------------------------- '
' Dependencies include : JsonVBA, LoggerVBA
' * These modules must be noted in exclude.json
'   and marked by the '@exclude.json decorator.
' ----------------------------------------------- '
Option Explicit
' IMPORT PATHS (Change these to the proper file paths)
Public Const GITREPO               As String = "C:\Users\cruff\source\SM - Final"
Public Const GLOBAL_PATH           As String = "S:\Data Manager"

' APP SETUP DESCRIPTIONS:
Public Const APP_UPDATE_AVAILABLE  As Long = 101
Public Const APP_UP_TO_DATE        As Long = 100
' UPDATER ERROR DESCRIPTIONS
Public Const MISSING_FILES         As Long = 1
Public Const TEST_FAILED           As Long = 2

Public update_available            As Boolean
Public checked_for_updates         As Boolean
Public ready_to_test               As Boolean

Private update_dir                 As String
Private modules_dir                As String
Private classes_dir                As String
Private forms_dir                  As String

Function CheckForUpdates(ByVal current_version As String) As Long
    ' Compare current app_version to the global app_version on the network global.json file.
    Dim global_version As String
    Dim local_version As String
    Debug.Print "Checking for updates . . . "
    local_version = GetLocalVersion
    global_version = GetGlobalVersion
    If global_version <> local_version Then
        update_available = "True"
        checked_for_updates = True
        CheckForUpdates = APP_UPDATE_AVAILABLE
    Else
        checked_for_updates = True
        CheckForUpdates = APP_UP_TO_DATE
    End If
End Function

Function GetGlobalVersion() As String
' Retrieve the global app version from the network version.json file
    GetGlobalVersion = GetJsonValue(GLOBAL_PATH & "\updates\global_version.json", "app_version")
End Function

Private Function GetLocalVersion() As String
' Retreieves the current app version from the version.json file
    Dim local_version_json As Object
    GetLocalVersion = JsonVBA.GetJsonValue(ThisWorkbook.path & "\config\local_version.json", "app_version")
End Function

Sub InitializeUpdater()
    Debug.Print "Loading update directories . . . "
    update_dir = GLOBAL_PATH & "\updates"
    modules_dir = update_dir & "\Modules"
    classes_dir = update_dir & "\Class Modules"
    forms_dir = update_dir & "\User Forms"
End Sub

Private Function RemovePreviousVersion() As String
' Removes file associated with the previous version of the app
    Dim log_buffer       As String
    Dim source_file      As Object
    Dim exclude_json     As Object
    Dim message          As String
    Debug.Print "Removing previous version"
    log_buffer = "Removing previous version" & vbNewLine
    On Error Resume Next
    Set exclude_json = JsonVBA.GetJsonObject(update_dir & "\exclude.json")
    For Each source_file In ThisWorkbook.VBProject.VBComponents
    ' This process ignores third-party libraries listed in the exclude.json file
        If exclude_json.exists(source_file.Name) Then
            Debug.Print "(excluded)" & source_file.Name
            log_buffer = log_buffer & "(excluded)" & source_file.Name & vbNewLine
         Else
            message = "Removed : " & source_file.Name
            If source_file.Type <> 100 Then
                source_file.Name = source_file.Name & "_OLD"
            End If
            ThisWorkbook.VBProject.VBComponents.Remove source_file
            Debug.Print message
            log_buffer = log_buffer & message & vbNewLine
        End If
    Next source_file
    RemovePreviousVersion = log_buffer
End Function

Sub ImportFromLocalGitRepo()
' Imports files from local git repo on CRUFF
    If Environ("Username") = "CRuff" Then
        InitializeUpdater
        ImportSourceCode GITREPO
    Else
        Debug.Print "Repository inaccessible from this network location."
    End If
End Sub

Public Sub UpdateSpecManager()
' Must be called from the excel GUI to prevent issues. This particular sub is specific to this application.
    CheckForUpdates current_version:=GetLocalVersion
    If update_available Then
        ' If not up to date this sub will be called from within the initialize app procedure
        InitializeUpdater
        ' Imports files associated with the new version from the network drive
        If ImportSourceCode(update_dir) = 0 Then
            ' Verify that the update was applied succesfully
            update_available = False
            UpdateLocalVersion_Json GetGlobalVersion
            ready_to_test = True
        Else
            MsgBox "The application failed to update. Contact the administrator."
            'AbortUpdateProcess
        End If
    Else
        ' Let the user know that the app is up to date.
        MsgBox "This is the newest version available."
    End If
End Sub

Private Function ImportSourceCode(ByVal import_directory As String) As Long

    Dim path As String
    Dim VerNum As String
    Dim strFile As String
    Dim wb As Workbook
    Dim log_buffer As String
    Dim exclude_json As Object
    
    Set wb = ActiveWorkbook
    ' Files must be removed before new versions can be imported. Otherwise there will be errors.
    log_buffer = RemovePreviousVersion
    
    On Error GoTo ImportFail
    Set exclude_json = JsonVBA.GetJsonObject(update_dir & "\exclude.json")
    path = import_directory & "\Modules\"
    strFile = Dir(path & "*.bas*")
    
    Do While Len(strFile) > 0
        ' Check to see if the code module should be exclude from import
        If Not exclude_json.exists(Left(strFile, Len(strFile) - 4)) Then
            log_buffer = log_buffer & path & strFile & vbNewLine
            wb.VBProject.VBComponents.Import path & strFile
            Debug.Print "Imported : " & path & strFile
        End If
        strFile = Dir
    Loop
    
    path = import_directory & "\User Forms\"
    strFile = Dir(path & "*.frm*") ' frx files must be in the same dir as the frm files
    
    Do While Len(strFile) > 0
        log_buffer = log_buffer & path & strFile & vbNewLine
        wb.VBProject.VBComponents.Import path & strFile
        Debug.Print "Imported : " & path & strFile
        strFile = Dir
    Loop
    
    path = import_directory & "\Class Modules\"
    strFile = Dir(path & "*.cls*") ' This folder must not contain sheet files or errors will occur
    
    Do While Len(strFile) > 0
        log_buffer = log_buffer & path & strFile & vbNewLine
        wb.VBProject.VBComponents.Import path & strFile
        Debug.Print "Imported : " & path & strFile
        strFile = Dir
    Loop
    Debug.Print "Source Code Import Successful!"
    log_buffer = log_buffer & "Import Successful!"
    ImportSourceCode = 0
    Exit Function

ImportFail:
    Debug.Print "Import Failed"
    ImportSourceCode = -1
End Function

Public Function VerifyUpdateIntegrity() As Long
' Checks that all source code files were imported successfully
' before attempting to intialize the test suite
    Dim source_file     As Variant
    Dim source_files    As Object
    Dim vb_components   As Object
    On Error GoTo SourceFileNotFoundException
    ' Test `Logger.bas` first to ensure results are recorded
    'source_file = "Logger.bas"
    Logger.Log "Verifying the integrity imported source code"
    ' Load include.json and use it to check for the modules which should be present.
    Set source_files = JsonVBA.GetJsonObject("S:\Data Manager\updates\include.json")
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    ' Any error will raise a SourceFileNotFoundException and fail the test
    For Each source_file In source_files
       Debug.Print CStr(vb_components.Item(source_files.Item(source_file)))
    Next source_file

    VerifyUpdateIntegrity = 0 'Return zero on success
    Exit Function

SourceFileNotFoundException:
    Debug.Print source_file & " not found. Import failed"
    VerifyUpdateIntegrity = MISSING_FILES
End Function

Public Sub UpdateLocalVersion_Json(new_value As String)
    Dim local_version_json As Object
    Set local_version_json = JsonVBA.GetJsonObject(ThisWorkbook.path & "\config\local_version.json")
    local_version_json.Item("app_version") = new_value
    JsonVBA.WriteJsonObject ThisWorkbook.path & "\config\local_version.json", local_version_json
End Sub

Public Sub UpdateAvailablePrompt()
	' If update.ps1 script is executed the all excel workbooks will be closed.
    ' This means we must allow the user the chance to save their work 
End Sub