Attribute VB_Name = "Updater"
' ----------------------------------------------- '
' Dependencies include : JsonVBA, LoggerVBA
' * These modules should be noted in exclude.json
' ----------------------------------------------- '
Option Explicit
' IMPORT PATHS (Change these to the proper file paths)
Public Const GITREPO               As String = "C:\Users\cruff\source\SM - Final"
Public Const GLOBAL_PATH           As String = "S:\Data Manager\"
Public Const WORKBOOK_NAME         As String = "Spec Manager (Test).xlsm"
' APP SETUP DESCRIPTIONS:
Public Const APP_UPDATE_AVAILABLE  As Long = 101
Public Const APP_UP_TO_DATE        As Long = 100
' UPDATER ERROR DESCRIPTIONS
Public Const MISSING_FILES         As Long = 1
Public Const TEST_FAILED           As Long = 2

Public update_available            As Boolean
Public checked_for_updates         As Boolean
Private update_dir                 As String
Private modules_dir                As String
Private classes_dir                As String
Private forms_dir                  As String

Function CheckForUpdates(ByVal current_version As String) As Long
    ' Compare current app_version to the global app_version on the network global.json file.
    Dim global_version As String
    Dim local_version As String
    Logger.Log "Checking for updates . . . "
    local_version = current_version
    global_version = GetGlobalVersion
    If CDbl(global_version) > CDbl(local_version) Then
        Updater.update_available = "True"
        CreateNewUpdateAlert
        checked_for_updates = True
        CheckForUpdates = APP_UPDATE_AVAILABLE
    Else
        checked_for_updates = True
        CheckForUpdates = APP_UP_TO_DATE
    End If
End Function

Public Sub CreateNewUpdateAlert()
    Dim btn As Object
    For Each btn In shtStart.Buttons
        If btn.Name = "Button 4" Then
            btn.text = "Update Available"
        End If
    Next btn
End Sub

Function GetGlobalVersion() As String
' Retrieve the global app version from the network version.json file
    GetGlobalVersion = GetJsonValue(GLOBAL_PATH & "\updates\version.json", "app_version")
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
    log_buffer = "Removing previous version" & vbNewLine
    On Error Resume Next
    Set exclude_json = JsonVBA.GetJsonObject(update_dir & "\exclude.json")
    For Each source_file In ActiveWorkbook.VBProject.VBComponents
    ' This process ignores third-party libraries listed in the exclude.json file
        If exclude_json.exists(source_file.Name) Then
            log_buffer = log_buffer & "(excluded)" & source_file.Name & vbNewLine
         Else
            ActiveWorkbook.VBProject.VBComponents.Remove source_file
            Debug.Print source_file.Name
            log_buffer = log_buffer & source_file.Name & vbNewLine
        End If
    Next source_file
    RemovePreviousVersion = log_buffer
End Function

Sub ApplyUpdate()
' Imports files associated with the new version from the network drive
    ImportSourceCode update_dir
End Sub

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
    SpecManager.StartApp
    If checked_for_updates = False Then
        CheckForUpdates App.version
    End If
    If update_available Then
        On Error GoTo UpdateFailedException
        ' If not up to date this sub will be called from within the initialize app procedure
        InitializeUpdater
        ' Remove old version and import new source code files
        ' Destroy all objects except Updater.bas to prevent errors
        SpecManager.StopApp
        ApplyUpdate
        ' Verify that the update was applied succesfully
        If VerifyUpdateIntegrity <> MISSING_FILES Then
            update_available = False
            App.current_user.ChangeSetting "app_version", GetGlobalVersion
            SpecManager.StartApp
            Tests.AllTests
        End If
    Else
        ' Let the user know that the app is up to date.
        MsgBox "This is the newest version available."
    End If
    
    Exit Sub
UpdateFailedException:
    MsgBox "The application failed to update. Contact the administrator."
    AbortUpdateProcess
End Sub

Private Sub AbortUpdateProcess()
' Aborts the update process in order to preserve the application upon update failure
    Dim w As Window
    If Windows.count > 1 Then
        Application.DisplayAlerts = False
        For Each w In Windows
            If w.Parent.Name = WORKBOOK_NAME Then
                w.Parent.Close
            End If
        Next w
        Application.DisplayAlerts = True
    Else
        Application.DisplayAlerts = False
        Application.Quit
    End If
End Sub

Private Sub ImportSourceCode(ByVal import_directory As String)

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
        Debug.Print strFile
        Debug.Print Left(strFile, Len(strFile) - 4)
        If Not exclude_json.exists(Left(strFile, Len(strFile) - 4)) Then
            log_buffer = log_buffer & path & strFile & vbNewLine
            wb.VBProject.VBComponents.Import path & strFile
        End If
        strFile = Dir
    Loop
    
    path = import_directory & "\User Forms\"
    strFile = Dir(path & "*.frm*") ' frx files must be in the same dir as the frm files
    
    Do While Len(strFile) > 0
        Debug.Print strFile
        log_buffer = log_buffer & path & strFile & vbNewLine
        wb.VBProject.VBComponents.Import path & strFile
        strFile = Dir
    Loop
    
    path = import_directory & "\Class Modules\"
    strFile = Dir(path & "*.cls*") ' This folder must not contain sheet files or errors will occur
    
    Do While Len(strFile) > 0
        Debug.Print strFile
        log_buffer = log_buffer & path & strFile & vbNewLine
        wb.VBProject.VBComponents.Import path & strFile
        strFile = Dir
    Loop
    'Logger.Log log_buffer & "Import Successful!"
    'Logger.SaveLog "import"
    Exit Sub

ImportFail:
    MsgBox "Import Failed"
End Sub

Private Function VerifyUpdateIntegrity() As Long
' Checks that all source code files were imported successfully
' before attempting to intialize the test suite
    Dim source_file     As Variant
    Dim source_files    As Object
    Dim vb_components   As Object
    On Error GoTo SourceFileNotFoundException
    ' Test `Logger.bas` first to ensure results are recorded
    source_file = "Logger.bas"
    Logger.Log "Verifying the integrity imported source code"
    ' Load include.json and use it to check for the modules which should be present.
    Set source_files = JsonVBA.GetJsonObject(update_dir & "\include.json")
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

