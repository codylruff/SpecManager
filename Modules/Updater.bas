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
Private update_dir                 As String
Private modules_dir                As String
Private classes_dir                As String
Private forms_dir                  As String

Function CheckForUpdates(ByVal current_version As String) As Long
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
    GetGlobalVersion = GetJsonValue(PublicDir & "\updates\version.json", "app_version")
End Function

Sub InitializeUpdater()
    Logger.Log "Loading update directories . . . "
    update_dir = GLOBAL_PATH & "\updates"
    modules_dir = update_dir & "\Modules"
    classes_dir = update_dir & "\Class Modules"
    forms_dir = update_dir & "\User Forms"
    RemovePreviousVersion
End Sub

Private Function RemovePreviousVersion() As String
' Removes file associated with the previous version of the app
    Dim log_buffer       As String
    Dim source_file      As Object
    Dim exclude_json     As Object
    Dim source_files
    log_buffer = "Removing previous version" & vbNewLine
    On Error Resume Next
    Set exclude_json = JsonVBA.GetJsonObject(update_dir & "\exclude.json")
    For Each source_file In ActiveWorkbook.VBProject.VBComponents
    ' This process ignores third-party libraries listed in the exclude.json file
        If Not exclude_json.Exists(source_file.Name) Then 
            ActiveWorkbook.VBProject.VBComponents.Remove source_file
            Debug.Print source_file.Name
            log_buffer = log_buffer & source_file.Name & vbNewLine
        Else
            log_buffer = log_buffer & "(excluded)" & source_file.Name  & vbNewLine
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
        ImportSourceCode GITREPO
    Else
        Logger.Log "Repository inaccessible from this network location."
    End If
End Sub

Sub UpdateSpecManager()
' Must be called from the excel GUI to prevent issues. This particular sub is specific to this application.
    If update_available Then
        On Error Goto UpdateFailedException
        ' If not up to date this sub will be called from within the initialize app procedure
        InitializeUpdater
        ' Remove old version and import new source code files
        ' Destroy all objects except Updater.bas to prevent errors
        SpecManager.StopSpecManager
        ApplyUpdate
        SpecManager.StartSpecManager
        ' Verify that the update was applied succesfully
        If VerifyUpdateIntegrity <> MISSING_FILES Then
            Tests.AllTests
        End If
    Else
        ' Let the user know that the app is up to date.
        MsgBox "This is the newest version availible."
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
    
    Set wb = ActiveWorkbook
    ' Files must be removed before new versions can be imported. Otherwise there will be errors.
    log_buffer = RemovePreviousVersion
    
    On Error GoTo ImportFail

    path = import_directory & "\Modules\"
    strFile = Dir(path & "*.bas*")
    
    Do While Len(strFile) > 0
        If strFile <> "Updater.bas" Then
            Debug.Print strFile
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
    Logger.Log log_buffer & "Import Successful!"
    Logger.SaveLog "import"
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
    On Error Goto SourceFileNotFoundException
    ' Test `Logger.bas` first to ensure results are recorded
    source_file = "Logger.bas"
    Logger.Log "Verifying the integrity imported source code"
    ' Load include.json and use it to check for the modules which should be present.
    Set source_files = JsonVBA.JsonVBA.GetJsonObject(update_dir & "\include.json")
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    ' Any error will raise a SourceFileNotFoundException and fail the test
    For Each source_file In source_files    
       Logger.Log CStr(vb_components.Item(source_files.Item(source_file)))
    Next source_file

    VerifyUpdateIntegrity = 0 'Return zero on success
    Exit Function

SourceFileNotFoundException:
    Debug.Print source_file & " not found. Import failed"
    VerifyUpdateIntegrity = MISSING_FILES
End Function