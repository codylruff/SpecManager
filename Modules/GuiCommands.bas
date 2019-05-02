Attribute VB_Name = "GuiCommands"
Option Explicit
'@Folder("Modules")

'=================================
' DESCRIPTION: Holds commands used
' through the GUI with exception
' of the import function.
'=================================
Public Sub DeinitializeApplication()
    SpecManager.StopApp
    If Application.VBE.MainWindow.Visible = True Then
        Application.VBE.MainWindow.Visible = False
    End If
    If Application.DisplayAlerts = False Then
        Application.DisplayAlerts = True
    End If
End Sub

Public Sub InitializeApplication()
    If update_available Then
        MsgBox "Please update the application to the current version."
        Exit Sub
    End If
    SpecManager.StartApp
    On Error GoTo UpdateFailure
    ' Check for updates and start up the app if it is up to date
    Updater.UpdateSpecManager
    If update_available Then Exit Sub
    On Error GoTo 0
    ' If the app is updated and you have already checked for updates the app will start.
    SpecManager.StartApp
    shtDeveloper.Visible = xlSheetVeryHidden
    GoToMain
    Exit Sub
UpdateFailure:
    Logger.Log "Update failed"
    MsgBox "Update Failed Contact Administrator!"
End Sub


Public Sub GoToMain()
'Opens the main menu form.
    formMainMenu.Show vbModeless
End Sub

Sub UnloadAllForms()
    Dim objLoop As Object

    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop

End Sub

Public Sub WarpingWsToDB()
    SpecManager.RestartApp
    Set App.current_template = SpecManager.GetTemplate("warping")
    SpecManager.WorksheetToDatabase
End Sub

Public Sub ExportAll()
' Exports the codebase to a project folder as text files
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long

    lngNumberOfTasks = 4
    lngCounter = 0

    Logger.ResetLog

    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Creating a New Version...", _
        False)
    'If App.current_user.Settings.Item("repo_path") = vbNullString Then
        directory = "C:\Users\cruff\source\SM - Final\"
    'Else
        'directory = "S:\Data Manager\updates\"
    'End If
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        1, _
        lngNumberOfTasks, _
        "Saving...", _
        False, _
        "Spec Manager")
    
    count = 0
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Creating Directory...", _
        False)
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Exporting Code Modules...", _
        False)

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        
        If VBComponent.Type <> Document Then
            Select Case VBComponent.Type
                Case ClassModule
                    extension = ".cls"
                    path = directory & "Class Modules\" & VBComponent.Name & extension
                Case Form
                    extension = ".frm"
                    path = directory & "User Forms\" & VBComponent.Name & extension
                    
                Case Module
                    extension = ".bas"
                    path = directory & "Modules\" & VBComponent.Name & extension
                    
                Case Else
                    extension = ".txt"
            End Select
            
            On Error Resume Next
            Err.Clear
            
            
            Call VBComponent.Export(path)
            
            If Err.Number <> 0 Then
                Logger.Log "Failed to export " & VBComponent.Name & " to " & path
            Else
                count = count + 1
                Logger.Log "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
            End If

            On Error GoTo 0
        End If

    Next
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Finishing...", _
        True)
        
    Logger.Log "Export Complete."
    Logger.ResetLog "export"
End Sub

Public Sub CloseConfig()
'Performs actions needed to close config.
    ThisWorkbook.Save
    shtDeveloper.Visible = xlSheetVeryHidden
    Application.VBE.MainWindow.Visible = False
    Application.DisplayAlerts = False
    GuiCommands.GoToMain
End Sub

Public Sub ExitApp()
'This exits the application after saving the thisworkbook.
    Dim w As Window
    SpecManager.StopApp
    If Windows.count > 1 Then
        For Each w In Windows
            If w.Parent.Name = ThisWorkbook.Name Then
                w.Parent.Save
                w.Parent.Close
            End If
        Next w
        If Application.DisplayAlerts = False Then
            Application.DisplayAlerts = True
        End If
    Else
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub

Public Sub ClearForm(frm)
'Clears the values from a user form.
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                ctl.text = vbNullString
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
            Case Else
                End Select
    Next ctl
End Sub

Public Sub DB2W_tblWarpingSpecs()
' Dumps warping specs to a new worksheet
    
End Sub

Public Sub DB2W_tblStyleSpecs()
' Dumps style specs to a new worksheet
    
End Sub

Public Sub ConsoleBoxToPdf()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = PublicDir & "\Specifications\" & App.current_spec.MaterialId & "_" & App.current_spec.Revision
    Set ws = Sheets("SpecificationForm")
    App.console.PrintObjectToSheet App.current_spec, ws
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    Logger.Log "PDF Saved : " & fileName
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail"
End Sub

Public Sub ConsoleBoxToPdf_Test()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = PublicDir & "\Specifications\" & App.current_spec.MaterialId & "_" & App.current_spec.Revision
    Set ws = Sheets("SpecificationForm")
    App.console.PrintObjectToSheet App.current_spec, ws
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    Logger.Log "PDF Saved : " & fileName
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail"
End Sub
