Attribute VB_Name = "GuiCommands"
Option Explicit
'@Folder("Modules")

'=================================
' DESCRIPTION: Holds commands used
' through the GUI with exception
' of the import function.
'=================================

Public Sub GoToMain()
'Opens the main menu form.
    SpecManager.StopSpecManager
    Application.Visible = False
    formMainMenu.Show
End Sub

Sub UnloadAllForms()
    Dim objLoop As Object

    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop
    GoToMain
End Sub

Public Sub WarpingWsToDB()
    SpecManager.StartSpecManager
    Set manager.current_template = SpecManager.GetTemplate("warping")
    SpecManager.WorksheetToDatabase
    SpecManager.StopSpecManager
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

    Logger.ClearBuffer

    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Creating a New Version...", _
        False)
        
    directory = GitRepo & "\"
    
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
    Logger.SaveLog "export"
End Sub

Public Sub ConfigControl()
'Initializes the password form for config access.
    If Environ("UserName") <> "CRuff" Then
        formPassword.Show
    Else
        Application.DisplayAlerts = True
        shtDeveloper.Visible = xlSheetVisible
        Application.Visible = True
        Application.VBE.MainWindow.Visible = True
        Application.SendKeys ("^r")
    End If
End Sub

Public Sub Open_Config(Password As String)
'Performs a password check and opens config.
    If Password = "@Wmp9296bm4ddw" Then
        Application.DisplayAlerts = True
        shtDeveloper.Visible = xlSheetVisible
        Application.Visible = True
        Application.VBE.MainWindow.Visible = True
        Application.SendKeys ("^r")
        Unload formPassword
    Else
        MsgBox "Access Denied", vbExclamation
        Exit Sub
    End If
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
    ThisWorkbook.Save
    Application.Quit
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
    Dim ws As Worksheet, initFileName As String, fileName As String
    On Error GoTo SaveFileError
    initFileName = PublicDir & "\" & manager.current_spec.MaterialId & "_" & manager.current_spec.Revision
    fileName = Application.GetSaveAsFilename(InitialFileName:=initFileName, _
                                     FileFilter:="PDF Files (*.pdf), *.pdf", _
                                     Title:="Select Path and Filename to save")
    Set ws = Sheets("SpecificationForm")
    manager.console.PrintObjectToSheet manager.current_spec, ws
    If fileName <> "False" Then
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=fileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    End If
    Exit Sub
    
SaveFileError:
    MsgBox "Failed to save file contact admin"
End Sub

Public Sub ConsoleBoxToPdf_Test()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = "C:\Users\" & Environ("Username") & "\Desktop\" & manager.current_spec.MaterialId & "_" & manager.current_spec.Revision
    Set ws = Sheets("SpecificationForm")
    manager.console.PrintObjectToSheet manager.current_spec, ws
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    Logger.Log "PDF Saved"
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail"
End Sub
