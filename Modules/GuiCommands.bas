Attribute VB_Name = "GuiCommands"
Option Explicit
'@Folder("Modules")
'=================================
' DESCRIPTION: Holds commands used
' through the GUI with exception
' of the import function which is
' kept in ThisWorkbook.
'=================================
Public Sub ResetExcelGUI()
' Sets visible sheets in the excel gui to only start
    HideAllSheets SAATI_Data_Manager.ThisWorkbook
End Sub

Private Sub HideAllSheets(wb As Workbook)
' Hides all visible sheets in the given workbook.
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws Is shtStart Then
            'Pass
        ElseIf ws.Visible = xlSheetVisible Then
            ws.Visible = xlSheetHidden
            App.logger.Log ws.Name & " was hidden."
        End If
    Next ws
End Sub

Public Sub ShowAllSheets(wb As Workbook)
' Makes all worksheets visible in the given workbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Next ws

End Sub

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
    SpecManager.StartApp
    shtDeveloper.Visible = xlSheetVeryHidden
    'formConsole.show vbModeless
    GoToMain
End Sub

Public Sub GoToMain()
'Opens the main menu form.
    formMainMenu.show vbModeless
End Sub

Sub UnloadAllForms()
    Dim objLoop As Object

    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop

End Sub

Public Sub ExportAll()
' Exports the codebase to a project folder as text files
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim Count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long

    lngNumberOfTasks = 4
    lngCounter = 0
    
    App.Start

    directory = ThisWorkbook.path & "\"
    


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
            err.Clear
            
            
            Call VBComponent.Export(path)
            
            If err.Number <> 0 Then
                App.logger.Log "Failed to export " & VBComponent.Name & " to " & path, ExportLog
            Else
                Count = Count + 1
                App.logger.Log "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path, ExportLog
            End If

            On Error GoTo 0
        End If

    Next
    
        
    App.logger.Log "Export Complete.", ExportLog
    App.logger.ResetLog ExportLog
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
    If Windows.Count > 1 Then
        For Each w In Windows
            If w.Parent.Name = ThisWorkbook.Name Then
                Application.Visible = True
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

Public Sub DocumentPrinterToPdf()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = PUBLIC_DIR & "\Specifications\" & App.current_spec.MaterialId & "_" & App.current_spec.Revision
    Set ws = Sheets("pdf")
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    App.logger.Log "PDF Saved : " & fileName
    Exit Sub
    
SaveFileError:
    App.logger.Log "Failed to save file PDF Fail"
End Sub

Public Sub DocumentPrinterToPdf_Test()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = PUBLIC_DIR & "\Specifications\" & App.current_spec.MaterialId & "_" & App.current_spec.Revision
    Set ws = Sheets("SpecificationForm")
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    App.logger.Log "PDF Saved : " & fileName, TestLog
    Exit Sub
    
SaveFileError:
    App.logger.Log "Failed to save file PDF Fail", TestLog
End Sub
