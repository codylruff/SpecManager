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
        ElseIf ws Is shtPlanning Then
            'Pass
        ElseIf ws.Visible = xlSheetVisible Then
            ws.Visible = xlSheetHidden
            Logger.Log ws.Name & " was hidden."
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
    shtAdmin.Visible = xlSheetVeryHidden
    GoToMain
End Sub

Public Sub GoToMain()
'Opens the main menu form.

End Sub

Sub UnloadAllForms()
    Dim objLoop As Object

    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop

End Sub

Public Sub ExportAll()
' Exports the codebase to a project folder as text files
    Const vbaModule = 1
    Const vbaClassModule = 2
    Const vbaForm = 3
    Const vbaDocument = 100
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
    
    'App.Start

    directory = ThisWorkbook.path & "\src\"
    
    Logger.ResetLog ExportLog
    Logger.SetLogLevel LOG_ALL
    Logger.SetImmediateLog ExportLog
    Logger.Log "Exporting Files . . . ", RuntimeLog
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        
        If VBComponent.Type <> vbaDocument Then
            Select Case VBComponent.Type
                Case vbaClassModule
                    extension = ".cls"
                    path = directory & VBComponent.Name & extension
                Case vbaForm
                    extension = ".frm"
                    path = directory & VBComponent.Name & extension
                    
                Case vbaModule
                    extension = ".bas"
                    path = directory & VBComponent.Name & extension
                    
                Case Else
                    extension = ".txt"
            End Select
            
            On Error Resume Next
            err.Clear
            
            
            Call VBComponent.Export(path)
            
            If err.Number <> 0 Then
                Logger.Log "Failed to export " & VBComponent.Name & " to " & path, ExportLog
            Else
                Count = Count + 1
                Logger.Log "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path, ExportLog
            End If

            On Error GoTo 0
        End If

    Next
    
    Logger.Log "Export Complete", RuntimeLog
    Logger.Log "Export Complete.", ExportLog
    Logger.SetImmediateLog RuntimeLog
    Logger.SaveLog ExportLog
    Logger.SetLogLevel LOG_LOW
End Sub

Public Sub CloseConfig()
'Performs actions needed to close config.
    ThisWorkbook.Save
    shtAdmin.Visible = xlSheetVeryHidden
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
                ctl.text = nullstr
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
    fileName = PUBLIC_DIR & "\Specifications\" & App.current_doc.MaterialId & "_" & App.current_doc.Revision
    Set ws = Sheets("pdf")
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    Logger.Log "PDF Saved : " & fileName
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail"
End Sub

Public Sub DocumentPrinterToPdf_Test()
    Dim ws As Worksheet
    Dim fileName As String
    On Error GoTo SaveFileError
    fileName = PUBLIC_DIR & "\Specifications\" & App.current_doc.MaterialId & "_" & App.current_doc.Revision
    Set ws = Sheets("DocumentForm")
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    Logger.Log "PDF Saved : " & fileName, TestLog
    Exit Sub
    
SaveFileError:
    Logger.Log "Failed to save file PDF Fail", TestLog
End Sub

Public Sub OpenConfiguration()
' Opens the configuration panel
    App.Start
    AccessControl.ConfigControl
End Sub

Public Sub GoToPlanning()
' Opens planning worksheet
    shtPlanning.Activate
End Sub

Public Sub CopyCurrentDocument()
' Makes a copy of the current doc, with a new material id
    Dim new_material_id As String
    Dim ret_val As Long
    new_material_id = Prompt.UserInput(SingleLineText, "Material Id", "Enter a material id for copy?")
    If new_material_id = nullstr Then
        Prompt.Error "You must enter a material id."
        Exit Sub
    End If
    ret_val = SpecManager.CreateDocumentFromCopy(App.current_doc, new_material_id)
    If ret_val = DB_PUSH_SUCCESS Then
        Prompt.Success "Copied Successfully"
    Else
        Prompt.Error "Copy Failed"
    End If
End Sub

Public Sub LoadExcelDocument()
    DocumentParser.LoadNewDocument "excel"
End Sub

Public Sub LoadJsonDocument()
    DocumentParser.LoadNewDocument "json"
End Sub

Public Sub CreateBallisticsDocument()
    Dim material_id As String
    Dim package_length_inches As Double
    Dim fabric_width_inches As Double
    Dim conditioned_weight_gsm As Double
    Dim target_psf As Double
    Dim ret_val As Long
    Dim machine_id As String

    App.Start
    material_id = shtAdmin.Range("material_id").value ' this is the material id (SAP Code)
    package_length_inches = shtAdmin.Range("package_length_inches")
    fabric_width_inches = shtAdmin.Range("fabric_width_inches")
    conditioned_weight_gsm = shtAdmin.Range("conditioned_weight_gsm")
    target_psf = shtAdmin.Range("target_psf")
    machine_id = CStr(shtAdmin.Range("machine_id").value)   ' This is the machine id (ie. loom number, warper, etc...)

    ret_val = SpecManager.BuildBallisticTestSpec(material_id, package_length_inches, fabric_width_inches, conditioned_weight_gsm, target_psf, machine_id, False)
    
    ' Parse return value.
    If ret_val = DB_PUSH_SUCCESS Then
        Prompt.Success "New Document Saved."
    ElseIf ret_val = MATERIAL_EXISTS_ERR Then
        Prompt.Error "Material Already Exists."
    Else
        Prompt.Error "Error Saving Document."
    End If
    
    App.Shutdown
End Sub
