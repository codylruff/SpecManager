Attribute VB_Name = "DeploymentTools"
Option Explicit

Public Sub DeployVersion(new_version As String)
' Runs through all the steps necessary to properly deploy a new version of an app
' to a the AutoUpdater (AU) directory where files can be imported into all production
' versions of an application.
    Logger.ResetLog
    ' Create the include, exclude, and version JSON files and write to the AU directory
    UpdateInclude_Json
    UpdateExclude_Json
    If UpdateGlobalVersion_Json(new_version) <> 0 Then
        Logger.Log "Error incrementing global version number"
        Logger.Log "Exiting deployment subroutine"
        Exit Sub
    Else
        Updater.UpdateLocalVersion_Json new_version
        Logger.Log "Global version incremeted sucessfully."
    End If
    DeploySourceCode
    Logger.Log "New version successfully deployed!"
    Logger.ResetLog "deploy"
End Sub

Private Function UpdateInclude_Json() As Long
' Creates the include.json file to list all updateable code blocks for AU
    Dim vb_components   As Object
    Dim component       As Object
    Dim include_json    As Object
    Dim exclude_json    As Object
    Dim dir_            As String
    Dim extension       As String
    dir_ = Updater.GLOBAL_PATH & "\updates"
    ' Loop through the code modules in this project
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    Set include_json = CreateObject("Scripting.Dictionary")
    Set exclude_json = JsonVBA.GetJsonObject(dir_ & "\exclude.json")
    For Each component In vb_components
        extension = Utils.ToFileExtension(component.Type)
        If Not exclude_json.exists(component.Name) Then
            If extension <> ".txt" Then
                include_json.Add component.Name, component.Name & extension
            End If
        End If
    Next component
    ' No Errors
    UpdateInclude_Json = JsonVBA.WriteJsonObject(dir_ & "\include.json", include_json)
End Function

Private Function UpdateGlobalVersion_Json(new_version As String) As Long
' Update the global_version.json file to detail the current global app-version
    Dim global_version_json    As Object
    Dim path_                  As String
    ' file path to the global updates folder
    path_ = Updater.GLOBAL_PATH & "\updates\global_version.json"
    Set global_version_json = JsonVBA.GetJsonObject(path_)
    global_version_json.Item("app_version") = new_version
    ' Return 0 on success
    UpdateGlobalVersion_Json = JsonVBA.WriteJsonObject(path_, global_version_json)
End Function

Private Function DeploySourceCode() As Long
' Deploys the new production source code to the global updates folder
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    Dim VBComponent As Object
    Dim path As String
    Dim directory As String
    Dim extension As String
    ' initialize the updates directory
    directory = Updater.GLOBAL_PATH & "\updates"
    ' Export all .bas, .cls, and .frm code modules
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        If VBComponent.Type <> Document Then
            Select Case VBComponent.Type
                Case ClassModule
                    extension = ".cls"
                    path = directory & "\Class Modules\" & VBComponent.Name & extension
                Case Form
                    extension = ".frm"
                    path = directory & "\User Forms\" & VBComponent.Name & extension
                Case Module
                    extension = ".bas"
                    path = directory & "\Modules\" & VBComponent.Name & extension
                Case Else
                    Logger.Log "Forced to export as .txt file"
                    extension = ".txt"
            End Select
            On Error Resume Next
            Err.Clear
            VBComponent.Export (path)
            If Err.Number <> 0 Then
                Logger.Log "Failed to export " & VBComponent.Name & " to " & path
            Else
                Logger.Log "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
            End If
            On Error GoTo 0
        End If
    Next
    Logger.Log "Source Deployment Complete."
    DeploySourceCode = 0

End Function

Private Sub UpdateExclude_Json()
' Creates the exclude.json file to list non-updatedable code blocks
    Dim vb_components   As Object
    Dim component       As Object
    Dim exclude_json    As Object
    Dim dir_            As String
    Dim decorated       As Boolean
    Dim extension       As String
    dir_ = Updater.GLOBAL_PATH & "\updates"
    ' Loop through the code modules in this project
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    Set exclude_json = JsonVBA.GetJsonObject(dir_ & "\exclude.json")
    For Each component In vb_components
        ' check each code module for the decorator
        decorated = IsDecorated(component.CodeModule, "'@exclude.json")
        extension = Utils.ToFileExtension(component.Type)
        If decorated And (Not exclude_json.exists(component.Name)) Then
            If extension <> ".txt" Then
                exclude_json.Add component.Name, component.Name & extension
            End If
        End If
    Next component
End Sub

Public Function IsDecorated(code_module As Object, decorator As String) As Boolean
' Check a vba code module for a decorator comment ie. `'@decorator-example`
    IsDecorated = code_module.Find(target:=decorator, StartLine:=1, StartColumn:=1, _
                EndLine:=code_module.CountOfLines, EndColumn:=255, _
                wholeword:=True, MatchCase:=True, patternsearch:=False)
End Function
