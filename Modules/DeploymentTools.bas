Public Sub DeployNewProductionVersion()
' Runs through all the steps necessary to properly deploy a new version of an app
' to a the AutoUpdater (AU) directory where files can be imported into all production
' versions of an application.

    ' Create the include, exclude, and version JSON files and write to the AU directory
    CreateInclude_Json
    CreateExclude_Json
    CreateVersion_Json
End Sub

Private Function CreateInclude_Json() As Long
' Creates the include.json file to list all updateable code blocks for AU
    Dim vb_components   As Object
    Dim component       As Object
    Dim include_dict    As Object
    ' Loop through the code modules in this project
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    Set include_dict = CreateObject("Scripting.Dictionary")
    For Each component In vb_components
        include_dict.Add component.Name, component.Name & "." & Type(component) '<--- Is this how to get the file extension???
    Next component
    ' No Errors
    CreateInclude_Json = 0
End Function

Private Function CreateExclude_Json() As Long
' Creates the exclude.json file to list non-updatedable code blocks
    Dim exclude_dict    As Object
    Set exclude_dict = CreateObject("Scripting.Dictionary")
    ' No Errors
    CreateExclude_Json = 0
End Function

Private Function CreateVersion_Json() As Long
' Create the version.json file to detail the current global app-version
    Dim version_dict    As Object
    Set version_dict = CreateObject("Scripting.Dictionary")
    ' No Errors
    CreateVersion_Json = 0
End Function