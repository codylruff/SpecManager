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
' Creates the include.json file to support AU functionality
    Dim vb_components   As Object
    Dim component       As Object
    Dim include_dict    As Object
    ' Loop through the code modules in this project
    Set vb_components = ActiveWorkbook.VBProject.VBComponents
    Set include_dict = CreateObject("Scripting.Dictionary")
    For Each component In vb_components
        include_dict.Add component.Name, component.Name & "." & Type(component) 'Is this how to get the file extension???
    Next component
End Function