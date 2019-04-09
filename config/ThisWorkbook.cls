'============================================================================
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'NOTES: This is the DEVELOPER VERSION of the Spec Manager Application.
' Excel macro to export all VBA source code in this
' project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust
' Center Settings/Macro Settings/Trust access to the VBA project object model
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'=============================================================================
Private retry As Long
Private Sub RemoveAll()
    On Error Resume Next
    Dim Element As Object
    For Each Element In ActiveWorkbook.VBProject.VBComponents
        ActiveWorkbook.VBProject.VBComponents.Remove Element
        Debug.Print Element.Name
    Next
End Sub

Sub UpdateSpecManager()
' Automatic updating from any version of Spec Manager
    Dim path As String

    ' If not up to date this sub will be called from within the initialize app procedure
    ' Remove All Modules except the update module
    Application.Run "Updater.RemovePreviousVersion"


End Sub

Sub VSImport()
    
    Dim path As String
    Dim VerNum As String
    Dim strFile As String
    Dim wb As Workbook
    
    Application.Run ("SpecManager.StopSpecManager")
    
    Set wb = ActiveWorkbook
    
    RemoveAll
    
    On Error GoTo ImportFail
    
    path = "C:\Users\cruff\source\SM - Final\Modules\"
    strFile = Dir(path & "*.bas*")
    
    Do While Len(strFile) > 0
        Debug.Print strFile
        wb.VBProject.VBComponents.Import path & strFile
        strFile = Dir
    Loop
    
    path = "C:\Users\cruff\source\SM - Final\User Forms\"
    strFile = Dir(path & "*.frm*")
    
    Do While Len(strFile) > 0
        Debug.Print strFile
        wb.VBProject.VBComponents.Import path & strFile
        strFile = Dir
    Loop
    
    path = "C:\Users\cruff\source\SM - Final\Class Modules\"
    strFile = Dir(path & "*.cls*")
    
    Do While Len(strFile) > 0
        Debug.Print strFile
        wb.VBProject.VBComponents.Import path & strFile
        strFile = Dir
    Loop
    Debug.Print "Import Successful!"
    Application.Run ("SpecManager.StartSpecManager")
    Exit Sub

ImportFail:
    MsgBox "Import Failed"
End Sub
