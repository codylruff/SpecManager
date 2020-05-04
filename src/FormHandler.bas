Attribute VB_Name = "FormHandler"
Option Explicit
'===================================
'DESCRIPTION: FormHandler
'===================================
Function ClearForm(name As String) As Boolean
    Dim current_form As IForm
    On Error GoTo catchError
    Set current_form = App.forms.Item(name)
    current_form.Clear
exitFunction:
    Set current_form = Nothing
    Exit Function
catchError:
    Logger.Log "Form not found."
    GoTo exitFunction
End Function