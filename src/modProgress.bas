Attribute VB_Name = "modProgress"
Option Explicit ' Always a good idea to use this
'@Folder("Modules")
'========================================================================
'ShowProgress : Macro that displays a progress bar. It needs 'ufProgress'
'to operate.
'Arguments:
'ActionNumber   - This is where you set the number of actions that have
'been executed. You may probably use a For-Loop counter here, or
'explicitly change it as the code progresses.
'TotalActions   - Here is where you tell the form how many actions are
'going to be performed, so it can calculate the proportion of actions
'completed.
'Title          - You may choose to set a custom title for the form here
'StatusMessage  - The form has a little status bar at the bottom and you
'can display custom messages there. Ideally, a short description of the
'action that is currently being performed.
'CloseWhenDone  - If this argument is set to True, the Form closes
'automatically when ActionNumber reaches TotalActions. Hence, remember to
'call ShowProgress with the last ActionNumber only when that action is
'complete.
'========================================================================
'Author     :   Ejaz Ahmed
'Date       :   27 March 2014
'Website    :   http://strugglingtoexcel.wordpress.com/
'Email      :   StrugglingToExcel@outlook.com
'========================================================================
Sub ShowProgress(ByVal ActionNumber As Long, _
                ByVal TotalActions As Long, _
                Optional ByVal StatusMessage As String = nullstr, _
                Optional ByVal CloseWhenDone As Boolean = True, _
                Optional ByVal title As String = nullstr)

    DoEvents 'to ensure that the code to display the form gets executed

    'Display the Proressbar
    If isFormOpen("ufProgress") Then
        
        'If the form is already open, just update the ActionNumbers and Status
        'message
        Call ufProgress.UpdateForm(ActionNumber, TotalActions, StatusMessage)

    Else
        
        'Center the form to the application window before showing it.
        'This was updated in V1.02 after receiving feedback from many
        'multi-monitor users
        Call CenterUserForm(ufProgress)
        
        'if the form is not already open, Show it
        ufProgress.show
        'set the title
        If Not title = nullstr Then
            ufProgress.Caption = title
        End If
        'then update the ActionNumber and Status Message
        Call ufProgress.UpdateForm(ActionNumber, TotalActions, StatusMessage)

    End If

    'If the user chose to close the form automatically when the last action
    'is reached, close it
    If CloseWhenDone And CBool(ActionNumber >= TotalActions) Then
        Unload ufProgress
    End If

End Sub

'========================================================================
'isFormOpen: A function to check if a form has already being showed. The
'function returns True if a form with a specified name is already open.
'Arguments:
'FormName   - Name of the form that needs to be checked
'========================================================================
'Author     :   Ejaz Ahmed
'Date       :   27 March 2014
'Website    :   http://strugglingtoexcel.wordpress.com/
'Email      :   StrugglingToExcel@outlook.com
'========================================================================
Function isFormOpen(ByVal FormName As String) As Boolean

    'Declare Function level Objects
    Dim ufForm As Object

    'Set the Function to False
    isFormOpen = False

    'Loop through all the open forms
    For Each ufForm In VBA.UserForms
        'Check the form names
        If ufForm.Name = FormName Then
            'if the form is open, set the function value to True
            isFormOpen = True
            'and exit the loop
            Exit For
        End If
    Next ufForm

End Function

'========================================================================
'CenterUserForm: A sub to center a userform to the application window.
'Arguments:
'WhichForm   - The Form Object
'========================================================================
'Author     :   Ejaz Ahmed
'Date       :   21 Feb 2018
'Website    :   http://strugglingtoexcel.wordpress.com/
'Email      :   StrugglingToExcel@outlook.com
'========================================================================
Sub CenterUserForm(ByRef WhichForm As Object)
    'This property lets you manually set the start up position
    WhichForm.StartUpPosition = 0
    'Center the form
    WhichForm.Left = Application.Left + (Application.Width / 2) - (WhichForm.Width / 2)
    WhichForm.Top = Application.Top + (Application.Height / 2) - (WhichForm.Height / 2)
End Sub
