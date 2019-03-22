VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Progress"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




















Option Explicit

Private Sub cmdAbort_Click()
    'Unload the form
    Unload ufProgress
    'And stop execution of all codes
    End
End Sub

Sub UpdateForm(ByVal ActionNumber As Long, ByVal TotalActions As Long, _
                Optional ByVal StatusMessage As String = vbNullString)
    'Declare Sub level Variables
    Dim strStatus As String
    'Preparing the Status Message
    strStatus = Format(ActionNumber, String(Len(CStr(TotalActions)), "0")) & " of " & TotalActions
    If Not StatusMessage = vbNullString Then
        strStatus = Left(strStatus & " | " & StatusMessage, 80) & "..."
    End If
    'Set the Status Message
    ufProgress.lblStatus.Caption = strStatus
    'Set the Proportion of actions completed
    ufProgress.lblPct.Caption = CStr(CLng(ActionNumber * 100 / TotalActions)) & "%"
    'Resize the Label Control
    ufProgress.lblFront.Width = ufProgress.lblBase.Width * (ActionNumber / TotalActions)
    'Check of all the actions have been completed
    If ActionNumber = TotalActions Then
        ufProgress.cmdAbort.Caption = "Close"
        ufProgress.lblStatus.Caption = "Complete. Press Close to exit."
    End If

End Sub

