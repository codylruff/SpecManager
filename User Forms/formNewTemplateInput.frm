VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewTemplateInput 
   Caption         =   "Create New Template"
   ClientHeight    =   2952
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "formNewTemplateInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNewTemplateInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












































Private Sub UserForm_Initialize()
    Logger.Log "--------- " & Me.Name & " ----------"
    PopulateCboProductLine
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdContinue_Click()
    If SpecManager.TemplateInput(txtTemplateName.Value) <> vbNullString Then
        If cboProductLine.Value <> vbNullString Then
            App.current_template.ProductLine = cboProductLine.Value
        Else
            MsgBox "Please select a product line!"
            Exit Sub
        End If
        Unload Me
        formCreateGeneric.Show vbModeless
    Else
        MsgBox "Please enter a template name !"
        Exit Sub
    End If
End Sub

Sub Continue()
    If SpecManager.TemplateInput(txtTemplateName.Value) <> vbNullString Then
        Logger.Log "Template Input Pass"
    Else
        Logger.Log "Template Input Fail"
    End If
End Sub

Private Sub PopulateCboProductLine()
    Dim prop As Variant
    Dim i As Integer
    Do While cboProductLine.ListCount > 0
        cboProductLine.RemoveItem 0
    Loop
    With cboProductLine
        .AddItem ("Protection")
        .AddItem ("Filtration")
        .AddItem ("Chemical")
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
