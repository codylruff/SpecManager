VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTitle              As String
Private mMessage            As String
Private mDialogType         As VbMsgBoxStyle
Private mButtons()          As Object
Private mThemeBG            As String
Private mTitleForeColour    As String
Private mDuration           As Long '/duration in milliseconds
Private mIsModal            As Boolean

Private JSONRoot As Object


Public Sub Class_Initialize()
    'Default values
    Me.AddButton "Yes", vbYes, True, "#F65656", "#FFFFFF"
    Me.AddButton "No", vbNo, False
    Me.Duration = 0
    
    mIsModal = True
    mTitleForeColour = "#FFFFFF"
    mThemeBG = "#F65656"
    Set JSONRoot = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get title() As String:                      title = mTitle:                         End Property
Public Property Let title(value As String):                 mTitle = value:                         End Property
Public Property Get TitleForeColour() As String:            TitleForeColour = mTitleForeColour:     End Property
Public Property Let TitleForeColour(value As String):       mTitleForeColour = value:               End Property

'ThemeBG is used as titel background colour as well as the bg colour for the active button
Public Property Get ThemeBg() As String:                    ThemeBg = mThemeBG:                     End Property
Public Property Let ThemeBg(value As String):               mThemeBG = value:                       End Property

Public Property Get message() As String:                    message = mMessage:                     End Property
Public Property Let message(value As String):               mMessage = value:                       End Property

Public Property Get dialogType() As VbMsgBoxStyle:          dialogType = mDialogType:               End Property
Public Property Let dialogType(value As VbMsgBoxStyle):     mDialogType = value:                    End Property

Public Property Get Duration() As Long:                     Duration = mDuration:                   End Property
Public Property Let Duration(value As Long):                mDuration = value:                      End Property

Public Property Get IsModal() As Long:                     IsModal = mIsModal:                      End Property
Public Property Let IsModal(value As Long):                mIsModal = value:                        End Property

Private Function addButtonsToCollection(Button As Object)
    On Error Resume Next
    Dim C As Long
    C = UBound(mButtons)
    If C < 0 Then C = 0
    
    C = C + 1
    
    ReDim Preserve mButtons(C)
    Set mButtons(C - 1) = Button
    
End Function

Public Function ClearButtons()
'---------------------------------------------------------------------------------------
' Procedure : ClearDefaultButtons
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   : Removes all the buttons from the button collection
' Returns   :   void
'---------------------------------------------------------------------------------------
'
    
    ReDim mButtons(0)
End Function

Public Function AddButton(ButtonLabel As String, ButtonValue As Integer, Optional IsDefault As Boolean = False, Optional BgColour As String = "#FFFFFF", Optional ForeColour As String = "#808080")
'---------------------------------------------------------------------------------------
' Procedure : AddButton
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   :  Add a new button to the dialog box
' Returns   : void
'---------------------------------------------------------------------------------------

    Dim localButton As Object
    Set localButton = CreateObject("Scripting.Dictionary")
    localButton.Add "ButtonLabel", ButtonLabel
    localButton.Add "ButtonValue", ButtonValue
    localButton.Add "IsDefault", IsDefault
    localButton.Add "BGColour", BgColour
    localButton.Add "ForeColour", ForeColour
    
    addButtonsToCollection localButton
End Function

Public Function ToJSON()
'---------------------------------------------------------------------------------------
' Procedure : ToJSON
' Author    : KRISH
' Date      : 07/05/2018
' Purpose   : Returns Json string of this class
' Returns   :
'---------------------------------------------------------------------------------------

    On Error Resume Next
    JSONRoot.Add "Title", title
    JSONRoot.Add "TitleForeColour", TitleForeColour
    JSONRoot.Add "ThemeBG", ThemeBg
    JSONRoot.Add "Message", message
    JSONRoot.Add "DialogType", dialogType
    JSONRoot.Add "Buttons", mButtons
    JSONRoot.Add "Duration", Duration
    JSONRoot.Add "IsModal", IsModal
       
    ToJSON = JsonVBA.ConvertToJson(JSONRoot)
    
End Function


