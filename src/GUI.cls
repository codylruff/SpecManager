VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'NOTE This is where the GUI should be handled.
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Type Settings
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    DisplayStatusBar As Boolean
    Calculation As Long
    LastSheetUsed As String
End Type

' Events
'REVIEW These events are not used at all. Why are they here?
Public Event BeforeRefresh()
Public Event AfterRefresh()
Public Event BeforeStart()
Public Event AfterStart()
Public Event BeforeShutdown()

Public kDll                      As New KrishDll
Public gMsg                      As String
Public gAppGuid                  As String
Public forms                     As VBA.Collection
Public PerformanceModeEnabled    As Boolean
Private settings_                As Settings
Public IsRunning                As Boolean

Public Property Get Krish() As KrishDll
' Wrapper for gDll krishDll library handle
    Set Krish = kDll
End Property

Sub Start()
    ' Set up forms for application
    Set forms = SetUpForms
    GUI.IsRunning = True
End Sub

Sub Shutdown()
    Set forms = Nothing
    GUI.IsRunning = False
End Sub

Public Sub Restart()
    Shutdown
    Start
End Sub

Public Sub ActivateForm(frm_name As String)
' Activates the specified form for navigation
    On Error GoTo Catch
    If Not GUI.IsRunning Then
        Start
    End If
    On Error GoTo 0
    GoTo Finally
Catch:
    Start
Finally:
    GUI.GetForm(frm_name).Sheet.Activate
End Sub

Public Sub ResetExcelGUI()
' Sets visible sheets in the excel gui to only start
    HideAllSheets SAATI_Data_Manager.ThisWorkbook
    ThisWorkbook.Sheets(settings_.LastSheetUsed).Visible = xlSheetVisible
    ThisWorkbook.Sheets(settings_.LastSheetUsed).Activate
End Sub

Private Sub HideAllSheets(wb As Workbook)
' Hides all visible sheets in the given workbook.
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws Is shtNavigation Then
            'Pass
        ElseIf ws Is shtPlanning Then
            'Pass
        ElseIf ws Is shtPortal Then
            'Pass
        ElseIf ws.Visible = xlSheetVisible Then
            ws.Visible = xlSheetHidden
            Logger.Log ws.Name & " was hidden."
        End If
    Next ws
End Sub

Public Function GetForm(frm_name As String) As IForm
' Returns the form object requested
    Set GetForm = Me.forms.item(frm_name)
End Function

Private Function SetUpForms() As VBA.Collection
' Prepare forms for use
    Dim coll As VBA.Collection
    Set coll = New VBA.Collection
    ' REVIEW Is this the right way to load the forms?
    With coll
        .Add Factory.CreateFormCreate, "FormCreate"
        .Add Factory.CreateFormPortal, "FormPortal"
        .Add Factory.CreateFormNavigation, "FormNavigation"
        '.Add Factory.CreateFormEdit, "FormEdit"
        '.Add Factory.CreateFormView, "FormView"
        .Add Factory.CreateFormPlanning, "FormPlanning"
        '.Add Factory.CreateFiltrationPlanningForm, "Filtration Planning"
        '.Add Factory.CreateAdminForm, "Admin Control Panel"
        '.Add Factory.DocumentConfigForm, "Document Config"
    End With
    Set SetUpForms = coll
End Function

Public Sub SaveSettings()
    settings_.LastSheetUsed = ActiveSheet.Name
    With Application
        settings_.ScreenUpdating = .ScreenUpdating
        settings_.DisplayAlerts = .DisplayAlerts
        settings_.DisplayStatusBar = .DisplayStatusBar
        settings_.Calculation = .Calculation
    End With
End Sub

Public Sub PerformanceMode(enable As Boolean, Optional DisableEvents As Boolean = False)
' This Sub turns off un-necessary gui functions to increase performance
    PerformanceModeEnabled = enable

    If enable Then
        Logger.Log "Performance Mode : ON", RuntimeLog
        ' Save current excel-gui settings
        SaveSettings
        ' Enabled performance mode
        With Application
            If .ScreenUpdating Then .ScreenUpdating = False
            If .DisplayAlerts Then .DisplayAlerts = False
            If .DisplayStatusBar Then .DisplayStatusBar = False
            'If .EnableEvents And DisableEvents Then .EnableEvents = False
            If .Calculation = xlCalculationAutomatic Then .Calculation = xlCalculationManual
        End With
    Else
        Logger.Log "Performance Mode : OFF", RuntimeLog
        With Application
            .ScreenUpdating = settings_.ScreenUpdating
            .DisplayAlerts = settings_.DisplayAlerts
            .DisplayStatusBar = settings_.DisplayStatusBar
            .Calculation = settings_.Calculation
        End With
    End If


End Sub

Sub HideExcelItems()

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveSheet.DisplayPageBreaks = False
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With

    'ThisWorkbook.Protect ""
End Sub

Sub UnHideExcelItems()

    'ThisWorkbook.Unprotect ""

    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True

    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
End Sub

' Forms Interface Methods
Public Property Get CurrentForm() As IForm
        Set CurrentForm = GUI.GetForm("Form" & Right(ActiveSheet.CodeName, Len(ActiveSheet.CodeName) - 3))
End Property

Public Sub PassFormCommand(cmd As String)
' Using the active sheet as the "form"
    ' Toggle Performance Mode
    If Not GUI.PerformanceModeEnabled Then GUI.PerformanceMode (True)
    On Error GoTo Catch
    If Not GUI.IsRunning Then
        Start
    End If
    On Error GoTo 0
    GoTo Finally
Catch:
    Start
Finally:
    Dim frm As IForm
    Set frm = GUI.CurrentForm
    frm.CmdByName cmd
    ' Turn off Performance Mode
    If GUI.PerformanceModeEnabled Then GUI.PerformanceMode (False)
End Sub

Function ClearForm(Name As String) As Boolean
    On Error GoTo catchError
    GUI.CurrentForm.Clear
exitFunction:
    Exit Function
catchError:
    Logger.Log "Form not found."
    GoTo exitFunction
End Function
