VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormPortal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Description
'====================================
'Factory Module Constructor function:
'Function CreateFormPortal() As FormPortal
Implements IForm
' DEFAULT PROPERTY VALUES
Private Const FORM_NAME As String = "FormPortal"
Private Const HELP_TEXT As String = "TODO" ' Thinking of loading this from a text file??
Private m_IForm_Console As Console
Private m_IForm_Fields(3) As Variant
'--------------
' I_PROPERTIES
'--------------
Public Property Set IForm_Console(value As Console): Set m_IForm_Console = value: End Property
Public Property Get IForm_Console() As Console
        Set IForm_Console = m_IForm_Console
End Property

Public Property Get IForm_Sheet() As Worksheet
    Set IForm_Sheet = shtPortal
End Property

Public Property Get IForm_Name() As String
    IForm_Name = FORM_NAME
End Property

Public Property Get IForm_HelpText() As String
    IForm_HelpText = HELP_TEXT
End Property

Public Property Get IForm_Fields() As Variant
    IForm_Fields = m_IForm_Fields
End Property
'----------
' METHODS
'----------
Private Sub PopulateDocumentIdList()
    Dim ids As Variant
    ids = App.DocumentsByUID.keys
    Me.IForm_ChangeFieldValue "document_id", nullstr
    With Me.IForm_Sheet.Range("document_id")
        .Validation.Delete
        .value = "Select a Document Id . . ."
        .Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=IIf(Utils.ArrayLength(ids) > 1, Join(ids, ","), CStr(ids(0)))
    End With
End Sub

Private Sub MaterialSearch()
    SpecManager.MaterialInput UCase(IForm_Sheet.Range("material_id"))
    Logger.Log "Listing Documents . . . "
    If Not App.specs Is Nothing Then
        App.printer.ListObjects App.DocumentsByUID
    Else
        App.printer.WriteLine "No specifications are available for this code."
    End If
    If Me.IForm_Console.text = nullstr Then
        Me.IForm_Console.text = "No specifications are available for this code."
    End If
End Sub

Private Sub ExportDocument(file_type As String)
' Exports a document object to an excel file.
    Dim doc As Document
    Dim ret_val As Long
    Dim material_id As String
    Dim description As String
    Dim file_dir As String
    Dim machine_id As String
    Dim Revision As String
    Dim document_id As String
    Dim dir As String

    ' Initialize document parameters
    material_id = CStr(Me.IForm_FieldValue("material_id")) ' this is the material id (SAP Code)
    machine_id = CStr(Me.IForm_FieldValue("machine_id"))   ' This is the machine id (ie. loom number, warper, etc...)
    document_id = CStr(Me.IForm_FieldValue("document_id")) ' This is the template the document is based on
    dir = Me.IForm_FieldValue("file_dir")
    On Error GoTo Catch
    If document_id = nullstr Then
        Prompt.Error "Please select a document to export."
        err.Raise E_NO_DOCUMENT_SELECTED
    End If
    
    On Error GoTo DocumentUIDNotFound
    Set doc = App.DocumentsByUID(document_id)
    On Error GoTo Catch
    App.printer.PrintObjectToSheet doc, Utils.SheetByName(doc.SpecType)
    Select Case file_type
        Case ".PDF"
            App.printer.ToPDF Sheets(doc.SpecType), False, dir & "\" & doc.fileName & ".pdf"
        Case ".XLSX"
            Sheets(doc.SpecType).Visible = xlSheetVisible
            Sheets(doc.SpecType).Copy
            With ActiveWorkbook
                .SaveAs _
                fileName:=dir & "\" & doc.fileName & ".xlsx", _
                FileFormat:=xlOpenXMLWorkbook
                .Close SaveChanges:=False
            End With
            Sheets(doc.SpecType).Visible = xlSheetHidden
            Logger.Log "File Saved : " & dir & "\" & doc.fileName & ".xlsx"
        Case Default
            Prompt.Error "No file type selected."
    End Select
    GoTo Finally
DocumentUIDNotFound:
    Prompt.Error "Could not retrieve requested document."
Catch:
    ' Pass
Finally:
    err.Clear
End Sub
'----------
' I_METHODS
'----------
Public Function IForm_Clear()
    ' Console must be cleared through object
    IForm_Console.Clear
    ' Fields can be cleared manually.
    For i = 0 To UBound(IForm_Fields) - 1
        IForm_Sheet.Range(IForm_Fields(i)).value = nullstr
    Next i
End Function

Public Function IForm_CmdByName(cmd As String) As Boolean
    Dim ret_val As Boolean
    Select Case cmd
        Case "Search"
            cmdSearch
            ret_val = True
        Case "SelectFolder"
            cmdSelectFolder
            ret_val = True
        Case "ExportAsXlsx"
            cmdExportAsXlsx
            ret_val = True
        Case "ExportAsPdf"
            cmdExportAsPdf
            ret_val = True
        Case Default
            ret_val = False
    End Select
    IForm_CmdByName = ret_val
End Function

Public Function IForm_ReformatSheet() As Boolean
    ' TODO Implement this fully
    err.Raise E_NOTIMPL, "Function Not Implemented"
    IForm_ReformatSheet = False
End Function

Public Function IForm_FieldValue(field As String) As Variant
' Reads the value of a field in the form
    Dim i As Long
    Dim val As Variant
    On Error GoTo Catch
    IForm_FieldValue = IForm_Sheet.Range(field).value
    GoTo Finally
Catch:
    err.Raise INTERNAL_ERR, "Field not found"
    IForm_FieldValue = nullstr
Finally:
    ' Pass
End Function

Public Function IForm_ChangeFieldValue(field As String, val As Variant)
' Changes the value of a field in the form
    IForm_Sheet.Range(field).value = val
    'err.Raise INTERNAL_ERR, "Field not found"
End Function

'----------
' COMMANDS
'----------
Private Sub cmdSearch()
    ' Check for any white space and remove it
    If Utils.RemoveWhiteSpace(IForm_Sheet.Range("material_id")) = nullstr Then
       Prompt.Error "Please enter a material id."
       Exit Sub
    End If
    ' Run search routine
    MaterialSearch
    PopulateDocumentIdList
End Sub

Private Sub cmdSelectFolder()
' Select default directory
    Me.IForm_ChangeFieldValue "file_dir", Prompt.SelectFolder
End Sub

Public Sub cmdExportAsPdf()
' Export selected document as a .pdf
    ExportDocument ".PDF"
End Sub

Public Sub cmdExportAsXlsx()
' Export selected document as a .xlsx
    ExportDocument ".XLSX"
End Sub

'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    Set Me.IForm_Console = Factory.CreateConsole(Me.IForm_Sheet)
    ' Initialize the fields array to save on Array() calls.
    'm_IForm_Fields = Array("material_id", "work_order", "machine_id")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    'REVIEW Is there anything that needs to be destroyed here?
End Sub
