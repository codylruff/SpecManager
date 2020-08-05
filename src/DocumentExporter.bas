Attribute VB_Name = "DocumentExporter"
Option Explicit
'===================================
'DESCRIPTION: DocumentExporter
'===================================
Private Sub Search()
' Retreives the documents for the specified material id
    Dim material_id As String
    ' Start the application
    App.Start
    material_id = CStr(GUI.CurrentForm.FieldValue("material_id")) ' this is the material id (SAP Code)
    SpecManager.MaterialInput (material_id)
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
    material_id = CStr(GUI.CurrentForm.FieldValue("material_id")) ' this is the material id (SAP Code)
    machine_id = CStr(GUI.CurrentForm.FieldValue("machine_id"))   ' This is the machine id (ie. loom number, warper, etc...)
    document_id = CStr(GUI.CurrentForm.FieldValue("document_id")) ' This is the template the document is based on
    dir = GUI.CurrentForm.FieldValue("file_dir")
    
    On Error GoTo DocumentUIDNotFound
    Set doc = App.DocumentsByUID(document_id)
    On Error GoTo 0

    Select Case file_type
        Case ".PDF"
            App.printer.ToPDF Sheets(doc.SpecType), CStr(dir)
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
        Case Default
            Prompt.Error "No file type selected."
    End Select
    GoTo Finally
DocumentUIDNotFound:
    Prompt.Error "Could not retrieve requested document."
Finally:
    App.Shutdown
End Sub

Public Sub ExportAsPdf()
' Export selected document as a .pdf
    ExportDocument ".PDF"
End Sub

Public Sub ExportAsXlsx()
' Export selected document as a .xlsx
    ExportDocument ".XLSX"
End Sub

'Private Sub ExportDocuments()
'FIXME
'' Exports all document objects associated with a single material_id to an excel file.
'    Dim shts As Variant
'    Dim Count As Long
'    Count = App.specs.Count
'    For i = 0 To Count - 1
'        shts(i) =
'    Next i
'    Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Copy
'    With ActiveWorkbook
'        .SaveAs fileName:=Environ("TEMP") & "\New3.xlsx", FileFormat:=xlOpenXMLWorkbook
'        .Close SaveChanges:=False
'    End With
'
'End Sub
