Attribute VB_Name = "DocumentParser"
Option Explicit
' ==============================================
' DOCUMENT PARSER
' ==============================================

Public Sub LoadNewDocument()
    Dim file_path As String
    Dim path_no_ext As String
    Dim path_len As Integer
    Dim char_count As Integer
    Dim char_buffer As String
    Dim material_number As String
    Dim json_string As String
    Dim progress_bar As Long
    Dim spec As Specification
    Dim ret_val As Long
    Dim template_type As String

    ' Start the application
    App.Start

    ' Prompt user to select file path
    file_path = PromptHandler.SelectSpecifcationFile
    
    ' Initialize an empty specification
    Set spec = CreateSpecification
    
    ' Prompt the user for a template name. Then Validate the template exists
    template_type = PromptHandler.EnterTemplateType
    On Error GoTo InvalidTemplateType
    Set spec.Template = App.templates(template_type)
    On Error GoTo 0

    ' Initialize the progress bar
    progress_bar = App.gDll.ShowProgressBar(4)

    ' Task 1 Extract material Id from file name.
    progress_bar = App.gDll.SetProgressBar(progress_bar, 1, "Task 1/4")
    path_no_ext = Replace(file_path, ".xlsx", vbNullString)
    path_len = Len(path_no_ext)
    char_count = path_len

    ' Throw error for improper file name
    On Error GoTo FileNamingError
    Do Until char_buffer = "_"
        char_count = char_count - 1
        char_buffer = Mid(path_no_ext, char_count, 1)
    Loop
    On Error GoTo 0
    
    ' Task 2 Parse the Document for a json string
    progress_bar = App.gDll.SetProgressBar(progress_bar, 2, "Task 2/4")
    material_number = Mid(path_no_ext, char_count + 1, path_len - char_count)
    json_string = JsonVBA.ConvertToJson(ParseDocument(file_path, template_type))

    ' Task 3 Convert json string into specification object
    progress_bar = App.gDll.SetProgressBar(progress_bar, 3, "Task 3/4")

    ' Create specification from json string
    spec.JsonToObject json_string
    spec.MaterialId = material_number
    spec.SpecType = template_type
    spec.Revision = "1.0"

    ' Task 4 Save specification object to the database
    progress_bar = App.gDll.SetProgressBar(progress_bar, 4, "Task 4/4", AutoClose:=True)

    ' Save Specification to database
    ret_val = SpecManager.SaveNewSpecification(spec)

    ' Parse return value.
    If ret_val = DB_PUSH_SUCCESS Then
        PromptHandler.Success "New Specification Saved."
    ElseIf ret_val = SM_MATERIAL_EXISTS Then
        PromptHandler.Error "Material Already Exists."
    Else
        PromptHandler.Error "Error Saving Specification."
    End If

    ' Stop the app
    App.Shutdown
    Exit Sub
    
FileNamingError:
    PromptHandler.Error "File named improperly."
    Exit Sub
InvalidTemplateType:
    PromptHandler.Error template_type & " Does Not Exist!"
End Sub

Public Function ParseDocument(path As String, template_type As String) As Object
    Dim wb As Workbook
    Dim doc_dict As Object
    Dim i As Long
    Dim arr
    
    ' Turn on Performance Mode
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Open work book and initialze dictionary
    Set wb = OpenWorkbook(path)
    Set doc_dict = CreateObject("Scripting.Dictionary")

    ' Retrieve names from the document
    ' The sheet must be named after the Spec_Type to work.
    arr = Utils.GetNames(wb, template_type)

    On Error GoTo ParsingError
    For i = 0 To UBound(arr, 1) - 1
        doc_dict.Add arr(i, 0), arr(i, 1)
    Next i
    On Error GoTo 0
    
    Set ParseDocument = doc_dict
    wb.Close

    ' Turn off Performance Mode
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Function

ParsingError:
    PromptHandler.Error "Parsing Error @" & CStr(arr(i, 0))
End Function

Public Sub CreateTemplateFromFile()
    Dim file_path As String
    Dim Template As SpecificationTemplate

    ' Start the application
    App.Start

    ' Prompt user to select file path
    file_path = PromptHandler.SelectSpecifcationFile

    ' Spec_Type will be name of file selected
    Set Template = Factory.CreateTemplateFromJsonFile(file_path)

    ' Parse database return value.
    If SpecManager.SaveSpecificationTemplate(Template) <> DB_PUSH_SUCCESS Then
        PromptHandler.Error "Failed to create " & Template.SpecType
        Exit Sub
    End If
    PromptHandler.Success Template.SpecType & " Created Successfully!"
End Sub

'Public Function ParsePsf(path As String) As Object
'    Dim wb As Workbook
'    Dim strFile As String
'    Dim psf_dict As Object
'    Dim prop As Variant
'    Dim nr As Name
'    Dim rng As Object
'
'    ' Turn on Performance Mode
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False

'    Debug.Print "Parsing PSF @ " & path
'    Dim ret_val As Long
'    Set wb = OpenWorkbook(path)
'    DeleteNames wb
'    Set psf_dict = CreateObject("Scripting.Dictionary")
'    Set psf_dict = CreatePsfNames(wb)
'    Set ParsePsf = psf_dict
'    Set psf_dict = Nothing
'    wb.Close
'
'    ' Turn off Performance Mode
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'
'End Function

' Public Function CreatePsfNames(wb As Workbook) As Object
'     Dim ws As Worksheet
'     Set ws = wb.Sheets("Spec Sheet")
'     Dim dict As Object
'     Set dict = CreateObject("Scripting.Dictionary")
'     With wb.Names
'         .Add Name:="bsf_max", RefersTo:=ws.Range("F24") ' Breaking Strength Weft Max
'         .Add Name:="bsf_min", RefersTo:=ws.Range("D24") ' Breaking Strength Weft Min
'         .Add Name:="bsf_nom", RefersTo:=ws.Range("E24") ' Breaking Strength Weft Nominal
'         .Add Name:="bsw_max", RefersTo:=ws.Range("F23") ' Breaking Strength Warp Max
'         .Add Name:="bsw_min", RefersTo:=ws.Range("D23") ' Breaking Strength Warp Minimum
'         .Add Name:="bsw_nom", RefersTo:=ws.Range("E23") ' Breaking Strength Warp Nominal
'         .Add Name:="cad_max", RefersTo:=ws.Range("F15") ' Conditioned Weight Max
'         .Add Name:="cad_min", RefersTo:=ws.Range("D15") ' Conditioned Weight Minimum
'         .Add Name:="cad_nom", RefersTo:=ws.Range("E15") ' Conditioned Weight Nominal
'         .Add Name:="coated_ad_max", RefersTo:=ws.Range("F39") ' Coated Fabric Weight Max
'         .Add Name:="coated_ad_min", RefersTo:=ws.Range("D39") ' Coated Fabric Weight Minimum
'         .Add Name:="coated_ad_nom", RefersTo:=ws.Range("E39") ' Coated Fabric Weight Nominal
'         .Add Name:="coating_description", RefersTo:=ws.Range("D38") ' Coating Description
'         .Add Name:="code", RefersTo:=ws.Range("B6") ' SAP Code
'         .Add Name:="core_id", RefersTo:=ws.Range("D33") ' Core Inside Diameter
'         .Add Name:="core_len", RefersTo:=ws.Range("D34") ' Core Length
'         .Add Name:="core_material", RefersTo:=ws.Range("G34") ' Core Length
'         .Add Name:="da_max", RefersTo:=ws.Range("F27") ' Dynamic Water Absorption Max
'         .Add Name:="da_min", RefersTo:=ws.Range("D27") ' Dynamic Water Absorption Minimum
'         .Add Name:="da_nom", RefersTo:=ws.Range("E27") ' Dynamic Water Absorption Nominal
'         .Add Name:="dad_max", RefersTo:=ws.Range("F16") ' Dry Weight Max
'         .Add Name:="dad_min", RefersTo:=ws.Range("D16") ' Dry Weight Minimum
'         .Add Name:="dad_nom", RefersTo:=ws.Range("E16") ' Dry Weight Nominal
'         .Add Name:="description", RefersTo:=ws.Range("B5") ' Description
'         .Add Name:="fiber", RefersTo:=ws.Range("D9") ' Fiber Description
'         .Add Name:="fil_nom", RefersTo:=ws.Range("E14") ' Weft Thread Count Max
'         .Add Name:="fill_max", RefersTo:=ws.Range("F14") ' Weft Thread Count Minimum
'         .Add Name:="fill_min", RefersTo:=ws.Range("D14") ' Weft Thread Count Nominal
'         .Add Name:="finish", RefersTo:=ws.Range("D25") ' Finishing Treatment
'         .Add Name:="flag_defects", RefersTo:=ws.Range("D31") ' Flag Defects
'         .Add Name:="fringe_max", RefersTo:=ws.Range("F21") ' Fringe Length Max
'         .Add Name:="fringe_min", RefersTo:=ws.Range("D21") ' Fringe Length Minimum
'         .Add Name:="fringe_nom", RefersTo:=ws.Range("E21") ' Fringe Length Nominal
'         .Add Name:="insp_docs", RefersTo:=ws.Range("D32") ' Inspection Documents
'         .Add Name:="lam_or_coat", RefersTo:=ws.Range("D37") ' Laminated or Coated Fabric
'         .Add Name:="len_max", RefersTo:=ws.Range("F19") ' Roll Length Max
'         .Add Name:="len_min", RefersTo:=ws.Range("D19") ' Roll Length Minimum
'         .Add Name:="len_nom", RefersTo:=ws.Range("E19") ' Roll Length Nominal
'         .Add Name:="min_acceptable_len", RefersTo:=ws.Range("D20") ' Min. Accept. Roll Length
'         .Add Name:="packaging_reqs", RefersTo:=ws.Range("C36") ' Packaging
'         .Add Name:="pallet_size", RefersTo:=ws.Range("D35") ' Pallet Size
'         .Add Name:="product", RefersTo:=ws.Range("B4") ' Product
'         .Add Name:="psf_no", RefersTo:=ws.Range("B3") ' PSF No
'         .Add Name:="rc_max", RefersTo:=ws.Range("F40") ' Resin Content Max
'         .Add Name:="rc_min", RefersTo:=ws.Range("D40") ' Resin Content Minimum
'         .Add Name:="rc_nom", RefersTo:=ws.Range("E40") ' Resin Content Nominal
'         .Add Name:="rev_no", RefersTo:=ws.Range("F3") ' Revision
'         .Add Name:="shrink", RefersTo:=ws.Range("D26") ' Residual Shrinkage
'         .Add Name:="sox_max", RefersTo:=ws.Range("F28") ' Extractable Level Max
'         .Add Name:="sox_min", RefersTo:=ws.Range("D28") ' Extractable Level Minimum
'         .Add Name:="sox_nom", RefersTo:=ws.Range("E28") ' Extractable Level Nominal
'         .Add Name:="storage_reqs", RefersTo:=ws.Range("B43") ' Storage Conditions
'         .Add Name:="tape_selvage", RefersTo:=ws.Range("D18") ' Tape Selvage
'         .Add Name:="tg_max", RefersTo:=ws.Range("F42") ' Glass Transition Temperature Max
'         .Add Name:="tg_min", RefersTo:=ws.Range("D42") ' Glass Transition Temperature Minimum
'         .Add Name:="tg_nom", RefersTo:=ws.Range("E42") ' Glass Transition Temperature Nominal
'         .Add Name:="thick_max", RefersTo:=ws.Range("F22") ' Final Fabric Thickness Max
'         .Add Name:="thick_min", RefersTo:=ws.Range("D22") ' Final Fabric Thickness Minimum
'         .Add Name:="thick_nom", RefersTo:=ws.Range("E22") ' Final Fabric Thickness Nominal
'         .Add Name:="v50_layers", RefersTo:=ws.Range("D30") ' Number of Layer
'         .Add Name:="v50_min", RefersTo:=ws.Range("D29") ' Minimum Average V50
'         .Add Name:="v50_test_reqs", RefersTo:=ws.Range("E29") ' Minimum Average V50
'         .Add Name:="vc_max", RefersTo:=ws.Range("F41") ' Volatile Content Max
'         .Add Name:="vc_min", RefersTo:=ws.Range("D41") ' Volatile Content Minimum
'         .Add Name:="vc_nom", RefersTo:=ws.Range("E41") ' Volatile Content Nominal
'         .Add Name:="warp_max", RefersTo:=ws.Range("F13") ' Warp Thread Count Max
'         .Add Name:="warp_min", RefersTo:=ws.Range("D13") ' Warp Thread Count Minimum
'         .Add Name:="warp_nom", RefersTo:=ws.Range("E13") ' Warp Thread Count Nominal
'         .Add Name:="warp_yarn", RefersTo:=ws.Range("D10") ' Warp Yarn
'         .Add Name:="weave", RefersTo:=ws.Range("D12") ' Weave Style
'         .Add Name:="weft_yarn", RefersTo:=ws.Range("D11") ' Weft Yarn
'         .Add Name:="wid_max", RefersTo:=ws.Range("F17") ' Final Product Useful Width Max
'         .Add Name:="wid_min", RefersTo:=ws.Range("D17") ' Final Product Useful Width Minimum
'         .Add Name:="wid_nom", RefersTo:=ws.Range("E17") ' Final Product Useful Width Nominal
'     End With
'     With dict
'         .Add Key:="bsf_max", item:=IIf(Range(wb.Names("bsf_max")).value = vbNullString, vbNullString, Range(wb.Names("bsf_max")).value)
'         .Add Key:="bsf_min", item:=IIf(Range(wb.Names("bsf_min")).value = vbNullString, vbNullString, Range(wb.Names("bsf_min")).value)
'         .Add Key:="bsf_nom", item:=IIf(Range(wb.Names("bsf_nom")).value = vbNullString, vbNullString, Range(wb.Names("bsf_nom")).value)
'         .Add Key:="bsw_max", item:=IIf(Range(wb.Names("bsw_max")).value = vbNullString, vbNullString, Range(wb.Names("bsw_max")).value)
'         .Add Key:="bsw_min", item:=IIf(Range(wb.Names("bsw_min")).value = vbNullString, vbNullString, Range(wb.Names("bsw_min")).value)
'         .Add Key:="bsw_nom", item:=IIf(Range(wb.Names("bsw_nom")).value = vbNullString, vbNullString, Range(wb.Names("bsw_nom")).value)
'         .Add Key:="cad_max", item:=IIf(Range(wb.Names("cad_max")).value = vbNullString, vbNullString, Range(wb.Names("cad_max")).value)
'         .Add Key:="cad_min", item:=IIf(Range(wb.Names("cad_min")).value = vbNullString, vbNullString, Range(wb.Names("cad_min")).value)
'         .Add Key:="cad_nom", item:=IIf(Range(wb.Names("cad_nom")).value = vbNullString, vbNullString, Range(wb.Names("cad_nom")).value)
'         .Add Key:="coated_ad_max", item:=IIf(Range(wb.Names("coated_ad_max")).value = vbNullString, vbNullString, Range(wb.Names("coated_ad_max")).value)
'         .Add Key:="coated_ad_min", item:=IIf(Range(wb.Names("coated_ad_min")).value = vbNullString, vbNullString, Range(wb.Names("coated_ad_min")).value)
'         .Add Key:="coated_ad_nom", item:=IIf(Range(wb.Names("coated_ad_nom")).value = vbNullString, vbNullString, Range(wb.Names("coated_ad_nom")).value)
'         .Add Key:="coating_description", item:=IIf(Range(wb.Names("coating_description")).value = vbNullString, vbNullString, Range(wb.Names("coating_description")).value)
'         .Add Key:="code", item:=IIf(Range(wb.Names("code")).value = vbNullString, vbNullString, Range(wb.Names("code")).value)
'         .Add Key:="core_id", item:=IIf(Range(wb.Names("core_id")).value = vbNullString, vbNullString, Range(wb.Names("core_id")).value)
'         .Add Key:="core_len", item:=IIf(Range(wb.Names("core_len")).value = vbNullString, vbNullString, Range(wb.Names("core_len")).value)
'         .Add Key:="core_material", item:=IIf(Range(wb.Names("core_material")).value = vbNullString, vbNullString, Range(wb.Names("core_material")).value)
'         .Add Key:="da_max", item:=IIf(Range(wb.Names("da_max")).value = vbNullString, vbNullString, Range(wb.Names("da_max")).value)
'         .Add Key:="da_min", item:=IIf(Range(wb.Names("da_min")).value = vbNullString, vbNullString, Range(wb.Names("da_min")).value)
'         .Add Key:="da_nom", item:=IIf(Range(wb.Names("da_nom")).value = vbNullString, vbNullString, Range(wb.Names("da_nom")).value)
'         .Add Key:="dad_max", item:=IIf(Range(wb.Names("dad_max")).value = vbNullString, vbNullString, Range(wb.Names("dad_max")).value)
'         .Add Key:="dad_min", item:=IIf(Range(wb.Names("dad_min")).value = vbNullString, vbNullString, Range(wb.Names("dad_min")).value)
'         .Add Key:="dad_nom", item:=IIf(Range(wb.Names("dad_nom")).value = vbNullString, vbNullString, Range(wb.Names("dad_nom")).value)
'         .Add Key:="description", item:=IIf(Range(wb.Names("description")).value = vbNullString, vbNullString, Range(wb.Names("description")).value)
'         .Add Key:="fiber", item:=IIf(Range(wb.Names("fiber")).value = vbNullString, vbNullString, Range(wb.Names("fiber")).value)
'         .Add Key:="fil_nom", item:=IIf(Range(wb.Names("fil_nom")).value = vbNullString, vbNullString, Range(wb.Names("fil_nom")).value)
'         .Add Key:="fill_max", item:=IIf(Range(wb.Names("fill_max")).value = vbNullString, vbNullString, Range(wb.Names("fill_max")).value)
'         .Add Key:="fill_min", item:=IIf(Range(wb.Names("fill_min")).value = vbNullString, vbNullString, Range(wb.Names("fill_min")).value)
'         .Add Key:="finish", item:=IIf(Range(wb.Names("finish")).value = vbNullString, vbNullString, Range(wb.Names("finish")).value)
'         .Add Key:="flag_defects", item:=IIf(Range(wb.Names("flag_defects")).value = vbNullString, vbNullString, Range(wb.Names("flag_defects")).value)
'         .Add Key:="fringe_max", item:=IIf(Range(wb.Names("fringe_max")).value = vbNullString, vbNullString, Range(wb.Names("fringe_max")).value)
'         .Add Key:="fringe_min", item:=IIf(Range(wb.Names("fringe_min")).value = vbNullString, vbNullString, Range(wb.Names("fringe_min")).value)
'         .Add Key:="fringe_nom", item:=IIf(Range(wb.Names("fringe_nom")).value = vbNullString, vbNullString, Range(wb.Names("fringe_nom")).value)
'         .Add Key:="insp_docs", item:=IIf(Range(wb.Names("insp_docs")).value = vbNullString, vbNullString, Range(wb.Names("insp_docs")).value)
'         .Add Key:="lam_or_coat", item:=IIf(Range(wb.Names("lam_or_coat")).value = vbNullString, vbNullString, Range(wb.Names("lam_or_coat")).value)
'         .Add Key:="len_max", item:=IIf(Range(wb.Names("len_max")).value = vbNullString, vbNullString, Range(wb.Names("len_max")).value)
'         .Add Key:="len_min", item:=IIf(Range(wb.Names("len_min")).value = vbNullString, vbNullString, Range(wb.Names("len_min")).value)
'         .Add Key:="len_nom", item:=IIf(Range(wb.Names("len_nom")).value = vbNullString, vbNullString, Range(wb.Names("len_nom")).value)
'         .Add Key:="min_acceptable_len", item:=IIf(Range(wb.Names("min_acceptable_len")).value = vbNullString, vbNullString, Range(wb.Names("min_acceptable_len")).value)
'         .Add Key:="packaging_reqs", item:=IIf(Range(wb.Names("packaging_reqs")).value = vbNullString, vbNullString, Range(wb.Names("packaging_reqs")).value)
'         .Add Key:="pallet_size", item:=IIf(Range(wb.Names("pallet_size")).value = vbNullString, vbNullString, Range(wb.Names("pallet_size")).value)
'         .Add Key:="product", item:=IIf(Range(wb.Names("product")).value = vbNullString, vbNullString, Range(wb.Names("product")).value)
'         .Add Key:="psf_no", item:=IIf(Range(wb.Names("psf_no")).value = vbNullString, vbNullString, Range(wb.Names("psf_no")).value)
'         .Add Key:="rc_max", item:=IIf(Range(wb.Names("rc_max")).value = vbNullString, vbNullString, Range(wb.Names("rc_max")).value)
'         .Add Key:="rc_min", item:=IIf(Range(wb.Names("rc_min")).value = vbNullString, vbNullString, Range(wb.Names("rc_min")).value)
'         .Add Key:="rc_nom", item:=IIf(Range(wb.Names("rc_nom")).value = vbNullString, vbNullString, Range(wb.Names("rc_nom")).value)
'         .Add Key:="rev_no", item:=IIf(Range(wb.Names("rev_no")).value = vbNullString, vbNullString, Range(wb.Names("rev_no")).value)
'         .Add Key:="shrink", item:=IIf(Range(wb.Names("shrink")).value = vbNullString, vbNullString, Range(wb.Names("shrink")).value)
'         .Add Key:="sox_max", item:=IIf(Range(wb.Names("sox_max")).value = vbNullString, vbNullString, Range(wb.Names("sox_max")).value)
'         .Add Key:="sox_min", item:=IIf(Range(wb.Names("sox_min")).value = vbNullString, vbNullString, Range(wb.Names("sox_min")).value)
'         .Add Key:="sox_nom", item:=IIf(Range(wb.Names("sox_nom")).value = vbNullString, vbNullString, Range(wb.Names("sox_nom")).value)
'         .Add Key:="storage_reqs", item:=IIf(Range(wb.Names("storage_reqs")).value = vbNullString, vbNullString, Range(wb.Names("storage_reqs")).value)
'         .Add Key:="tape_selvage", item:=IIf(Range(wb.Names("tape_selvage")).value = vbNullString, vbNullString, Range(wb.Names("tape_selvage")).value)
'         .Add Key:="tg_max", item:=IIf(Range(wb.Names("tg_max")).value = vbNullString, vbNullString, Range(wb.Names("tg_max")).value)
'         .Add Key:="tg_min", item:=IIf(Range(wb.Names("tg_min")).value = vbNullString, vbNullString, Range(wb.Names("tg_min")).value)
'         .Add Key:="tg_nom", item:=IIf(Range(wb.Names("tg_nom")).value = vbNullString, vbNullString, Range(wb.Names("tg_nom")).value)
'         .Add Key:="thick_max", item:=IIf(Range(wb.Names("thick_max")).value = vbNullString, vbNullString, Range(wb.Names("thick_max")).value)
'         .Add Key:="thick_min", item:=IIf(Range(wb.Names("thick_min")).value = vbNullString, vbNullString, Range(wb.Names("thick_min")).value)
'         .Add Key:="thick_nom", item:=IIf(Range(wb.Names("thick_nom")).value = vbNullString, vbNullString, Range(wb.Names("thick_nom")).value)
'         .Add Key:="v50_layers", item:=IIf(Range(wb.Names("v50_layers")).value = vbNullString, vbNullString, Range(wb.Names("v50_layers")).value)
'         .Add Key:="v50_min", item:=IIf(Range(wb.Names("v50_min")).value = vbNullString, vbNullString, Range(wb.Names("v50_min")).value)
'         .Add Key:="v50_test_reqs", item:=IIf(Range(wb.Names("v50_test_reqs")).value = vbNullString, vbNullString, Range(wb.Names("v50_test_reqs")).value)
'         .Add Key:="vc_max", item:=IIf(Range(wb.Names("vc_max")).value = vbNullString, vbNullString, Range(wb.Names("vc_max")).value)
'         .Add Key:="vc_min", item:=IIf(Range(wb.Names("vc_min")).value = vbNullString, vbNullString, Range(wb.Names("vc_min")).value)
'         .Add Key:="vc_nom", item:=IIf(Range(wb.Names("vc_nom")).value = vbNullString, vbNullString, Range(wb.Names("vc_nom")).value)
'         .Add Key:="warp_max", item:=IIf(Range(wb.Names("warp_max")).value = vbNullString, vbNullString, Range(wb.Names("warp_max")).value)
'         .Add Key:="warp_min", item:=IIf(Range(wb.Names("warp_min")).value = vbNullString, vbNullString, Range(wb.Names("warp_min")).value)
'         .Add Key:="warp_nom", item:=IIf(Range(wb.Names("warp_nom")).value = vbNullString, vbNullString, Range(wb.Names("warp_nom")).value)
'         .Add Key:="warp_yarn", item:=IIf(Range(wb.Names("warp_yarn")).value = vbNullString, vbNullString, Range(wb.Names("warp_yarn")).value)
'         .Add Key:="weave", item:=IIf(Range(wb.Names("weave")).value = vbNullString, vbNullString, Range(wb.Names("weave")).value)
'         .Add Key:="weft_yarn", item:=IIf(Range(wb.Names("weft_yarn")).value = vbNullString, vbNullString, Range(wb.Names("weft_yarn")).value)
'         .Add Key:="wid_max", item:=IIf(Range(wb.Names("wid_max")).value = vbNullString, vbNullString, Range(wb.Names("wid_max")).value)
'         .Add Key:="wid_min", item:=IIf(Range(wb.Names("wid_min")).value = vbNullString, vbNullString, Range(wb.Names("wid_min")).value)
'         .Add Key:="wid_nom", item:=IIf(Range(wb.Names("wid_nom")).value = vbNullString, vbNullString, Range(wb.Names("wid_nom")).value)
'     End With

'     Set CreatePsfNames = dict

' End Function

Public Sub RenameRBAs()

    Dim material_number As String
    Dim r As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Set ws = shtDocumentParser
    
    For r = 1 To 19
        material_number = ws.Cells(r, 1)
        Set wb = OpenWorkbook(ThisWorkbook.path & "\RBAs\" & ws.Cells(r, 3) & ".xlsx")
        wb.SaveAs ThisWorkbook.path & "\RBAs\" & material_number & ".xlsx"
        wb.Close
    Next r
End Sub

Public Function ApplyNames(arr As Variant, wb As Workbook, ws_name As String) As Object
    Dim i As Long
    Dim ws As Worksheet
    Dim Name As String
    Dim addr As String
    Dim dict As Object
    Set ws = wb.Sheets(ws_name)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        addr = Chr(34) & arr(i, 1) & Chr(34)
        Name = Chr(34) & arr(i, 0) & Chr(34)
        wb.Names.Add Name:=Name, RefersTo:=ws.Range(addr)
        dict.Add Key:=Name, item:=IIf(Range(wb.Names(Name)).value = vbNullString, _
                                      vbNullString, Range(wb.Names(Name)).value)
    Next i
    Set ApplyNames = dict
End Function
