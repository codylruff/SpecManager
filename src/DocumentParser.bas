Attribute VB_Name = "DocumentParser"
Option Explicit
' ==============================================
' DOCUMENT PARSER
' ==============================================
'Public Sub ParseAll(Optional material_keyword As String = "code")
'    Dim material_number As String
'    Dim json_string As String
'    Dim k As Long
'    Dim json_object As Object
'    Dim json_file_path As String
'    Dim prop As Variant
'    Dim r As Long
'    Dim wb As Workbook
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("FormParser")
'    With ws
'        .Visible = True
'        Dim arrFiles As Variant
'        arrFiles = Utils.GetFiles(pfilters:=Array("xls", "xlsx"))
'        For r = 1 To UBound(arrFiles) - 1
'            Set json_object = ParsePsf(CStr(arrFiles(r)))
'            ' Clean data by removing units and filling in missing values
'            k = 2
'            .Cells(r, 1).value = json_object(material_keyword)
'            For Each prop In json_object
'                If json_object.item(prop) = Chr(34) Then json_object.item(prop) = Chr(34) + Chr(34)
'                If r = 1 Then
'                    .Cells(1, k).value = CStr(prop)
'                    .Cells(r + 1, k).value = json_object(prop)
'                Else
'                    .Cells(r, k).value = json_object(prop)
'                End If
'                 k = k + 1
'            Next prop
'            json_string = JsonVBA.ConvertToJson(json_object)
'            .Cells(r, k + 1).value = json_string
'        Next r
'    End With
'End Sub

Public Sub LoadNewRBA()
    Dim json_object As Object
    Dim file_path As String
    Dim path_no_ext As String
    Dim path_len As Integer
    Dim char_count As Integer
    Dim char_buffer As String
    Dim material_number As String
    Dim json_string As String
    Dim progress_bar As Long
    App.Start
    file_path = SelectSpecifcationFile
    progress_bar = App.gDll.ShowProgressBar(4)
    ' Task 1
    progress_bar = App.gDll.SetProgressBar(progress_bar, 1, "Task 1/4")
    path_no_ext = Replace(file_path, ".xlsx", vbNullString)
    path_len = Len(path_no_ext)
    char_count = path_len
    On Error GoTo FileNamingError
    Do Until char_buffer = "_"
        char_count = char_count - 1
        char_buffer = Mid(path_no_ext, char_count, 1)
    Loop
    On Error GoTo 0
    material_number = Mid(path_no_ext, char_count + 1, path_len - char_count)
    ' Task 2
    progress_bar = App.gDll.SetProgressBar(progress_bar, 2, "Task 2/4")
    Set json_object = ParseRBA(file_path)
    json_string = JsonVBA.ConvertToJson(json_object)
    ' Task 3
    progress_bar = App.gDll.SetProgressBar(progress_bar, 3, "Task 3/4")
    Dim spec As Specification
    Dim ret_val As Long
    Set spec = CreateSpecification
    Set spec.Template = App.templates("Weaving RBA")
    spec.JsonToObject json_string
    spec.MaterialId = material_number
    spec.SpecType = "Weaving RBA"
    spec.Revision = "1.0"
    ' Task 4
    progress_bar = App.gDll.SetProgressBar(progress_bar, 4, "Task 4/4", AutoClose:=True)
    ret_val = SpecManager.SaveNewSpecification(spec)
    Debug.Print spec.PropertiesJson
    If ret_val = DB_PUSH_SUCCESS Then
        PromptHandler.Success "New Specification Saved."
    ElseIf ret_val = SM_MATERIAL_EXISTS Then
        PromptHandler.Error "Material Already Exists."
    Else
        PromptHandler.Error "Error Saving Specification."
    End If
    App.Shutdown
    Exit Sub
FileNamingError:
    PromptHandler.Error "File named improperly."
End Sub

Public Function SelectSpecifcationFile() As String
' Select an RBA file from the file dialog.
    SelectSpecifcationFile = App.gDll.OpenFile("Select Specification Document . . .")
End Function

Public Function ParseRBA(path As String) As Object
    Dim wb As Workbook
    Dim strFile As String
    Dim rba_dict As Object
    Dim prop As Variant
    Dim nr As Name
    Dim rng As Object
    Dim i As Long
    
    ' Turn on Performance Mode
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ret_val As Long
    Set wb = OpenWorkbook(path)
    'DeleteNames wb
     Set rba_dict = CreateObject("Scripting.Dictionary")
    ' Clean data by removing units and filling in missing values
    Dim arr
    Dim rngs() As Variant
    arr = Utils.GetNames(wb, "Weaving RBA")
    ReDim rngs(UBound(arr, 1), 1)
    On Error Resume Next
    For i = 0 To UBound(arr, 1) - 1
        rba_dict.Add arr(i, 0), arr(i, 1)
    Next i
    On Error GoTo 0
    ret_val = JsonVBA.WriteJsonObject(path & ".json", rba_dict)
    Set ParseRBA = rba_dict
    'Set rba_dict = Nothing
    wb.Close
    ' Turn off Performance Mode
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function

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

' Public Function AddMoreRbaNames(dict As Object, wb As Workbook) As Object
    
'     With dict
'         .Add Key:="actual_weft_count", item:=IIf(Range(wb.Names("actual_weft_count")).value = vbNullString, vbNullString, Range(wb.Names("actual_weft_count")).value)
'         .Add Key:="article_code", item:=IIf(Range(wb.Names("article_code")).value = vbNullString, vbNullString, Range(wb.Names("article_code")).value)
'         .Add Key:="aux_selvedges_closing_degrees", item:=IIf(Range(wb.Names("aux_selvedges_closing_degrees")).value = vbNullString, vbNullString, Range(wb.Names("aux_selvedges_closing_degrees")).value)
'         .Add Key:="bottom_rapier_clamps", item:=IIf(Range(wb.Names("bottom_rapier_clamps")).value = vbNullString, vbNullString, Range(wb.Names("bottom_rapier_clamps")).value)
'         .Add Key:="bottom_spreader_bars", item:=IIf(Range(wb.Names("bottom_spreader_bars")).value = vbNullString, vbNullString, Range(wb.Names("bottom_spreader_bars")).value)
'         .Add Key:="central_selvedges_drawing_in", item:=IIf(Range(wb.Names("central_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_drawing_in")).value)
'         .Add Key:="central_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("central_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_ends_per_dent")).value)
'         .Add Key:="central_selvedges_number_ends", item:=IIf(Range(wb.Names("central_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_number_ends")).value)
'         .Add Key:="central_selvedges_weave", item:=IIf(Range(wb.Names("central_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_weave")).value)
'         .Add Key:="central_selvedges_yarn_count", item:=IIf(Range(wb.Names("central_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_yarn_count")).value)
'         .Add Key:="cutting_degrees", item:=IIf(Range(wb.Names("cutting_degrees")).value = vbNullString, vbNullString, Range(wb.Names("cutting_degrees")).value)
'         .Add Key:="date", item:=IIf(Range(wb.Names("date")).value = vbNullString, vbNullString, Range(wb.Names("date")).value)
'         .Add Key:="dorn_left_selvedges_drawing_in", item:=IIf(Range(wb.Names("dorn_left_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_drawing_in")).value)
'         .Add Key:="dorn_left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value)
'         .Add Key:="dorn_left_selvedges_number_ends", item:=IIf(Range(wb.Names("dorn_left_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_number_ends")).value)
'         .Add Key:="dorn_left_selvedges_weave", item:=IIf(Range(wb.Names("dorn_left_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_weave")).value)
'         .Add Key:="dorn_left_selvedges_yarn_count", item:=IIf(Range(wb.Names("dorn_left_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_yarn_count")).value)
'         .Add Key:="draw_in_harness", item:=IIf(Range(wb.Names("draw_in_harness")).value = vbNullString, vbNullString, Range(wb.Names("draw_in_harness")).value)
'         .Add Key:="draw_in_reed", item:=IIf(Range(wb.Names("draw_in_reed")).value = vbNullString, vbNullString, Range(wb.Names("draw_in_reed")).value)
'         .Add Key:="fabric_width", item:=IIf(Range(wb.Names("fabric_width")).value = vbNullString, vbNullString, Range(wb.Names("fabric_width")).value)
'         .Add Key:="first_heddle", item:=IIf(Range(wb.Names("first_heddle")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle")).value)
'         .Add Key:="first_heddle_1", item:=IIf(Range(wb.Names("first_heddle_1")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle_1")).value)
'         .Add Key:="first_heddle_guide", item:=IIf(Range(wb.Names("first_heddle_guide")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle_guide")).value)
'         .Add Key:="harness_configuration", item:=IIf(Range(wb.Names("harness_configuration")).value = vbNullString, vbNullString, Range(wb.Names("harness_configuration")).value)
'         .Add Key:="horizontal_back_rest_roller", item:=IIf(Range(wb.Names("horizontal_back_rest_roller")).value = vbNullString, vbNullString, Range(wb.Names("horizontal_back_rest_roller")).value)
'         .Add Key:="last_heddle", item:=IIf(Range(wb.Names("last_heddle")).value = vbNullString, vbNullString, Range(wb.Names("last_heddle")).value)
'         .Add Key:="last_heddle_guide", item:=IIf(Range(wb.Names("last_heddle_guide")).value = vbNullString, vbNullString, Range(wb.Names("last_heddle_guide")).value)
'         .Add Key:="left_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_main_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_drawing_in")).value)
'         .Add Key:="left_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_main_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_ends_per_dent")).value)
'         .Add Key:="left_main_selvedges_number_ends", item:=IIf(Range(wb.Names("left_main_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_number_ends")).value)
'         .Add Key:="left_main_selvedges_weave", item:=IIf(Range(wb.Names("left_main_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_weave")).value)
'         .Add Key:="left_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_main_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_yarn_count")).value)
'         .Add Key:="left_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_drawing_in")).value)
'         .Add Key:="left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_ends_per_dent")).value)
'         .Add Key:="left_selvedges_number_ends", item:=IIf(Range(wb.Names("left_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_number_ends")).value)
'         .Add Key:="left_selvedges_weave", item:=IIf(Range(wb.Names("left_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_weave")).value)
'         .Add Key:="left_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_yarn_count")).value)
'         .Add Key:="loom_number", item:=IIf(Range(wb.Names("loom_number")).value = vbNullString, vbNullString, Range(wb.Names("loom_number")).value)
'         .Add Key:="loom_type", item:=IIf(Range(wb.Names("loom_type")).value = vbNullString, vbNullString, Range(wb.Names("loom_type")).value)
'         .Add Key:="number_ends_wo_selvedges", item:=IIf(Range(wb.Names("number_ends_wo_selvedges")).value = vbNullString, vbNullString, Range(wb.Names("number_ends_wo_selvedges")).value)
'         .Add Key:="number_harnesses", item:=IIf(Range(wb.Names("number_harnesses")).value = vbNullString, vbNullString, Range(wb.Names("number_harnesses")).value)
'         .Add Key:="pinch_roller_felt_type", item:=IIf(Range(wb.Names("pinch_roller_felt_type")).value = vbNullString, vbNullString, Range(wb.Names("pinch_roller_felt_type")).value)
'         .Add Key:="press_roller_type", item:=IIf(Range(wb.Names("press_roller_type")).value = vbNullString, vbNullString, Range(wb.Names("press_roller_type")).value)
'         .Add Key:="rba_number", item:=IIf(Range(wb.Names("rba_number")).value = vbNullString, vbNullString, Range(wb.Names("rba_number")).value)
'         .Add Key:="reed", item:=IIf(Range(wb.Names("reed")).value = vbNullString, vbNullString, Range(wb.Names("reed")).value)
'         .Add Key:="reed_width", item:=IIf(Range(wb.Names("reed_width")).value = vbNullString, vbNullString, Range(wb.Names("reed_width")).value)
'         .Add Key:="right_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_main_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_drawing_in")).value)
'         .Add Key:="right_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_main_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_ends_per_dent")).value)
'         .Add Key:="right_main_selvedges_number_ends", item:=IIf(Range(wb.Names("right_main_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_number_ends")).value)
'         .Add Key:="right_main_selvedges_weave", item:=IIf(Range(wb.Names("right_main_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_weave")).value)
'         .Add Key:="right_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_main_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_yarn_count")).value)
'         .Add Key:="right_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_drawing_in")).value)
'         .Add Key:="right_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_ends_per_dent")).value)
'         .Add Key:="right_selvedges_number_ends", item:=IIf(Range(wb.Names("right_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_number_ends")).value)
'         .Add Key:="right_selvedges_weave", item:=IIf(Range(wb.Names("right_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_weave")).value)
'         .Add Key:="right_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_yarn_count")).value)
'         .Add Key:="sand_roller_type", item:=IIf(Range(wb.Names("sand_roller_type")).value = vbNullString, vbNullString, Range(wb.Names("sand_roller_type")).value)
'         .Add Key:="selvedges_type", item:=IIf(Range(wb.Names("selvedges_type")).value = vbNullString, vbNullString, Range(wb.Names("selvedges_type")).value)
'         .Add Key:="shed_closing_degrees", item:=IIf(Range(wb.Names("shed_closing_degrees")).value = vbNullString, vbNullString, Range(wb.Names("shed_closing_degrees")).value)
'         .Add Key:="speed", item:=IIf(Range(wb.Names("speed")).value = vbNullString, vbNullString, Range(wb.Names("speed")).value)
'         .Add Key:="springs_type", item:=IIf(Range(wb.Names("springs_type")).value = vbNullString, vbNullString, Range(wb.Names("springs_type")).value)
'         .Add Key:="style_number", item:=IIf(Range(wb.Names("style_number")).value = vbNullString, vbNullString, Range(wb.Names("style_number")).value)
'         .Add Key:="temples_composition", item:=IIf(Range(wb.Names("temples_composition")).value = vbNullString, vbNullString, Range(wb.Names("temples_composition")).value)
'         .Add Key:="upper_rapier_clamps", item:=IIf(Range(wb.Names("upper_rapier_clamps")).value = vbNullString, vbNullString, Range(wb.Names("upper_rapier_clamps")).value)
'         .Add Key:="upper_spreader_bars", item:=IIf(Range(wb.Names("upper_spreader_bars")).value = vbNullString, vbNullString, Range(wb.Names("upper_spreader_bars")).value)
'         .Add Key:="vertical_back_rest_roller", item:=IIf(Range(wb.Names("vertical_back_rest_roller")).value = vbNullString, vbNullString, Range(wb.Names("vertical_back_rest_roller")).value)
'         .Add Key:="warp_tension", item:=IIf(Range(wb.Names("warp_tension")).value = vbNullString, vbNullString, Range(wb.Names("warp_tension")).value)
'         .Add Key:="weave_pattern", item:=IIf(Range(wb.Names("weave_pattern")).value = vbNullString, vbNullString, Range(wb.Names("weave_pattern")).value)
'         .Add Key:="weft_count_set_point", item:=IIf(Range(wb.Names("weft_count_set_point")).value = vbNullString, vbNullString, Range(wb.Names("weft_count_set_point")).value)
'         .Add Key:="notes1", item:=IIf(Range(wb.Names("notes1")).value = vbNullString, vbNullString, Range(wb.Names("notes1")).value)
'         .Add Key:="notes2", item:=IIf(Range(wb.Names("notes2")).value = vbNullString, vbNullString, Range(wb.Names("notes2")).value)
'         .Add Key:="notes3", item:=IIf(Range(wb.Names("notes3")).value = vbNullString, vbNullString, Range(wb.Names("notes3")).value)
'         .Add Key:="notes4", item:=IIf(Range(wb.Names("notes4")).value = vbNullString, vbNullString, Range(wb.Names("notes4")).value)
'         .Add Key:="notes5", item:=IIf(Range(wb.Names("notes5")).value = vbNullString, vbNullString, Range(wb.Names("notes5")).value)
'         .Add Key:="notes6", item:=IIf(Range(wb.Names("notes6")).value = vbNullString, vbNullString, Range(wb.Names("notes6")).value)
'         .Add Key:="notes7", item:=IIf(Range(wb.Names("notes7")).value = vbNullString, vbNullString, Range(wb.Names("notes7")).value)
'         .Add Key:="notes8", item:=IIf(Range(wb.Names("notes8")).value = vbNullString, vbNullString, Range(wb.Names("notes8")).value)
'         .Add Key:="roll_length", item:=IIf(Range(wb.Names("roll_length")).value = vbNullString, vbNullString, Range(wb.Names("roll_length")).value)
'         .Add Key:="fd_09", item:=IIf(Range(wb.Names("fd_09")).value = vbNullString, vbNullString, Range(wb.Names("fd_09")).value)
'         .Add Key:="fd_19", item:=IIf(Range(wb.Names("fd_19")).value = vbNullString, vbNullString, Range(wb.Names("fd_19")).value)
'         .Add Key:="fd_29", item:=IIf(Range(wb.Names("fd_29")).value = vbNullString, vbNullString, Range(wb.Names("fd_29")).value)
'         .Add Key:="fd_39", item:=IIf(Range(wb.Names("fd_39")).value = vbNullString, vbNullString, Range(wb.Names("fd_39")).value)
'         .Add Key:="fd_49", item:=IIf(Range(wb.Names("fd_49")).value = vbNullString, vbNullString, Range(wb.Names("fd_49")).value)
'         .Add Key:="fd_59", item:=IIf(Range(wb.Names("fd_59")).value = vbNullString, vbNullString, Range(wb.Names("fd_59")).value)
'         .Add Key:="fd_69", item:=IIf(Range(wb.Names("fd_69")).value = vbNullString, vbNullString, Range(wb.Names("fd_69")).value)
'         .Add Key:="fd_79", item:=IIf(Range(wb.Names("fd_79")).value = vbNullString, vbNullString, Range(wb.Names("fd_79")).value)
'         .Add Key:="fd_89", item:=IIf(Range(wb.Names("fd_89")).value = vbNullString, vbNullString, Range(wb.Names("fd_89")).value)
'         .Add Key:="fd_99", item:=IIf(Range(wb.Names("fd_99")).value = vbNullString, vbNullString, Range(wb.Names("fd_99")).value)
'         .Add Key:="fd_08", item:=IIf(Range(wb.Names("fd_08")).value = vbNullString, vbNullString, Range(wb.Names("fd_08")).value)
'         .Add Key:="fd_18", item:=IIf(Range(wb.Names("fd_18")).value = vbNullString, vbNullString, Range(wb.Names("fd_18")).value)
'         .Add Key:="fd_28", item:=IIf(Range(wb.Names("fd_28")).value = vbNullString, vbNullString, Range(wb.Names("fd_28")).value)
'         .Add Key:="fd_38", item:=IIf(Range(wb.Names("fd_38")).value = vbNullString, vbNullString, Range(wb.Names("fd_38")).value)
'         .Add Key:="fd_48", item:=IIf(Range(wb.Names("fd_48")).value = vbNullString, vbNullString, Range(wb.Names("fd_48")).value)
'         .Add Key:="fd_58", item:=IIf(Range(wb.Names("fd_58")).value = vbNullString, vbNullString, Range(wb.Names("fd_58")).value)
'         .Add Key:="fd_68", item:=IIf(Range(wb.Names("fd_68")).value = vbNullString, vbNullString, Range(wb.Names("fd_68")).value)
'         .Add Key:="fd_78", item:=IIf(Range(wb.Names("fd_78")).value = vbNullString, vbNullString, Range(wb.Names("fd_78")).value)
'         .Add Key:="fd_88", item:=IIf(Range(wb.Names("fd_88")).value = vbNullString, vbNullString, Range(wb.Names("fd_88")).value)
'         .Add Key:="fd_98", item:=IIf(Range(wb.Names("fd_98")).value = vbNullString, vbNullString, Range(wb.Names("fd_98")).value)
'         .Add Key:="fd_07", item:=IIf(Range(wb.Names("fd_07")).value = vbNullString, vbNullString, Range(wb.Names("fd_07")).value)
'         .Add Key:="fd_17", item:=IIf(Range(wb.Names("fd_17")).value = vbNullString, vbNullString, Range(wb.Names("fd_17")).value)
'         .Add Key:="fd_27", item:=IIf(Range(wb.Names("fd_27")).value = vbNullString, vbNullString, Range(wb.Names("fd_27")).value)
'         .Add Key:="fd_37", item:=IIf(Range(wb.Names("fd_37")).value = vbNullString, vbNullString, Range(wb.Names("fd_37")).value)
'         .Add Key:="fd_47", item:=IIf(Range(wb.Names("fd_47")).value = vbNullString, vbNullString, Range(wb.Names("fd_47")).value)
'         .Add Key:="fd_57", item:=IIf(Range(wb.Names("fd_57")).value = vbNullString, vbNullString, Range(wb.Names("fd_57")).value)
'         .Add Key:="fd_67", item:=IIf(Range(wb.Names("fd_67")).value = vbNullString, vbNullString, Range(wb.Names("fd_67")).value)
'         .Add Key:="fd_77", item:=IIf(Range(wb.Names("fd_77")).value = vbNullString, vbNullString, Range(wb.Names("fd_77")).value)
'         .Add Key:="fd_87", item:=IIf(Range(wb.Names("fd_87")).value = vbNullString, vbNullString, Range(wb.Names("fd_87")).value)
'         .Add Key:="fd_97", item:=IIf(Range(wb.Names("fd_97")).value = vbNullString, vbNullString, Range(wb.Names("fd_97")).value)
'         .Add Key:="fd_06", item:=IIf(Range(wb.Names("fd_06")).value = vbNullString, vbNullString, Range(wb.Names("fd_06")).value)
'         .Add Key:="fd_16", item:=IIf(Range(wb.Names("fd_16")).value = vbNullString, vbNullString, Range(wb.Names("fd_16")).value)
'         .Add Key:="fd_26", item:=IIf(Range(wb.Names("fd_26")).value = vbNullString, vbNullString, Range(wb.Names("fd_26")).value)
'         .Add Key:="fd_36", item:=IIf(Range(wb.Names("fd_36")).value = vbNullString, vbNullString, Range(wb.Names("fd_36")).value)
'         .Add Key:="fd_46", item:=IIf(Range(wb.Names("fd_46")).value = vbNullString, vbNullString, Range(wb.Names("fd_46")).value)
'         .Add Key:="fd_56", item:=IIf(Range(wb.Names("fd_56")).value = vbNullString, vbNullString, Range(wb.Names("fd_56")).value)
'         .Add Key:="fd_66", item:=IIf(Range(wb.Names("fd_66")).value = vbNullString, vbNullString, Range(wb.Names("fd_66")).value)
'         .Add Key:="fd_76", item:=IIf(Range(wb.Names("fd_76")).value = vbNullString, vbNullString, Range(wb.Names("fd_76")).value)
'         .Add Key:="fd_86", item:=IIf(Range(wb.Names("fd_86")).value = vbNullString, vbNullString, Range(wb.Names("fd_86")).value)
'         .Add Key:="fd_96", item:=IIf(Range(wb.Names("fd_96")).value = vbNullString, vbNullString, Range(wb.Names("fd_96")).value)
'         .Add Key:="fd_05", item:=IIf(Range(wb.Names("fd_05")).value = vbNullString, vbNullString, Range(wb.Names("fd_05")).value)
'         .Add Key:="fd_15", item:=IIf(Range(wb.Names("fd_15")).value = vbNullString, vbNullString, Range(wb.Names("fd_15")).value)
'         .Add Key:="fd_25", item:=IIf(Range(wb.Names("fd_25")).value = vbNullString, vbNullString, Range(wb.Names("fd_25")).value)
'         .Add Key:="fd_35", item:=IIf(Range(wb.Names("fd_35")).value = vbNullString, vbNullString, Range(wb.Names("fd_35")).value)
'         .Add Key:="fd_45", item:=IIf(Range(wb.Names("fd_45")).value = vbNullString, vbNullString, Range(wb.Names("fd_45")).value)
'         .Add Key:="fd_55", item:=IIf(Range(wb.Names("fd_55")).value = vbNullString, vbNullString, Range(wb.Names("fd_55")).value)
'         .Add Key:="fd_65", item:=IIf(Range(wb.Names("fd_65")).value = vbNullString, vbNullString, Range(wb.Names("fd_65")).value)
'         .Add Key:="fd_75", item:=IIf(Range(wb.Names("fd_75")).value = vbNullString, vbNullString, Range(wb.Names("fd_75")).value)
'         .Add Key:="fd_85", item:=IIf(Range(wb.Names("fd_85")).value = vbNullString, vbNullString, Range(wb.Names("fd_85")).value)
'         .Add Key:="fd_95", item:=IIf(Range(wb.Names("fd_95")).value = vbNullString, vbNullString, Range(wb.Names("fd_95")).value)
'         .Add Key:="fd_04", item:=IIf(Range(wb.Names("fd_04")).value = vbNullString, vbNullString, Range(wb.Names("fd_04")).value)
'         .Add Key:="fd_14", item:=IIf(Range(wb.Names("fd_14")).value = vbNullString, vbNullString, Range(wb.Names("fd_14")).value)
'         .Add Key:="fd_24", item:=IIf(Range(wb.Names("fd_24")).value = vbNullString, vbNullString, Range(wb.Names("fd_24")).value)
'         .Add Key:="fd_34", item:=IIf(Range(wb.Names("fd_34")).value = vbNullString, vbNullString, Range(wb.Names("fd_34")).value)
'         .Add Key:="fd_44", item:=IIf(Range(wb.Names("fd_44")).value = vbNullString, vbNullString, Range(wb.Names("fd_44")).value)
'         .Add Key:="fd_54", item:=IIf(Range(wb.Names("fd_54")).value = vbNullString, vbNullString, Range(wb.Names("fd_54")).value)
'         .Add Key:="fd_64", item:=IIf(Range(wb.Names("fd_64")).value = vbNullString, vbNullString, Range(wb.Names("fd_64")).value)
'         .Add Key:="fd_74", item:=IIf(Range(wb.Names("fd_74")).value = vbNullString, vbNullString, Range(wb.Names("fd_74")).value)
'         .Add Key:="fd_84", item:=IIf(Range(wb.Names("fd_84")).value = vbNullString, vbNullString, Range(wb.Names("fd_84")).value)
'         .Add Key:="fd_94", item:=IIf(Range(wb.Names("fd_94")).value = vbNullString, vbNullString, Range(wb.Names("fd_94")).value)
'         .Add Key:="fd_03", item:=IIf(Range(wb.Names("fd_03")).value = vbNullString, vbNullString, Range(wb.Names("fd_03")).value)
'         .Add Key:="fd_13", item:=IIf(Range(wb.Names("fd_13")).value = vbNullString, vbNullString, Range(wb.Names("fd_13")).value)
'         .Add Key:="fd_23", item:=IIf(Range(wb.Names("fd_23")).value = vbNullString, vbNullString, Range(wb.Names("fd_23")).value)
'         .Add Key:="fd_33", item:=IIf(Range(wb.Names("fd_33")).value = vbNullString, vbNullString, Range(wb.Names("fd_33")).value)
'         .Add Key:="fd_43", item:=IIf(Range(wb.Names("fd_43")).value = vbNullString, vbNullString, Range(wb.Names("fd_43")).value)
'         .Add Key:="fd_53", item:=IIf(Range(wb.Names("fd_53")).value = vbNullString, vbNullString, Range(wb.Names("fd_53")).value)
'         .Add Key:="fd_63", item:=IIf(Range(wb.Names("fd_63")).value = vbNullString, vbNullString, Range(wb.Names("fd_63")).value)
'         .Add Key:="fd_73", item:=IIf(Range(wb.Names("fd_73")).value = vbNullString, vbNullString, Range(wb.Names("fd_73")).value)
'         .Add Key:="fd_83", item:=IIf(Range(wb.Names("fd_83")).value = vbNullString, vbNullString, Range(wb.Names("fd_83")).value)
'         .Add Key:="fd_93", item:=IIf(Range(wb.Names("fd_93")).value = vbNullString, vbNullString, Range(wb.Names("fd_93")).value)
'         .Add Key:="fd_02", item:=IIf(Range(wb.Names("fd_02")).value = vbNullString, vbNullString, Range(wb.Names("fd_02")).value)
'         .Add Key:="fd_12", item:=IIf(Range(wb.Names("fd_12")).value = vbNullString, vbNullString, Range(wb.Names("fd_12")).value)
'         .Add Key:="fd_22", item:=IIf(Range(wb.Names("fd_22")).value = vbNullString, vbNullString, Range(wb.Names("fd_22")).value)
'         .Add Key:="fd_32", item:=IIf(Range(wb.Names("fd_32")).value = vbNullString, vbNullString, Range(wb.Names("fd_32")).value)
'         .Add Key:="fd_42", item:=IIf(Range(wb.Names("fd_42")).value = vbNullString, vbNullString, Range(wb.Names("fd_42")).value)
'         .Add Key:="fd_52", item:=IIf(Range(wb.Names("fd_52")).value = vbNullString, vbNullString, Range(wb.Names("fd_52")).value)
'         .Add Key:="fd_62", item:=IIf(Range(wb.Names("fd_62")).value = vbNullString, vbNullString, Range(wb.Names("fd_62")).value)
'         .Add Key:="fd_72", item:=IIf(Range(wb.Names("fd_72")).value = vbNullString, vbNullString, Range(wb.Names("fd_72")).value)
'         .Add Key:="fd_82", item:=IIf(Range(wb.Names("fd_82")).value = vbNullString, vbNullString, Range(wb.Names("fd_82")).value)
'         .Add Key:="fd_92", item:=IIf(Range(wb.Names("fd_92")).value = vbNullString, vbNullString, Range(wb.Names("fd_92")).value)
'         .Add Key:="fd_01", item:=IIf(Range(wb.Names("fd_01")).value = vbNullString, vbNullString, Range(wb.Names("fd_01")).value)
'         .Add Key:="fd_11", item:=IIf(Range(wb.Names("fd_11")).value = vbNullString, vbNullString, Range(wb.Names("fd_11")).value)
'         .Add Key:="fd_21", item:=IIf(Range(wb.Names("fd_21")).value = vbNullString, vbNullString, Range(wb.Names("fd_21")).value)
'         .Add Key:="fd_31", item:=IIf(Range(wb.Names("fd_31")).value = vbNullString, vbNullString, Range(wb.Names("fd_31")).value)
'         .Add Key:="fd_41", item:=IIf(Range(wb.Names("fd_41")).value = vbNullString, vbNullString, Range(wb.Names("fd_41")).value)
'         .Add Key:="fd_51", item:=IIf(Range(wb.Names("fd_51")).value = vbNullString, vbNullString, Range(wb.Names("fd_51")).value)
'         .Add Key:="fd_61", item:=IIf(Range(wb.Names("fd_61")).value = vbNullString, vbNullString, Range(wb.Names("fd_61")).value)
'         .Add Key:="fd_71", item:=IIf(Range(wb.Names("fd_71")).value = vbNullString, vbNullString, Range(wb.Names("fd_71")).value)
'         .Add Key:="fd_81", item:=IIf(Range(wb.Names("fd_81")).value = vbNullString, vbNullString, Range(wb.Names("fd_81")).value)
'         .Add Key:="fd_91", item:=IIf(Range(wb.Names("fd_91")).value = vbNullString, vbNullString, Range(wb.Names("fd_91")).value)
'         .Add Key:="fd_00", item:=IIf(Range(wb.Names("fd_00")).value = vbNullString, vbNullString, Range(wb.Names("fd_00")).value)
'         .Add Key:="fd_10", item:=IIf(Range(wb.Names("fd_10")).value = vbNullString, vbNullString, Range(wb.Names("fd_10")).value)
'         .Add Key:="fd_20", item:=IIf(Range(wb.Names("fd_20")).value = vbNullString, vbNullString, Range(wb.Names("fd_20")).value)
'         .Add Key:="fd_30", item:=IIf(Range(wb.Names("fd_30")).value = vbNullString, vbNullString, Range(wb.Names("fd_30")).value)
'         .Add Key:="fd_40", item:=IIf(Range(wb.Names("fd_40")).value = vbNullString, vbNullString, Range(wb.Names("fd_40")).value)
'         .Add Key:="fd_50", item:=IIf(Range(wb.Names("fd_50")).value = vbNullString, vbNullString, Range(wb.Names("fd_50")).value)
'         .Add Key:="fd_60", item:=IIf(Range(wb.Names("fd_60")).value = vbNullString, vbNullString, Range(wb.Names("fd_60")).value)
'         .Add Key:="fd_70", item:=IIf(Range(wb.Names("fd_70")).value = vbNullString, vbNullString, Range(wb.Names("fd_70")).value)
'         .Add Key:="fd_80", item:=IIf(Range(wb.Names("fd_80")).value = vbNullString, vbNullString, Range(wb.Names("fd_80")).value)
'         .Add Key:="fd_90", item:=IIf(Range(wb.Names("fd_90")).value = vbNullString, vbNullString, Range(wb.Names("fd_90")).value)
'         .Add Key:="di_09", item:=IIf(Range(wb.Names("di_09")).value = vbNullString, vbNullString, Range(wb.Names("di_09")).value)
'         .Add Key:="di_19", item:=IIf(Range(wb.Names("di_19")).value = vbNullString, vbNullString, Range(wb.Names("di_19")).value)
'         .Add Key:="di_29", item:=IIf(Range(wb.Names("di_29")).value = vbNullString, vbNullString, Range(wb.Names("di_29")).value)
'         .Add Key:="di_39", item:=IIf(Range(wb.Names("di_39")).value = vbNullString, vbNullString, Range(wb.Names("di_39")).value)
'         .Add Key:="di_49", item:=IIf(Range(wb.Names("di_49")).value = vbNullString, vbNullString, Range(wb.Names("di_49")).value)
'         .Add Key:="di_59", item:=IIf(Range(wb.Names("di_59")).value = vbNullString, vbNullString, Range(wb.Names("di_59")).value)
'         .Add Key:="di_69", item:=IIf(Range(wb.Names("di_69")).value = vbNullString, vbNullString, Range(wb.Names("di_69")).value)
'         .Add Key:="di_79", item:=IIf(Range(wb.Names("di_79")).value = vbNullString, vbNullString, Range(wb.Names("di_79")).value)
'         .Add Key:="di_89", item:=IIf(Range(wb.Names("di_89")).value = vbNullString, vbNullString, Range(wb.Names("di_89")).value)
'         .Add Key:="di_99", item:=IIf(Range(wb.Names("di_99")).value = vbNullString, vbNullString, Range(wb.Names("di_99")).value)
'         .Add Key:="di_08", item:=IIf(Range(wb.Names("di_08")).value = vbNullString, vbNullString, Range(wb.Names("di_08")).value)
'         .Add Key:="di_18", item:=IIf(Range(wb.Names("di_18")).value = vbNullString, vbNullString, Range(wb.Names("di_18")).value)
'         .Add Key:="di_28", item:=IIf(Range(wb.Names("di_28")).value = vbNullString, vbNullString, Range(wb.Names("di_28")).value)
'         .Add Key:="di_38", item:=IIf(Range(wb.Names("di_38")).value = vbNullString, vbNullString, Range(wb.Names("di_38")).value)
'         .Add Key:="di_48", item:=IIf(Range(wb.Names("di_48")).value = vbNullString, vbNullString, Range(wb.Names("di_48")).value)
'         .Add Key:="di_58", item:=IIf(Range(wb.Names("di_58")).value = vbNullString, vbNullString, Range(wb.Names("di_58")).value)
'         .Add Key:="di_68", item:=IIf(Range(wb.Names("di_68")).value = vbNullString, vbNullString, Range(wb.Names("di_68")).value)
'         .Add Key:="di_78", item:=IIf(Range(wb.Names("di_78")).value = vbNullString, vbNullString, Range(wb.Names("di_78")).value)
'         .Add Key:="di_88", item:=IIf(Range(wb.Names("di_88")).value = vbNullString, vbNullString, Range(wb.Names("di_88")).value)
'         .Add Key:="di_98", item:=IIf(Range(wb.Names("di_98")).value = vbNullString, vbNullString, Range(wb.Names("di_98")).value)
'         .Add Key:="di_07", item:=IIf(Range(wb.Names("di_07")).value = vbNullString, vbNullString, Range(wb.Names("di_07")).value)
'         .Add Key:="di_17", item:=IIf(Range(wb.Names("di_17")).value = vbNullString, vbNullString, Range(wb.Names("di_17")).value)
'         .Add Key:="di_27", item:=IIf(Range(wb.Names("di_27")).value = vbNullString, vbNullString, Range(wb.Names("di_27")).value)
'         .Add Key:="di_37", item:=IIf(Range(wb.Names("di_37")).value = vbNullString, vbNullString, Range(wb.Names("di_37")).value)
'         .Add Key:="di_47", item:=IIf(Range(wb.Names("di_47")).value = vbNullString, vbNullString, Range(wb.Names("di_47")).value)
'         .Add Key:="di_57", item:=IIf(Range(wb.Names("di_57")).value = vbNullString, vbNullString, Range(wb.Names("di_57")).value)
'         .Add Key:="di_67", item:=IIf(Range(wb.Names("di_67")).value = vbNullString, vbNullString, Range(wb.Names("di_67")).value)
'         .Add Key:="di_77", item:=IIf(Range(wb.Names("di_77")).value = vbNullString, vbNullString, Range(wb.Names("di_77")).value)
'         .Add Key:="di_87", item:=IIf(Range(wb.Names("di_87")).value = vbNullString, vbNullString, Range(wb.Names("di_87")).value)
'         .Add Key:="di_97", item:=IIf(Range(wb.Names("di_97")).value = vbNullString, vbNullString, Range(wb.Names("di_97")).value)
'         .Add Key:="di_06", item:=IIf(Range(wb.Names("di_06")).value = vbNullString, vbNullString, Range(wb.Names("di_06")).value)
'         .Add Key:="di_16", item:=IIf(Range(wb.Names("di_16")).value = vbNullString, vbNullString, Range(wb.Names("di_16")).value)
'         .Add Key:="di_26", item:=IIf(Range(wb.Names("di_26")).value = vbNullString, vbNullString, Range(wb.Names("di_26")).value)
'         .Add Key:="di_36", item:=IIf(Range(wb.Names("di_36")).value = vbNullString, vbNullString, Range(wb.Names("di_36")).value)
'         .Add Key:="di_46", item:=IIf(Range(wb.Names("di_46")).value = vbNullString, vbNullString, Range(wb.Names("di_46")).value)
'         .Add Key:="di_56", item:=IIf(Range(wb.Names("di_56")).value = vbNullString, vbNullString, Range(wb.Names("di_56")).value)
'         .Add Key:="di_66", item:=IIf(Range(wb.Names("di_66")).value = vbNullString, vbNullString, Range(wb.Names("di_66")).value)
'         .Add Key:="di_76", item:=IIf(Range(wb.Names("di_76")).value = vbNullString, vbNullString, Range(wb.Names("di_76")).value)
'         .Add Key:="di_86", item:=IIf(Range(wb.Names("di_86")).value = vbNullString, vbNullString, Range(wb.Names("di_86")).value)
'         .Add Key:="di_96", item:=IIf(Range(wb.Names("di_96")).value = vbNullString, vbNullString, Range(wb.Names("di_96")).value)
'         .Add Key:="di_05", item:=IIf(Range(wb.Names("di_05")).value = vbNullString, vbNullString, Range(wb.Names("di_05")).value)
'         .Add Key:="di_15", item:=IIf(Range(wb.Names("di_15")).value = vbNullString, vbNullString, Range(wb.Names("di_15")).value)
'         .Add Key:="di_25", item:=IIf(Range(wb.Names("di_25")).value = vbNullString, vbNullString, Range(wb.Names("di_25")).value)
'         .Add Key:="di_35", item:=IIf(Range(wb.Names("di_35")).value = vbNullString, vbNullString, Range(wb.Names("di_35")).value)
'         .Add Key:="di_45", item:=IIf(Range(wb.Names("di_45")).value = vbNullString, vbNullString, Range(wb.Names("di_45")).value)
'         .Add Key:="di_55", item:=IIf(Range(wb.Names("di_55")).value = vbNullString, vbNullString, Range(wb.Names("di_55")).value)
'         .Add Key:="di_65", item:=IIf(Range(wb.Names("di_65")).value = vbNullString, vbNullString, Range(wb.Names("di_65")).value)
'         .Add Key:="di_75", item:=IIf(Range(wb.Names("di_75")).value = vbNullString, vbNullString, Range(wb.Names("di_75")).value)
'         .Add Key:="di_85", item:=IIf(Range(wb.Names("di_85")).value = vbNullString, vbNullString, Range(wb.Names("di_85")).value)
'         .Add Key:="di_95", item:=IIf(Range(wb.Names("di_95")).value = vbNullString, vbNullString, Range(wb.Names("di_95")).value)
'         .Add Key:="di_04", item:=IIf(Range(wb.Names("di_04")).value = vbNullString, vbNullString, Range(wb.Names("di_04")).value)
'         .Add Key:="di_14", item:=IIf(Range(wb.Names("di_14")).value = vbNullString, vbNullString, Range(wb.Names("di_14")).value)
'         .Add Key:="di_24", item:=IIf(Range(wb.Names("di_24")).value = vbNullString, vbNullString, Range(wb.Names("di_24")).value)
'         .Add Key:="di_34", item:=IIf(Range(wb.Names("di_34")).value = vbNullString, vbNullString, Range(wb.Names("di_34")).value)
'         .Add Key:="di_44", item:=IIf(Range(wb.Names("di_44")).value = vbNullString, vbNullString, Range(wb.Names("di_44")).value)
'         .Add Key:="di_54", item:=IIf(Range(wb.Names("di_54")).value = vbNullString, vbNullString, Range(wb.Names("di_54")).value)
'         .Add Key:="di_64", item:=IIf(Range(wb.Names("di_64")).value = vbNullString, vbNullString, Range(wb.Names("di_64")).value)
'         .Add Key:="di_74", item:=IIf(Range(wb.Names("di_74")).value = vbNullString, vbNullString, Range(wb.Names("di_74")).value)
'         .Add Key:="di_84", item:=IIf(Range(wb.Names("di_84")).value = vbNullString, vbNullString, Range(wb.Names("di_84")).value)
'         .Add Key:="di_94", item:=IIf(Range(wb.Names("di_94")).value = vbNullString, vbNullString, Range(wb.Names("di_94")).value)
'         .Add Key:="di_03", item:=IIf(Range(wb.Names("di_03")).value = vbNullString, vbNullString, Range(wb.Names("di_03")).value)
'         .Add Key:="di_13", item:=IIf(Range(wb.Names("di_13")).value = vbNullString, vbNullString, Range(wb.Names("di_13")).value)
'         .Add Key:="di_23", item:=IIf(Range(wb.Names("di_23")).value = vbNullString, vbNullString, Range(wb.Names("di_23")).value)
'         .Add Key:="di_33", item:=IIf(Range(wb.Names("di_33")).value = vbNullString, vbNullString, Range(wb.Names("di_33")).value)
'         .Add Key:="di_43", item:=IIf(Range(wb.Names("di_43")).value = vbNullString, vbNullString, Range(wb.Names("di_43")).value)
'         .Add Key:="di_53", item:=IIf(Range(wb.Names("di_53")).value = vbNullString, vbNullString, Range(wb.Names("di_53")).value)
'         .Add Key:="di_63", item:=IIf(Range(wb.Names("di_63")).value = vbNullString, vbNullString, Range(wb.Names("di_63")).value)
'         .Add Key:="di_73", item:=IIf(Range(wb.Names("di_73")).value = vbNullString, vbNullString, Range(wb.Names("di_73")).value)
'         .Add Key:="di_83", item:=IIf(Range(wb.Names("di_83")).value = vbNullString, vbNullString, Range(wb.Names("di_83")).value)
'         .Add Key:="di_93", item:=IIf(Range(wb.Names("di_93")).value = vbNullString, vbNullString, Range(wb.Names("di_93")).value)
'         .Add Key:="di_02", item:=IIf(Range(wb.Names("di_02")).value = vbNullString, vbNullString, Range(wb.Names("di_02")).value)
'         .Add Key:="di_12", item:=IIf(Range(wb.Names("di_12")).value = vbNullString, vbNullString, Range(wb.Names("di_12")).value)
'         .Add Key:="di_22", item:=IIf(Range(wb.Names("di_22")).value = vbNullString, vbNullString, Range(wb.Names("di_22")).value)
'         .Add Key:="di_32", item:=IIf(Range(wb.Names("di_32")).value = vbNullString, vbNullString, Range(wb.Names("di_32")).value)
'         .Add Key:="di_42", item:=IIf(Range(wb.Names("di_42")).value = vbNullString, vbNullString, Range(wb.Names("di_42")).value)
'         .Add Key:="di_52", item:=IIf(Range(wb.Names("di_52")).value = vbNullString, vbNullString, Range(wb.Names("di_52")).value)
'         .Add Key:="di_62", item:=IIf(Range(wb.Names("di_62")).value = vbNullString, vbNullString, Range(wb.Names("di_62")).value)
'         .Add Key:="di_72", item:=IIf(Range(wb.Names("di_72")).value = vbNullString, vbNullString, Range(wb.Names("di_72")).value)
'         .Add Key:="di_82", item:=IIf(Range(wb.Names("di_82")).value = vbNullString, vbNullString, Range(wb.Names("di_82")).value)
'         .Add Key:="di_92", item:=IIf(Range(wb.Names("di_92")).value = vbNullString, vbNullString, Range(wb.Names("di_92")).value)
'         .Add Key:="di_01", item:=IIf(Range(wb.Names("di_01")).value = vbNullString, vbNullString, Range(wb.Names("di_01")).value)
'         .Add Key:="di_11", item:=IIf(Range(wb.Names("di_11")).value = vbNullString, vbNullString, Range(wb.Names("di_11")).value)
'         .Add Key:="di_21", item:=IIf(Range(wb.Names("di_21")).value = vbNullString, vbNullString, Range(wb.Names("di_21")).value)
'         .Add Key:="di_31", item:=IIf(Range(wb.Names("di_31")).value = vbNullString, vbNullString, Range(wb.Names("di_31")).value)
'         .Add Key:="di_41", item:=IIf(Range(wb.Names("di_41")).value = vbNullString, vbNullString, Range(wb.Names("di_41")).value)
'         .Add Key:="di_51", item:=IIf(Range(wb.Names("di_51")).value = vbNullString, vbNullString, Range(wb.Names("di_51")).value)
'         .Add Key:="di_61", item:=IIf(Range(wb.Names("di_61")).value = vbNullString, vbNullString, Range(wb.Names("di_61")).value)
'         .Add Key:="di_71", item:=IIf(Range(wb.Names("di_71")).value = vbNullString, vbNullString, Range(wb.Names("di_71")).value)
'         .Add Key:="di_81", item:=IIf(Range(wb.Names("di_81")).value = vbNullString, vbNullString, Range(wb.Names("di_81")).value)
'         .Add Key:="di_91", item:=IIf(Range(wb.Names("di_91")).value = vbNullString, vbNullString, Range(wb.Names("di_91")).value)
'         .Add Key:="di_00", item:=IIf(Range(wb.Names("di_00")).value = vbNullString, vbNullString, Range(wb.Names("di_00")).value)
'         .Add Key:="di_10", item:=IIf(Range(wb.Names("di_10")).value = vbNullString, vbNullString, Range(wb.Names("di_10")).value)
'         .Add Key:="di_20", item:=IIf(Range(wb.Names("di_20")).value = vbNullString, vbNullString, Range(wb.Names("di_20")).value)
'         .Add Key:="di_30", item:=IIf(Range(wb.Names("di_30")).value = vbNullString, vbNullString, Range(wb.Names("di_30")).value)
'         .Add Key:="di_40", item:=IIf(Range(wb.Names("di_40")).value = vbNullString, vbNullString, Range(wb.Names("di_40")).value)
'         .Add Key:="di_50", item:=IIf(Range(wb.Names("di_50")).value = vbNullString, vbNullString, Range(wb.Names("di_50")).value)
'         .Add Key:="di_60", item:=IIf(Range(wb.Names("di_60")).value = vbNullString, vbNullString, Range(wb.Names("di_60")).value)
'         .Add Key:="di_70", item:=IIf(Range(wb.Names("di_70")).value = vbNullString, vbNullString, Range(wb.Names("di_70")).value)
'         .Add Key:="di_80", item:=IIf(Range(wb.Names("di_80")).value = vbNullString, vbNullString, Range(wb.Names("di_80")).value)
'         .Add Key:="di_90", item:=IIf(Range(wb.Names("di_90")).value = vbNullString, vbNullString, Range(wb.Names("di_90")).value)
'         .Add Key:="ld_09", item:=IIf(Range(wb.Names("ld_09")).value = vbNullString, vbNullString, Range(wb.Names("ld_09")).value)
'         .Add Key:="ld_19", item:=IIf(Range(wb.Names("ld_19")).value = vbNullString, vbNullString, Range(wb.Names("ld_19")).value)
'         .Add Key:="ld_29", item:=IIf(Range(wb.Names("ld_29")).value = vbNullString, vbNullString, Range(wb.Names("ld_29")).value)
'         .Add Key:="ld_39", item:=IIf(Range(wb.Names("ld_39")).value = vbNullString, vbNullString, Range(wb.Names("ld_39")).value)
'         .Add Key:="ld_49", item:=IIf(Range(wb.Names("ld_49")).value = vbNullString, vbNullString, Range(wb.Names("ld_49")).value)
'         .Add Key:="ld_59", item:=IIf(Range(wb.Names("ld_59")).value = vbNullString, vbNullString, Range(wb.Names("ld_59")).value)
'         .Add Key:="ld_69", item:=IIf(Range(wb.Names("ld_69")).value = vbNullString, vbNullString, Range(wb.Names("ld_69")).value)
'         .Add Key:="ld_79", item:=IIf(Range(wb.Names("ld_79")).value = vbNullString, vbNullString, Range(wb.Names("ld_79")).value)
'         .Add Key:="ld_89", item:=IIf(Range(wb.Names("ld_89")).value = vbNullString, vbNullString, Range(wb.Names("ld_89")).value)
'         .Add Key:="ld_99", item:=IIf(Range(wb.Names("ld_99")).value = vbNullString, vbNullString, Range(wb.Names("ld_99")).value)
'         .Add Key:="ld_08", item:=IIf(Range(wb.Names("ld_08")).value = vbNullString, vbNullString, Range(wb.Names("ld_08")).value)
'         .Add Key:="ld_18", item:=IIf(Range(wb.Names("ld_18")).value = vbNullString, vbNullString, Range(wb.Names("ld_18")).value)
'         .Add Key:="ld_28", item:=IIf(Range(wb.Names("ld_28")).value = vbNullString, vbNullString, Range(wb.Names("ld_28")).value)
'         .Add Key:="ld_38", item:=IIf(Range(wb.Names("ld_38")).value = vbNullString, vbNullString, Range(wb.Names("ld_38")).value)
'         .Add Key:="ld_48", item:=IIf(Range(wb.Names("ld_48")).value = vbNullString, vbNullString, Range(wb.Names("ld_48")).value)
'         .Add Key:="ld_58", item:=IIf(Range(wb.Names("ld_58")).value = vbNullString, vbNullString, Range(wb.Names("ld_58")).value)
'         .Add Key:="ld_68", item:=IIf(Range(wb.Names("ld_68")).value = vbNullString, vbNullString, Range(wb.Names("ld_68")).value)
'         .Add Key:="ld_78", item:=IIf(Range(wb.Names("ld_78")).value = vbNullString, vbNullString, Range(wb.Names("ld_78")).value)
'         .Add Key:="ld_88", item:=IIf(Range(wb.Names("ld_88")).value = vbNullString, vbNullString, Range(wb.Names("ld_88")).value)
'         .Add Key:="ld_98", item:=IIf(Range(wb.Names("ld_98")).value = vbNullString, vbNullString, Range(wb.Names("ld_98")).value)
'         .Add Key:="ld_07", item:=IIf(Range(wb.Names("ld_07")).value = vbNullString, vbNullString, Range(wb.Names("ld_07")).value)
'         .Add Key:="ld_17", item:=IIf(Range(wb.Names("ld_17")).value = vbNullString, vbNullString, Range(wb.Names("ld_17")).value)
'         .Add Key:="ld_27", item:=IIf(Range(wb.Names("ld_27")).value = vbNullString, vbNullString, Range(wb.Names("ld_27")).value)
'         .Add Key:="ld_37", item:=IIf(Range(wb.Names("ld_37")).value = vbNullString, vbNullString, Range(wb.Names("ld_37")).value)
'         .Add Key:="ld_47", item:=IIf(Range(wb.Names("ld_47")).value = vbNullString, vbNullString, Range(wb.Names("ld_47")).value)
'         .Add Key:="ld_57", item:=IIf(Range(wb.Names("ld_57")).value = vbNullString, vbNullString, Range(wb.Names("ld_57")).value)
'         .Add Key:="ld_67", item:=IIf(Range(wb.Names("ld_67")).value = vbNullString, vbNullString, Range(wb.Names("ld_67")).value)
'         .Add Key:="ld_77", item:=IIf(Range(wb.Names("ld_77")).value = vbNullString, vbNullString, Range(wb.Names("ld_77")).value)
'         .Add Key:="ld_87", item:=IIf(Range(wb.Names("ld_87")).value = vbNullString, vbNullString, Range(wb.Names("ld_87")).value)
'         .Add Key:="ld_97", item:=IIf(Range(wb.Names("ld_97")).value = vbNullString, vbNullString, Range(wb.Names("ld_97")).value)
'         .Add Key:="ld_06", item:=IIf(Range(wb.Names("ld_06")).value = vbNullString, vbNullString, Range(wb.Names("ld_06")).value)
'         .Add Key:="ld_16", item:=IIf(Range(wb.Names("ld_16")).value = vbNullString, vbNullString, Range(wb.Names("ld_16")).value)
'         .Add Key:="ld_26", item:=IIf(Range(wb.Names("ld_26")).value = vbNullString, vbNullString, Range(wb.Names("ld_26")).value)
'         .Add Key:="ld_36", item:=IIf(Range(wb.Names("ld_36")).value = vbNullString, vbNullString, Range(wb.Names("ld_36")).value)
'         .Add Key:="ld_46", item:=IIf(Range(wb.Names("ld_46")).value = vbNullString, vbNullString, Range(wb.Names("ld_46")).value)
'         .Add Key:="ld_56", item:=IIf(Range(wb.Names("ld_56")).value = vbNullString, vbNullString, Range(wb.Names("ld_56")).value)
'         .Add Key:="ld_66", item:=IIf(Range(wb.Names("ld_66")).value = vbNullString, vbNullString, Range(wb.Names("ld_66")).value)
'         .Add Key:="ld_76", item:=IIf(Range(wb.Names("ld_76")).value = vbNullString, vbNullString, Range(wb.Names("ld_76")).value)
'         .Add Key:="ld_86", item:=IIf(Range(wb.Names("ld_86")).value = vbNullString, vbNullString, Range(wb.Names("ld_86")).value)
'         .Add Key:="ld_96", item:=IIf(Range(wb.Names("ld_96")).value = vbNullString, vbNullString, Range(wb.Names("ld_96")).value)
'         .Add Key:="ld_05", item:=IIf(Range(wb.Names("ld_05")).value = vbNullString, vbNullString, Range(wb.Names("ld_05")).value)
'         .Add Key:="ld_15", item:=IIf(Range(wb.Names("ld_15")).value = vbNullString, vbNullString, Range(wb.Names("ld_15")).value)
'         .Add Key:="ld_25", item:=IIf(Range(wb.Names("ld_25")).value = vbNullString, vbNullString, Range(wb.Names("ld_25")).value)
'         .Add Key:="ld_35", item:=IIf(Range(wb.Names("ld_35")).value = vbNullString, vbNullString, Range(wb.Names("ld_35")).value)
'         .Add Key:="ld_45", item:=IIf(Range(wb.Names("ld_45")).value = vbNullString, vbNullString, Range(wb.Names("ld_45")).value)
'         .Add Key:="ld_55", item:=IIf(Range(wb.Names("ld_55")).value = vbNullString, vbNullString, Range(wb.Names("ld_55")).value)
'         .Add Key:="ld_65", item:=IIf(Range(wb.Names("ld_65")).value = vbNullString, vbNullString, Range(wb.Names("ld_65")).value)
'         .Add Key:="ld_75", item:=IIf(Range(wb.Names("ld_75")).value = vbNullString, vbNullString, Range(wb.Names("ld_75")).value)
'         .Add Key:="ld_85", item:=IIf(Range(wb.Names("ld_85")).value = vbNullString, vbNullString, Range(wb.Names("ld_85")).value)
'         .Add Key:="ld_95", item:=IIf(Range(wb.Names("ld_95")).value = vbNullString, vbNullString, Range(wb.Names("ld_95")).value)
'         .Add Key:="ld_04", item:=IIf(Range(wb.Names("ld_04")).value = vbNullString, vbNullString, Range(wb.Names("ld_04")).value)
'         .Add Key:="ld_14", item:=IIf(Range(wb.Names("ld_14")).value = vbNullString, vbNullString, Range(wb.Names("ld_14")).value)
'         .Add Key:="ld_24", item:=IIf(Range(wb.Names("ld_24")).value = vbNullString, vbNullString, Range(wb.Names("ld_24")).value)
'         .Add Key:="ld_34", item:=IIf(Range(wb.Names("ld_34")).value = vbNullString, vbNullString, Range(wb.Names("ld_34")).value)
'         .Add Key:="ld_44", item:=IIf(Range(wb.Names("ld_44")).value = vbNullString, vbNullString, Range(wb.Names("ld_44")).value)
'         .Add Key:="ld_54", item:=IIf(Range(wb.Names("ld_54")).value = vbNullString, vbNullString, Range(wb.Names("ld_54")).value)
'         .Add Key:="ld_64", item:=IIf(Range(wb.Names("ld_64")).value = vbNullString, vbNullString, Range(wb.Names("ld_64")).value)
'         .Add Key:="ld_74", item:=IIf(Range(wb.Names("ld_74")).value = vbNullString, vbNullString, Range(wb.Names("ld_74")).value)
'         .Add Key:="ld_84", item:=IIf(Range(wb.Names("ld_84")).value = vbNullString, vbNullString, Range(wb.Names("ld_84")).value)
'         .Add Key:="ld_94", item:=IIf(Range(wb.Names("ld_94")).value = vbNullString, vbNullString, Range(wb.Names("ld_94")).value)
'         .Add Key:="ld_03", item:=IIf(Range(wb.Names("ld_03")).value = vbNullString, vbNullString, Range(wb.Names("ld_03")).value)
'         .Add Key:="ld_13", item:=IIf(Range(wb.Names("ld_13")).value = vbNullString, vbNullString, Range(wb.Names("ld_13")).value)
'         .Add Key:="ld_23", item:=IIf(Range(wb.Names("ld_23")).value = vbNullString, vbNullString, Range(wb.Names("ld_23")).value)
'         .Add Key:="ld_33", item:=IIf(Range(wb.Names("ld_33")).value = vbNullString, vbNullString, Range(wb.Names("ld_33")).value)
'         .Add Key:="ld_43", item:=IIf(Range(wb.Names("ld_43")).value = vbNullString, vbNullString, Range(wb.Names("ld_43")).value)
'         .Add Key:="ld_53", item:=IIf(Range(wb.Names("ld_53")).value = vbNullString, vbNullString, Range(wb.Names("ld_53")).value)
'         .Add Key:="ld_63", item:=IIf(Range(wb.Names("ld_63")).value = vbNullString, vbNullString, Range(wb.Names("ld_63")).value)
'         .Add Key:="ld_73", item:=IIf(Range(wb.Names("ld_73")).value = vbNullString, vbNullString, Range(wb.Names("ld_73")).value)
'         .Add Key:="ld_83", item:=IIf(Range(wb.Names("ld_83")).value = vbNullString, vbNullString, Range(wb.Names("ld_83")).value)
'         .Add Key:="ld_93", item:=IIf(Range(wb.Names("ld_93")).value = vbNullString, vbNullString, Range(wb.Names("ld_93")).value)
'         .Add Key:="ld_02", item:=IIf(Range(wb.Names("ld_02")).value = vbNullString, vbNullString, Range(wb.Names("ld_02")).value)
'         .Add Key:="ld_12", item:=IIf(Range(wb.Names("ld_12")).value = vbNullString, vbNullString, Range(wb.Names("ld_12")).value)
'         .Add Key:="ld_22", item:=IIf(Range(wb.Names("ld_22")).value = vbNullString, vbNullString, Range(wb.Names("ld_22")).value)
'         .Add Key:="ld_32", item:=IIf(Range(wb.Names("ld_32")).value = vbNullString, vbNullString, Range(wb.Names("ld_32")).value)
'         .Add Key:="ld_42", item:=IIf(Range(wb.Names("ld_42")).value = vbNullString, vbNullString, Range(wb.Names("ld_42")).value)
'         .Add Key:="ld_52", item:=IIf(Range(wb.Names("ld_52")).value = vbNullString, vbNullString, Range(wb.Names("ld_52")).value)
'         .Add Key:="ld_62", item:=IIf(Range(wb.Names("ld_62")).value = vbNullString, vbNullString, Range(wb.Names("ld_62")).value)
'         .Add Key:="ld_72", item:=IIf(Range(wb.Names("ld_72")).value = vbNullString, vbNullString, Range(wb.Names("ld_72")).value)
'         .Add Key:="ld_82", item:=IIf(Range(wb.Names("ld_82")).value = vbNullString, vbNullString, Range(wb.Names("ld_82")).value)
'         .Add Key:="ld_92", item:=IIf(Range(wb.Names("ld_92")).value = vbNullString, vbNullString, Range(wb.Names("ld_92")).value)
'         .Add Key:="ld_01", item:=IIf(Range(wb.Names("ld_01")).value = vbNullString, vbNullString, Range(wb.Names("ld_01")).value)
'         .Add Key:="ld_11", item:=IIf(Range(wb.Names("ld_11")).value = vbNullString, vbNullString, Range(wb.Names("ld_11")).value)
'         .Add Key:="ld_21", item:=IIf(Range(wb.Names("ld_21")).value = vbNullString, vbNullString, Range(wb.Names("ld_21")).value)
'         .Add Key:="ld_31", item:=IIf(Range(wb.Names("ld_31")).value = vbNullString, vbNullString, Range(wb.Names("ld_31")).value)
'         .Add Key:="ld_41", item:=IIf(Range(wb.Names("ld_41")).value = vbNullString, vbNullString, Range(wb.Names("ld_41")).value)
'         .Add Key:="ld_51", item:=IIf(Range(wb.Names("ld_51")).value = vbNullString, vbNullString, Range(wb.Names("ld_51")).value)
'         .Add Key:="ld_61", item:=IIf(Range(wb.Names("ld_61")).value = vbNullString, vbNullString, Range(wb.Names("ld_61")).value)
'         .Add Key:="ld_71", item:=IIf(Range(wb.Names("ld_71")).value = vbNullString, vbNullString, Range(wb.Names("ld_71")).value)
'         .Add Key:="ld_81", item:=IIf(Range(wb.Names("ld_81")).value = vbNullString, vbNullString, Range(wb.Names("ld_81")).value)
'         .Add Key:="ld_91", item:=IIf(Range(wb.Names("ld_91")).value = vbNullString, vbNullString, Range(wb.Names("ld_91")).value)
'         .Add Key:="ld_00", item:=IIf(Range(wb.Names("ld_00")).value = vbNullString, vbNullString, Range(wb.Names("ld_00")).value)
'         .Add Key:="ld_10", item:=IIf(Range(wb.Names("ld_10")).value = vbNullString, vbNullString, Range(wb.Names("ld_10")).value)
'         .Add Key:="ld_20", item:=IIf(Range(wb.Names("ld_20")).value = vbNullString, vbNullString, Range(wb.Names("ld_20")).value)
'         .Add Key:="ld_30", item:=IIf(Range(wb.Names("ld_30")).value = vbNullString, vbNullString, Range(wb.Names("ld_30")).value)
'         .Add Key:="ld_40", item:=IIf(Range(wb.Names("ld_40")).value = vbNullString, vbNullString, Range(wb.Names("ld_40")).value)
'         .Add Key:="ld_50", item:=IIf(Range(wb.Names("ld_50")).value = vbNullString, vbNullString, Range(wb.Names("ld_50")).value)
'         .Add Key:="ld_60", item:=IIf(Range(wb.Names("ld_60")).value = vbNullString, vbNullString, Range(wb.Names("ld_60")).value)
'         .Add Key:="ld_70", item:=IIf(Range(wb.Names("ld_70")).value = vbNullString, vbNullString, Range(wb.Names("ld_70")).value)
'         .Add Key:="ld_80", item:=IIf(Range(wb.Names("ld_80")).value = vbNullString, vbNullString, Range(wb.Names("ld_80")).value)
'         .Add Key:="ld_90", item:=IIf(Range(wb.Names("ld_90")).value = vbNullString, vbNullString, Range(wb.Names("ld_90")).value)
'     End With
'     Set AddMoreRbaNames = dict
' End Function

Public Function AddRbaNames(dict As Object, wb As Workbook, tag As String, r_start As Long, r_end As Long, c_start As Long, c_end As Long) As Object
    Dim sht As Worksheet
    Dim nr As String
    Dim ret_val As Variant
    Set sht = wb.Sheets("Weaving RBA")
    Dim r, C, rw, cl As Long
    For r = r_start To r_end
        cl = 0
        For C = c_start To c_end
            rw = Abs(r_end - r)
            nr = tag & "_" & cl & rw
            ret_val = CreateNamedRange(wb, nr, sht, CLng(r), CLng(C))
            dict.Add nr, IIf(ret_val = vbNullString, vbNullString, ret_val)
            cl = cl + 1
        Next C
    Next r
    Set AddRbaNames = dict
End Function

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
