Attribute VB_Name = "RbaParser"
Option Explicit
' ==============================================
' RBA PARSER
' ==============================================
Public Sub ParseAll(Optional material_keyword As String = "code")
    Dim material_number As String
    Dim json_string As String
    Dim json_object As Object
    Dim json_file_path As String
    Dim prop As Variant
    Dim r As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FormParser")
    ws.Visible = True
    Dim arrFiles As Variant
    arrFiles = Utils.GetFiles(pfilters:=Array("xls", "xlsx"))
    On Error Resume Next
    For r = LBound(arrFiles) To UBound(arrFiles)
        Set json_object = ParsePsf(CStr(arrFiles(r)))
        ' Clean data by removing units and filling in missing values
        For Each prop In json_object
            ' If IsNumeric(Left(json_object(prop), 1)) Then
            '     json_object.item(prop) = Utils.CleanString(json_object(prop), _
            '             Array(Chr(34),), _
            '             True)
            ' End If
            If json_object.item(prop) = Chr(34) Then json_object.item(prop) = Chr(34) + Chr(34)
        Next prop
        Cells(r, 1) = json_object(material_keyword)
        json_string = JsonVBA.ConvertToJson(json_object)
        ws.Cells(r, 2).value = json_string
    Next r
End Sub

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
    'App.Start
    file_path = SelectRBAFile
    'progress_bar = App.gDll.ShowProgressBar(4)
    ' Task 1
    'progress_bar = App.gDll.SetProgressBar(progress_bar, 1, "Task 1/4")
    path_no_ext = Replace(file_path, ".xlsx", nullstr)
    path_len = Len(path_no_ext)
    char_count = path_len
    Do Until char_buffer = "_"
        char_count = char_count - 1
        char_buffer = Mid(path_no_ext, char_count, 1)
    Loop
    material_number = Mid(path_no_ext, char_count + 1, path_len - char_count)
    ' Task 2
    'progress_bar = App.gDll.SetProgressBar(progress_bar, 2, "Task 2/4")
    Set json_object = ParseRBA(file_path)
    json_string = JsonVBA.ConvertToJson(json_object)
    ' Task 3
    'progress_bar = App.gDll.SetProgressBar(progress_bar, 3, "Task 3/4")
    Dim spec As Specification
    Set spec = CreateSpecification
    spec.JsonToObject json_string
    spec.MaterialId = material_number
    spec.SpecType = "Weaving RBA"
    spec.Revision = "1.0"
    ' Task 4
    'progress_bar = App.gDll.SetProgressBar(progress_bar, 4, "Task 4/4", AutoClose:=True)
    If SpecManager.SaveNewSpecification(spec) = DB_PUSH_SUCCESS Then
        PromptHandler.Success "New Specification Saved."
    Else
        PromptHandler.Error "Specification Not Saved."
    End If
    App.Shutdown
End Sub

Public Function SelectRBAFile() As String
' Select an RBA file from the file dialog.
    SelectRBAFile = App.gDll.OpenFile("Select RBA File . . .")
End Function

Public Function ParseRBA(path As String) As Object
    Dim wb As Workbook
    Dim strFile As String
    Dim rba_dict As Object
    Dim prop As Variant
    Dim nr As Name
    Dim rng As Object
    
    ' Turn on Performance Mode
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ret_val As Long
    Set wb = OpenWorkbook(path)
    DeleteNames wb
    Set rba_dict = CreateObject("Scripting.Dictionary")
    Set rba_dict = AddRbaNames(rba_dict, wb, "fd", 73, 82, 2, 11)
    Set rba_dict = AddRbaNames(rba_dict, wb, "di", 73, 82, 15, 24)
    Set rba_dict = AddRbaNames(rba_dict, wb, "ld", 73, 82, 28, 37)
    Set rba_dict = AddMoreRbaNames(rba_dict, wb)
    ' Clean data by removing units and filling in missing values
    For Each prop In rba_dict
        If IsNumeric(Left(rba_dict(prop), 1)) Then
            rba_dict.item(prop) = Utils.CleanString(rba_dict(prop), _
                    Array("mm", "cm", "CM", "IN", "inches", "in", "inch", "ppi", "cN/filo", "RPM", "yards", "yds", "YARDS", "rpm", "cn", "perdent"), _
                    True)
        End If
    Next prop
    ret_val = JsonVBA.WriteJsonObject(path & ".json", rba_dict)
    Set ParseRBA = rba_dict
    Set rba_dict = Nothing
    wb.Close
    ' Turn off Performance Mode
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function

Public Function ParsePsf(path As String) As Object
    Dim wb As Workbook
    Dim strFile As String
    Dim psf_dict As Object
    Dim prop As Variant
    Dim nr As Name
    Dim rng As Object

    ' Turn on Performance Mode
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ret_val As Long
    Set wb = OpenWorkbook(path)
    DeleteNames wb
    Set psf_dict = CreateObject("Scripting.Dictionary")
    Set psf_dict = CreatePsfNames(wb)
    Set ParsePsf = psf_dict
    Set psf_dict = Nothing
    wb.Close

    ' Turn off Performance Mode
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Function

Public Function CreatePsfNames(wb As Workbook) As Object
    Dim ws As Worksheet
    Set ws = wb.Sheets("Spec Sheet")
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    With wb.Names
        .Add Name:="bsf_max", RefersTo:=ws.Range("F24") ' Breaking Strength Weft Max
        .Add Name:="bsf_min", RefersTo:=ws.Range("D24") ' Breaking Strength Weft Min
        .Add Name:="bsf_nom", RefersTo:=ws.Range("E24") ' Breaking Strength Weft Nominal
        .Add Name:="bsw_max", RefersTo:=ws.Range("F23") ' Breaking Strength Warp Max
        .Add Name:="bsw_min", RefersTo:=ws.Range("D23") ' Breaking Strength Warp Minimum
        .Add Name:="bsw_nom", RefersTo:=ws.Range("E23") ' Breaking Strength Warp Nominal
        .Add Name:="cad_max", RefersTo:=ws.Range("F15") ' Conditioned Weight Max
        .Add Name:="cad_min", RefersTo:=ws.Range("D15") ' Conditioned Weight Minimum
        .Add Name:="cad_nom", RefersTo:=ws.Range("E15") ' Conditioned Weight Nominal
        .Add Name:="coated_ad_max", RefersTo:=ws.Range("F39") ' Coated Fabric Weight Max
        .Add Name:="coated_ad_min", RefersTo:=ws.Range("D39") ' Coated Fabric Weight Minimum
        .Add Name:="coated_ad_nom", RefersTo:=ws.Range("E39") ' Coated Fabric Weight Nominal
        .Add Name:="coating_description", RefersTo:=ws.Range("D38") ' Coating Description
        .Add Name:="code", RefersTo:=ws.Range("B6") ' SAP Code
        .Add Name:="core_id", RefersTo:=ws.Range("D33") ' Core Inside Diameter
        .Add Name:="core_len", RefersTo:=ws.Range("D34") ' Core Length
        .Add Name:="core_material", RefersTo:=ws.Range("G34") ' Core Length
        .Add Name:="da_max", RefersTo:=ws.Range("F27") ' Dynamic Water Absorption Max
        .Add Name:="da_min", RefersTo:=ws.Range("D27") ' Dynamic Water Absorption Minimum
        .Add Name:="da_nom", RefersTo:=ws.Range("E27") ' Dynamic Water Absorption Nominal
        .Add Name:="dad_max", RefersTo:=ws.Range("F16") ' Dry Weight Max
        .Add Name:="dad_min", RefersTo:=ws.Range("D16") ' Dry Weight Minimum
        .Add Name:="dad_nom", RefersTo:=ws.Range("E16") ' Dry Weight Nominal
        .Add Name:="description", RefersTo:=ws.Range("B5") ' Description
        .Add Name:="fiber", RefersTo:=ws.Range("D9") ' Fiber Description
        .Add Name:="fil_nom", RefersTo:=ws.Range("E14") ' Weft Thread Count Max
        .Add Name:="fill_max", RefersTo:=ws.Range("F14") ' Weft Thread Count Minimum
        .Add Name:="fill_min", RefersTo:=ws.Range("D14") ' Weft Thread Count Nominal
        .Add Name:="finish", RefersTo:=ws.Range("D25") ' Finishing Treatment
        .Add Name:="flag_defects", RefersTo:=ws.Range("D31") ' Flag Defects
        .Add Name:="fringe_max", RefersTo:=ws.Range("F21") ' Fringe Length Max
        .Add Name:="fringe_min", RefersTo:=ws.Range("D21") ' Fringe Length Minimum
        .Add Name:="fringe_nom", RefersTo:=ws.Range("E21") ' Fringe Length Nominal
        .Add Name:="insp_docs", RefersTo:=ws.Range("D32") ' Inspection Documents
        .Add Name:="lam_or_coat", RefersTo:=ws.Range("D37") ' Laminated or Coated Fabric
        .Add Name:="len_max", RefersTo:=ws.Range("F19") ' Roll Length Max
        .Add Name:="len_min", RefersTo:=ws.Range("D19") ' Roll Length Minimum
        .Add Name:="len_nom", RefersTo:=ws.Range("E19") ' Roll Length Nominal
        .Add Name:="min_acceptable_len", RefersTo:=ws.Range("D20") ' Min. Accept. Roll Length
        .Add Name:="packaging_reqs", RefersTo:=ws.Range("C36") ' Packaging
        .Add Name:="pallet_size", RefersTo:=ws.Range("D35") ' Pallet Size
        .Add Name:="product", RefersTo:=ws.Range("B4") ' Product
        .Add Name:="psf_no", RefersTo:=ws.Range("B3") ' PSF No
        .Add Name:="rc_max", RefersTo:=ws.Range("F40") ' Resin Content Max
        .Add Name:="rc_min", RefersTo:=ws.Range("D40") ' Resin Content Minimum
        .Add Name:="rc_nom", RefersTo:=ws.Range("E40") ' Resin Content Nominal
        .Add Name:="rev_no", RefersTo:=ws.Range("F3") ' Revision
        .Add Name:="shrink", RefersTo:=ws.Range("D26") ' Residual Shrinkage
        .Add Name:="sox_max", RefersTo:=ws.Range("F28") ' Extractable Level Max
        .Add Name:="sox_min", RefersTo:=ws.Range("D28") ' Extractable Level Minimum
        .Add Name:="sox_nom", RefersTo:=ws.Range("E28") ' Extractable Level Nominal
        .Add Name:="storage_reqs", RefersTo:=ws.Range("B43") ' Storage Conditions
        .Add Name:="tape_selvage", RefersTo:=ws.Range("D18") ' Tape Selvage
        .Add Name:="tg_max", RefersTo:=ws.Range("F42") ' Glass Transition Temperature Max
        .Add Name:="tg_min", RefersTo:=ws.Range("D42") ' Glass Transition Temperature Minimum
        .Add Name:="tg_nom", RefersTo:=ws.Range("E42") ' Glass Transition Temperature Nominal
        .Add Name:="thick_max", RefersTo:=ws.Range("F22") ' Final Fabric Thickness Max
        .Add Name:="thick_min", RefersTo:=ws.Range("D22") ' Final Fabric Thickness Minimum
        .Add Name:="thick_nom", RefersTo:=ws.Range("E22") ' Final Fabric Thickness Nominal
        .Add Name:="v50_layers", RefersTo:=ws.Range("D30") ' Number of Layer
        .Add Name:="v50_min", RefersTo:=ws.Range("D29") ' Minimum Average V50
        .Add Name:="v50_test_reqs", RefersTo:=ws.Range("E29") ' Minimum Average V50
        .Add Name:="vc_max", RefersTo:=ws.Range("F41") ' Volatile Content Max
        .Add Name:="vc_min", RefersTo:=ws.Range("D41") ' Volatile Content Minimum
        .Add Name:="vc_nom", RefersTo:=ws.Range("E41") ' Volatile Content Nominal
        .Add Name:="warp_max", RefersTo:=ws.Range("F13") ' Warp Thread Count Max
        .Add Name:="warp_min", RefersTo:=ws.Range("D13") ' Warp Thread Count Minimum
        .Add Name:="warp_nom", RefersTo:=ws.Range("E13") ' Warp Thread Count Nominal
        .Add Name:="warp_yarn", RefersTo:=ws.Range("D10") ' Warp Yarn
        .Add Name:="weave", RefersTo:=ws.Range("D12") ' Weave Style
        .Add Name:="weft_yarn", RefersTo:=ws.Range("D11") ' Weft Yarn
        .Add Name:="wid_max", RefersTo:=ws.Range("F17") ' Final Product Useful Width Max
        .Add Name:="wid_min", RefersTo:=ws.Range("D17") ' Final Product Useful Width Minimum
        .Add Name:="wid_nom", RefersTo:=ws.Range("E17") ' Final Product Useful Width Nominal
    End With
    With dict
        .Add Key:="bsf_max", item:=IIf(Range(wb.Names("bsf_max")).value = nullstr, nullstr, Range(wb.Names("bsf_max")).value)
        .Add Key:="bsf_min", item:=IIf(Range(wb.Names("bsf_min")).value = nullstr, nullstr, Range(wb.Names("bsf_min")).value)
        .Add Key:="bsf_nom", item:=IIf(Range(wb.Names("bsf_nom")).value = nullstr, nullstr, Range(wb.Names("bsf_nom")).value)
        .Add Key:="bsw_max", item:=IIf(Range(wb.Names("bsw_max")).value = nullstr, nullstr, Range(wb.Names("bsw_max")).value)
        .Add Key:="bsw_min", item:=IIf(Range(wb.Names("bsw_min")).value = nullstr, nullstr, Range(wb.Names("bsw_min")).value)
        .Add Key:="bsw_nom", item:=IIf(Range(wb.Names("bsw_nom")).value = nullstr, nullstr, Range(wb.Names("bsw_nom")).value)
        .Add Key:="cad_max", item:=IIf(Range(wb.Names("cad_max")).value = nullstr, nullstr, Range(wb.Names("cad_max")).value)
        .Add Key:="cad_min", item:=IIf(Range(wb.Names("cad_min")).value = nullstr, nullstr, Range(wb.Names("cad_min")).value)
        .Add Key:="cad_nom", item:=IIf(Range(wb.Names("cad_nom")).value = nullstr, nullstr, Range(wb.Names("cad_nom")).value)
        .Add Key:="coated_ad_max", item:=IIf(Range(wb.Names("coated_ad_max")).value = nullstr, nullstr, Range(wb.Names("coated_ad_max")).value)
        .Add Key:="coated_ad_min", item:=IIf(Range(wb.Names("coated_ad_min")).value = nullstr, nullstr, Range(wb.Names("coated_ad_min")).value)
        .Add Key:="coated_ad_nom", item:=IIf(Range(wb.Names("coated_ad_nom")).value = nullstr, nullstr, Range(wb.Names("coated_ad_nom")).value)
        .Add Key:="coating_description", item:=IIf(Range(wb.Names("coating_description")).value = nullstr, nullstr, Range(wb.Names("coating_description")).value)
        .Add Key:="code", item:=IIf(Range(wb.Names("code")).value = nullstr, nullstr, Range(wb.Names("code")).value)
        .Add Key:="core_id", item:=IIf(Range(wb.Names("core_id")).value = nullstr, nullstr, Range(wb.Names("core_id")).value)
        .Add Key:="core_len", item:=IIf(Range(wb.Names("core_len")).value = nullstr, nullstr, Range(wb.Names("core_len")).value)
        .Add Key:="core_material", item:=IIf(Range(wb.Names("core_material")).value = nullstr, nullstr, Range(wb.Names("core_material")).value)
        .Add Key:="da_max", item:=IIf(Range(wb.Names("da_max")).value = nullstr, nullstr, Range(wb.Names("da_max")).value)
        .Add Key:="da_min", item:=IIf(Range(wb.Names("da_min")).value = nullstr, nullstr, Range(wb.Names("da_min")).value)
        .Add Key:="da_nom", item:=IIf(Range(wb.Names("da_nom")).value = nullstr, nullstr, Range(wb.Names("da_nom")).value)
        .Add Key:="dad_max", item:=IIf(Range(wb.Names("dad_max")).value = nullstr, nullstr, Range(wb.Names("dad_max")).value)
        .Add Key:="dad_min", item:=IIf(Range(wb.Names("dad_min")).value = nullstr, nullstr, Range(wb.Names("dad_min")).value)
        .Add Key:="dad_nom", item:=IIf(Range(wb.Names("dad_nom")).value = nullstr, nullstr, Range(wb.Names("dad_nom")).value)
        .Add Key:="description", item:=IIf(Range(wb.Names("description")).value = nullstr, nullstr, Range(wb.Names("description")).value)
        .Add Key:="fiber", item:=IIf(Range(wb.Names("fiber")).value = nullstr, nullstr, Range(wb.Names("fiber")).value)
        .Add Key:="fil_nom", item:=IIf(Range(wb.Names("fil_nom")).value = nullstr, nullstr, Range(wb.Names("fil_nom")).value)
        .Add Key:="fill_max", item:=IIf(Range(wb.Names("fill_max")).value = nullstr, nullstr, Range(wb.Names("fill_max")).value)
        .Add Key:="fill_min", item:=IIf(Range(wb.Names("fill_min")).value = nullstr, nullstr, Range(wb.Names("fill_min")).value)
        .Add Key:="finish", item:=IIf(Range(wb.Names("finish")).value = nullstr, nullstr, Range(wb.Names("finish")).value)
        .Add Key:="flag_defects", item:=IIf(Range(wb.Names("flag_defects")).value = nullstr, nullstr, Range(wb.Names("flag_defects")).value)
        .Add Key:="fringe_max", item:=IIf(Range(wb.Names("fringe_max")).value = nullstr, nullstr, Range(wb.Names("fringe_max")).value)
        .Add Key:="fringe_min", item:=IIf(Range(wb.Names("fringe_min")).value = nullstr, nullstr, Range(wb.Names("fringe_min")).value)
        .Add Key:="fringe_nom", item:=IIf(Range(wb.Names("fringe_nom")).value = nullstr, nullstr, Range(wb.Names("fringe_nom")).value)
        .Add Key:="insp_docs", item:=IIf(Range(wb.Names("insp_docs")).value = nullstr, nullstr, Range(wb.Names("insp_docs")).value)
        .Add Key:="lam_or_coat", item:=IIf(Range(wb.Names("lam_or_coat")).value = nullstr, nullstr, Range(wb.Names("lam_or_coat")).value)
        .Add Key:="len_max", item:=IIf(Range(wb.Names("len_max")).value = nullstr, nullstr, Range(wb.Names("len_max")).value)
        .Add Key:="len_min", item:=IIf(Range(wb.Names("len_min")).value = nullstr, nullstr, Range(wb.Names("len_min")).value)
        .Add Key:="len_nom", item:=IIf(Range(wb.Names("len_nom")).value = nullstr, nullstr, Range(wb.Names("len_nom")).value)
        .Add Key:="min_acceptable_len", item:=IIf(Range(wb.Names("min_acceptable_len")).value = nullstr, nullstr, Range(wb.Names("min_acceptable_len")).value)
        .Add Key:="packaging_reqs", item:=IIf(Range(wb.Names("packaging_reqs")).value = nullstr, nullstr, Range(wb.Names("packaging_reqs")).value)
        .Add Key:="pallet_size", item:=IIf(Range(wb.Names("pallet_size")).value = nullstr, nullstr, Range(wb.Names("pallet_size")).value)
        .Add Key:="product", item:=IIf(Range(wb.Names("product")).value = nullstr, nullstr, Range(wb.Names("product")).value)
        .Add Key:="psf_no", item:=IIf(Range(wb.Names("psf_no")).value = nullstr, nullstr, Range(wb.Names("psf_no")).value)
        .Add Key:="rc_max", item:=IIf(Range(wb.Names("rc_max")).value = nullstr, nullstr, Range(wb.Names("rc_max")).value)
        .Add Key:="rc_min", item:=IIf(Range(wb.Names("rc_min")).value = nullstr, nullstr, Range(wb.Names("rc_min")).value)
        .Add Key:="rc_nom", item:=IIf(Range(wb.Names("rc_nom")).value = nullstr, nullstr, Range(wb.Names("rc_nom")).value)
        .Add Key:="rev_no", item:=IIf(Range(wb.Names("rev_no")).value = nullstr, nullstr, Range(wb.Names("rev_no")).value)
        .Add Key:="shrink", item:=IIf(Range(wb.Names("shrink")).value = nullstr, nullstr, Range(wb.Names("shrink")).value)
        .Add Key:="sox_max", item:=IIf(Range(wb.Names("sox_max")).value = nullstr, nullstr, Range(wb.Names("sox_max")).value)
        .Add Key:="sox_min", item:=IIf(Range(wb.Names("sox_min")).value = nullstr, nullstr, Range(wb.Names("sox_min")).value)
        .Add Key:="sox_nom", item:=IIf(Range(wb.Names("sox_nom")).value = nullstr, nullstr, Range(wb.Names("sox_nom")).value)
        .Add Key:="storage_reqs", item:=IIf(Range(wb.Names("storage_reqs")).value = nullstr, nullstr, Range(wb.Names("storage_reqs")).value)
        .Add Key:="tape_selvage", item:=IIf(Range(wb.Names("tape_selvage")).value = nullstr, nullstr, Range(wb.Names("tape_selvage")).value)
        .Add Key:="tg_max", item:=IIf(Range(wb.Names("tg_max")).value = nullstr, nullstr, Range(wb.Names("tg_max")).value)
        .Add Key:="tg_min", item:=IIf(Range(wb.Names("tg_min")).value = nullstr, nullstr, Range(wb.Names("tg_min")).value)
        .Add Key:="tg_nom", item:=IIf(Range(wb.Names("tg_nom")).value = nullstr, nullstr, Range(wb.Names("tg_nom")).value)
        .Add Key:="thick_max", item:=IIf(Range(wb.Names("thick_max")).value = nullstr, nullstr, Range(wb.Names("thick_max")).value)
        .Add Key:="thick_min", item:=IIf(Range(wb.Names("thick_min")).value = nullstr, nullstr, Range(wb.Names("thick_min")).value)
        .Add Key:="thick_nom", item:=IIf(Range(wb.Names("thick_nom")).value = nullstr, nullstr, Range(wb.Names("thick_nom")).value)
        .Add Key:="v50_layers", item:=IIf(Range(wb.Names("v50_layers")).value = nullstr, nullstr, Range(wb.Names("v50_layers")).value)
        .Add Key:="v50_min", item:=IIf(Range(wb.Names("v50_min")).value = nullstr, nullstr, Range(wb.Names("v50_min")).value)
        .Add Key:="v50_test_reqs", item:=IIf(Range(wb.Names("v50_test_reqs")).value = nullstr, nullstr, Range(wb.Names("v50_test_reqs")).value)
        .Add Key:="vc_max", item:=IIf(Range(wb.Names("vc_max")).value = nullstr, nullstr, Range(wb.Names("vc_max")).value)
        .Add Key:="vc_min", item:=IIf(Range(wb.Names("vc_min")).value = nullstr, nullstr, Range(wb.Names("vc_min")).value)
        .Add Key:="vc_nom", item:=IIf(Range(wb.Names("vc_nom")).value = nullstr, nullstr, Range(wb.Names("vc_nom")).value)
        .Add Key:="warp_max", item:=IIf(Range(wb.Names("warp_max")).value = nullstr, nullstr, Range(wb.Names("warp_max")).value)
        .Add Key:="warp_min", item:=IIf(Range(wb.Names("warp_min")).value = nullstr, nullstr, Range(wb.Names("warp_min")).value)
        .Add Key:="warp_nom", item:=IIf(Range(wb.Names("warp_nom")).value = nullstr, nullstr, Range(wb.Names("warp_nom")).value)
        .Add Key:="warp_yarn", item:=IIf(Range(wb.Names("warp_yarn")).value = nullstr, nullstr, Range(wb.Names("warp_yarn")).value)
        .Add Key:="weave", item:=IIf(Range(wb.Names("weave")).value = nullstr, nullstr, Range(wb.Names("weave")).value)
        .Add Key:="weft_yarn", item:=IIf(Range(wb.Names("weft_yarn")).value = nullstr, nullstr, Range(wb.Names("weft_yarn")).value)
        .Add Key:="wid_max", item:=IIf(Range(wb.Names("wid_max")).value = nullstr, nullstr, Range(wb.Names("wid_max")).value)
        .Add Key:="wid_min", item:=IIf(Range(wb.Names("wid_min")).value = nullstr, nullstr, Range(wb.Names("wid_min")).value)
        .Add Key:="wid_nom", item:=IIf(Range(wb.Names("wid_nom")).value = nullstr, nullstr, Range(wb.Names("wid_nom")).value)
    End With

    Set CreatePsfNames = dict

End Function

Public Function AddMoreRbaNames(dict As Object, wb As Workbook) As Object
    
    With wb.Names
        .Add Name:="actual_weft_count", RefersTo:=wb.Sheets("ENG").Range("AC26")
        .Add Name:="article_code", RefersTo:=wb.Sheets("ENG").Range("J14")
        .Add Name:="aux_selvedges_closing_degrees", RefersTo:=wb.Sheets("ENG").Range("AC36")
        .Add Name:="bottom_rapier_clamps", RefersTo:=wb.Sheets("ENG").Range("AC49")
        .Add Name:="bottom_spreader_bars", RefersTo:=wb.Sheets("ENG").Range("AC59")
        .Add Name:="central_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T69")
        .Add Name:="central_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z69")
        .Add Name:="central_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J69")
        .Add Name:="central_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF69")
        .Add Name:="central_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N69")
        .Add Name:="cutting_degrees", RefersTo:=wb.Sheets("ENG").Range("J38")
        .Add Name:="date", RefersTo:=wb.Sheets("ENG").Range("AC8")
        .Add Name:="dorn_left_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T66")
        .Add Name:="dorn_left_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z66")
        .Add Name:="dorn_left_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J66")
        .Add Name:="dorn_left_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF66")
        .Add Name:="dorn_left_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N66")
        .Add Name:="draw_in_harness", RefersTo:=wb.Sheets("ENG").Range("AC18")
        .Add Name:="draw_in_reed", RefersTo:=wb.Sheets("ENG").Range("AC20")
        .Add Name:="fabric_width", RefersTo:=wb.Sheets("ENG").Range("J12")
        .Add Name:="first_heddle", RefersTo:=wb.Sheets("ENG").Range("J30")
        .Add Name:="first_heddle_1", RefersTo:=wb.Sheets("ENG").Range("J30")
        .Add Name:="first_heddle_guide", RefersTo:=wb.Sheets("ENG").Range("J34")
        .Add Name:="harness_configuration", RefersTo:=wb.Sheets("ENG").Range("J22")
        .Add Name:="horizontal_back_rest_roller", RefersTo:=wb.Sheets("ENG").Range("J42")
        .Add Name:="last_heddle", RefersTo:=wb.Sheets("ENG").Range("AC30")
        .Add Name:="last_heddle_guide", RefersTo:=wb.Sheets("ENG").Range("AC34")
        .Add Name:="left_main_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T67")
        .Add Name:="left_main_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z67")
        .Add Name:="left_main_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J67")
        .Add Name:="left_main_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF67")
        .Add Name:="left_main_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N67")
        .Add Name:="left_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T64")
        .Add Name:="left_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z64")
        .Add Name:="left_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J64")
        .Add Name:="left_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF64")
        .Add Name:="left_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N64")
        .Add Name:="loom_number", RefersTo:=wb.Sheets("ENG").Range("AC10")
        .Add Name:="loom_type", RefersTo:=wb.Sheets("ENG").Range("AC12")
        .Add Name:="number_ends_wo_selvedges", RefersTo:=wb.Sheets("ENG").Range("J24")
        .Add Name:="number_harnesses", RefersTo:=wb.Sheets("ENG").Range("J20")
        .Add Name:="pinch_roller_felt_type", RefersTo:=wb.Sheets("ENG").Range("J55")
        .Add Name:="press_roller_type", RefersTo:=wb.Sheets("ENG").Range("J53")
        .Add Name:="rba_number", RefersTo:=wb.Sheets("ENG").Range("J8")
        .Add Name:="reed", RefersTo:=wb.Sheets("ENG").Range("J16")
        .Add Name:="reed_width", RefersTo:=wb.Sheets("ENG").Range("AC16")
        .Add Name:="right_main_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T68")
        .Add Name:="right_main_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z68")
        .Add Name:="right_main_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J68")
        .Add Name:="right_main_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF68")
        .Add Name:="right_main_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N68")
        .Add Name:="right_selvedges_drawing_in", RefersTo:=wb.Sheets("ENG").Range("T65")
        .Add Name:="right_selvedges_ends_per_dent", RefersTo:=wb.Sheets("ENG").Range("Z65")
        .Add Name:="right_selvedges_number_ends", RefersTo:=wb.Sheets("ENG").Range("J65")
        .Add Name:="right_selvedges_weave", RefersTo:=wb.Sheets("ENG").Range("AF65")
        .Add Name:="right_selvedges_yarn_count", RefersTo:=wb.Sheets("ENG").Range("N65")
        .Add Name:="sand_roller_type", RefersTo:=wb.Sheets("ENG").Range("AC53")
        .Add Name:="selvedges_type", RefersTo:=wb.Sheets("ENG").Range("AC22")
        .Add Name:="shed_closing_degrees", RefersTo:=wb.Sheets("ENG").Range("J36")
        .Add Name:="speed", RefersTo:=wb.Sheets("ENG").Range("AC14")
        .Add Name:="springs_type", RefersTo:=wb.Sheets("ENG").Range("J44")
        .Add Name:="style_number", RefersTo:=wb.Sheets("ENG").Range("J10")
        .Add Name:="temples_composition", RefersTo:=wb.Sheets("ENG").Range("AC44")
        .Add Name:="upper_rapier_clamps", RefersTo:=wb.Sheets("ENG").Range("J49")
        .Add Name:="upper_spreader_bars", RefersTo:=wb.Sheets("ENG").Range("J59")
        .Add Name:="vertical_back_rest_roller", RefersTo:=wb.Sheets("ENG").Range("AC42")
        .Add Name:="warp_tension", RefersTo:=wb.Sheets("ENG").Range("J26")
        .Add Name:="weave_pattern", RefersTo:=wb.Sheets("ENG").Range("J18")
        .Add Name:="weft_count_set_point", RefersTo:=wb.Sheets("ENG").Range("AC24")
        .Add Name:="notes1", RefersTo:=wb.Sheets("ENG").Range("H86")
        .Add Name:="notes2", RefersTo:=wb.Sheets("ENG").Range("H87")
        .Add Name:="notes3", RefersTo:=wb.Sheets("ENG").Range("H88")
        .Add Name:="notes4", RefersTo:=wb.Sheets("ENG").Range("H89")
        .Add Name:="notes5", RefersTo:=wb.Sheets("ENG").Range("H90")
        .Add Name:="notes6", RefersTo:=wb.Sheets("ENG").Range("H91")
        .Add Name:="notes7", RefersTo:=wb.Sheets("ENG").Range("H92")
        .Add Name:="notes8", RefersTo:=wb.Sheets("ENG").Range("H93")
        .Add Name:="roll_length", RefersTo:=wb.Sheets("ENG").Range("AC84")
    End With
    With dict
        .Add Key:="actual_weft_count", item:=IIf(Range(wb.Names("actual_weft_count")).value = nullstr, nullstr, Range(wb.Names("actual_weft_count")).value)
        .Add Key:="article_code", item:=IIf(Range(wb.Names("article_code")).value = nullstr, nullstr, Range(wb.Names("article_code")).value)
        .Add Key:="aux_selvedges_closing_degrees", item:=IIf(Range(wb.Names("aux_selvedges_closing_degrees")).value = nullstr, nullstr, Range(wb.Names("aux_selvedges_closing_degrees")).value)
        .Add Key:="bottom_rapier_clamps", item:=IIf(Range(wb.Names("bottom_rapier_clamps")).value = nullstr, nullstr, Range(wb.Names("bottom_rapier_clamps")).value)
        .Add Key:="bottom_spreader_bars", item:=IIf(Range(wb.Names("bottom_spreader_bars")).value = nullstr, nullstr, Range(wb.Names("bottom_spreader_bars")).value)
        .Add Key:="central_selvedges_drawing_in", item:=IIf(Range(wb.Names("central_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("central_selvedges_drawing_in")).value)
        .Add Key:="central_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("central_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("central_selvedges_ends_per_dent")).value)
        .Add Key:="central_selvedges_number_ends", item:=IIf(Range(wb.Names("central_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("central_selvedges_number_ends")).value)
        .Add Key:="central_selvedges_weave", item:=IIf(Range(wb.Names("central_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("central_selvedges_weave")).value)
        .Add Key:="central_selvedges_yarn_count", item:=IIf(Range(wb.Names("central_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("central_selvedges_yarn_count")).value)
        .Add Key:="cutting_degrees", item:=IIf(Range(wb.Names("cutting_degrees")).value = nullstr, nullstr, Range(wb.Names("cutting_degrees")).value)
        .Add Key:="date", item:=IIf(Range(wb.Names("date")).value = nullstr, nullstr, Range(wb.Names("date")).value)
        .Add Key:="dorn_left_selvedges_drawing_in", item:=IIf(Range(wb.Names("dorn_left_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("dorn_left_selvedges_drawing_in")).value)
        .Add Key:="dorn_left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value)
        .Add Key:="dorn_left_selvedges_number_ends", item:=IIf(Range(wb.Names("dorn_left_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("dorn_left_selvedges_number_ends")).value)
        .Add Key:="dorn_left_selvedges_weave", item:=IIf(Range(wb.Names("dorn_left_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("dorn_left_selvedges_weave")).value)
        .Add Key:="dorn_left_selvedges_yarn_count", item:=IIf(Range(wb.Names("dorn_left_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("dorn_left_selvedges_yarn_count")).value)
        .Add Key:="draw_in_harness", item:=IIf(Range(wb.Names("draw_in_harness")).value = nullstr, nullstr, Range(wb.Names("draw_in_harness")).value)
        .Add Key:="draw_in_reed", item:=IIf(Range(wb.Names("draw_in_reed")).value = nullstr, nullstr, Range(wb.Names("draw_in_reed")).value)
        .Add Key:="fabric_width", item:=IIf(Range(wb.Names("fabric_width")).value = nullstr, nullstr, Range(wb.Names("fabric_width")).value)
        .Add Key:="first_heddle", item:=IIf(Range(wb.Names("first_heddle")).value = nullstr, nullstr, Range(wb.Names("first_heddle")).value)
        .Add Key:="first_heddle_1", item:=IIf(Range(wb.Names("first_heddle_1")).value = nullstr, nullstr, Range(wb.Names("first_heddle_1")).value)
        .Add Key:="first_heddle_guide", item:=IIf(Range(wb.Names("first_heddle_guide")).value = nullstr, nullstr, Range(wb.Names("first_heddle_guide")).value)
        .Add Key:="harness_configuration", item:=IIf(Range(wb.Names("harness_configuration")).value = nullstr, nullstr, Range(wb.Names("harness_configuration")).value)
        .Add Key:="horizontal_back_rest_roller", item:=IIf(Range(wb.Names("horizontal_back_rest_roller")).value = nullstr, nullstr, Range(wb.Names("horizontal_back_rest_roller")).value)
        .Add Key:="last_heddle", item:=IIf(Range(wb.Names("last_heddle")).value = nullstr, nullstr, Range(wb.Names("last_heddle")).value)
        .Add Key:="last_heddle_guide", item:=IIf(Range(wb.Names("last_heddle_guide")).value = nullstr, nullstr, Range(wb.Names("last_heddle_guide")).value)
        .Add Key:="left_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_main_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("left_main_selvedges_drawing_in")).value)
        .Add Key:="left_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_main_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("left_main_selvedges_ends_per_dent")).value)
        .Add Key:="left_main_selvedges_number_ends", item:=IIf(Range(wb.Names("left_main_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("left_main_selvedges_number_ends")).value)
        .Add Key:="left_main_selvedges_weave", item:=IIf(Range(wb.Names("left_main_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("left_main_selvedges_weave")).value)
        .Add Key:="left_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_main_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("left_main_selvedges_yarn_count")).value)
        .Add Key:="left_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("left_selvedges_drawing_in")).value)
        .Add Key:="left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("left_selvedges_ends_per_dent")).value)
        .Add Key:="left_selvedges_number_ends", item:=IIf(Range(wb.Names("left_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("left_selvedges_number_ends")).value)
        .Add Key:="left_selvedges_weave", item:=IIf(Range(wb.Names("left_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("left_selvedges_weave")).value)
        .Add Key:="left_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("left_selvedges_yarn_count")).value)
        .Add Key:="loom_number", item:=IIf(Range(wb.Names("loom_number")).value = nullstr, nullstr, Range(wb.Names("loom_number")).value)
        .Add Key:="loom_type", item:=IIf(Range(wb.Names("loom_type")).value = nullstr, nullstr, Range(wb.Names("loom_type")).value)
        .Add Key:="number_ends_wo_selvedges", item:=IIf(Range(wb.Names("number_ends_wo_selvedges")).value = nullstr, nullstr, Range(wb.Names("number_ends_wo_selvedges")).value)
        .Add Key:="number_harnesses", item:=IIf(Range(wb.Names("number_harnesses")).value = nullstr, nullstr, Range(wb.Names("number_harnesses")).value)
        .Add Key:="pinch_roller_felt_type", item:=IIf(Range(wb.Names("pinch_roller_felt_type")).value = nullstr, nullstr, Range(wb.Names("pinch_roller_felt_type")).value)
        .Add Key:="press_roller_type", item:=IIf(Range(wb.Names("press_roller_type")).value = nullstr, nullstr, Range(wb.Names("press_roller_type")).value)
        .Add Key:="rba_number", item:=IIf(Range(wb.Names("rba_number")).value = nullstr, nullstr, Range(wb.Names("rba_number")).value)
        .Add Key:="reed", item:=IIf(Range(wb.Names("reed")).value = nullstr, nullstr, Range(wb.Names("reed")).value)
        .Add Key:="reed_width", item:=IIf(Range(wb.Names("reed_width")).value = nullstr, nullstr, Range(wb.Names("reed_width")).value)
        .Add Key:="right_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_main_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("right_main_selvedges_drawing_in")).value)
        .Add Key:="right_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_main_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("right_main_selvedges_ends_per_dent")).value)
        .Add Key:="right_main_selvedges_number_ends", item:=IIf(Range(wb.Names("right_main_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("right_main_selvedges_number_ends")).value)
        .Add Key:="right_main_selvedges_weave", item:=IIf(Range(wb.Names("right_main_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("right_main_selvedges_weave")).value)
        .Add Key:="right_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_main_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("right_main_selvedges_yarn_count")).value)
        .Add Key:="right_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_selvedges_drawing_in")).value = nullstr, nullstr, Range(wb.Names("right_selvedges_drawing_in")).value)
        .Add Key:="right_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_selvedges_ends_per_dent")).value = nullstr, nullstr, Range(wb.Names("right_selvedges_ends_per_dent")).value)
        .Add Key:="right_selvedges_number_ends", item:=IIf(Range(wb.Names("right_selvedges_number_ends")).value = nullstr, nullstr, Range(wb.Names("right_selvedges_number_ends")).value)
        .Add Key:="right_selvedges_weave", item:=IIf(Range(wb.Names("right_selvedges_weave")).value = nullstr, nullstr, Range(wb.Names("right_selvedges_weave")).value)
        .Add Key:="right_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_selvedges_yarn_count")).value = nullstr, nullstr, Range(wb.Names("right_selvedges_yarn_count")).value)
        .Add Key:="sand_roller_type", item:=IIf(Range(wb.Names("sand_roller_type")).value = nullstr, nullstr, Range(wb.Names("sand_roller_type")).value)
        .Add Key:="selvedges_type", item:=IIf(Range(wb.Names("selvedges_type")).value = nullstr, nullstr, Range(wb.Names("selvedges_type")).value)
        .Add Key:="shed_closing_degrees", item:=IIf(Range(wb.Names("shed_closing_degrees")).value = nullstr, nullstr, Range(wb.Names("shed_closing_degrees")).value)
        .Add Key:="speed", item:=IIf(Range(wb.Names("speed")).value = nullstr, nullstr, Range(wb.Names("speed")).value)
        .Add Key:="springs_type", item:=IIf(Range(wb.Names("springs_type")).value = nullstr, nullstr, Range(wb.Names("springs_type")).value)
        .Add Key:="style_number", item:=IIf(Range(wb.Names("style_number")).value = nullstr, nullstr, Range(wb.Names("style_number")).value)
        .Add Key:="temples_composition", item:=IIf(Range(wb.Names("temples_composition")).value = nullstr, nullstr, Range(wb.Names("temples_composition")).value)
        .Add Key:="upper_rapier_clamps", item:=IIf(Range(wb.Names("upper_rapier_clamps")).value = nullstr, nullstr, Range(wb.Names("upper_rapier_clamps")).value)
        .Add Key:="upper_spreader_bars", item:=IIf(Range(wb.Names("upper_spreader_bars")).value = nullstr, nullstr, Range(wb.Names("upper_spreader_bars")).value)
        .Add Key:="vertical_back_rest_roller", item:=IIf(Range(wb.Names("vertical_back_rest_roller")).value = nullstr, nullstr, Range(wb.Names("vertical_back_rest_roller")).value)
        .Add Key:="warp_tension", item:=IIf(Range(wb.Names("warp_tension")).value = nullstr, nullstr, Range(wb.Names("warp_tension")).value)
        .Add Key:="weave_pattern", item:=IIf(Range(wb.Names("weave_pattern")).value = nullstr, nullstr, Range(wb.Names("weave_pattern")).value)
        .Add Key:="weft_count_set_point", item:=IIf(Range(wb.Names("weft_count_set_point")).value = nullstr, nullstr, Range(wb.Names("weft_count_set_point")).value)
        .Add Key:="notes1", item:=IIf(Range(wb.Names("notes1")).value = nullstr, nullstr, Range(wb.Names("notes1")).value)
        .Add Key:="notes2", item:=IIf(Range(wb.Names("notes2")).value = nullstr, nullstr, Range(wb.Names("notes2")).value)
        .Add Key:="notes3", item:=IIf(Range(wb.Names("notes3")).value = nullstr, nullstr, Range(wb.Names("notes3")).value)
        .Add Key:="notes4", item:=IIf(Range(wb.Names("notes4")).value = nullstr, nullstr, Range(wb.Names("notes4")).value)
        .Add Key:="notes5", item:=IIf(Range(wb.Names("notes5")).value = nullstr, nullstr, Range(wb.Names("notes5")).value)
        .Add Key:="notes6", item:=IIf(Range(wb.Names("notes6")).value = nullstr, nullstr, Range(wb.Names("notes6")).value)
        .Add Key:="notes7", item:=IIf(Range(wb.Names("notes7")).value = nullstr, nullstr, Range(wb.Names("notes7")).value)
        .Add Key:="notes8", item:=IIf(Range(wb.Names("notes8")).value = nullstr, nullstr, Range(wb.Names("notes8")).value)
        .Add Key:="roll_length", item:=IIf(Range(wb.Names("roll_length")).value = nullstr, nullstr, Range(wb.Names("roll_length")).value)
    End With
    Set AddMoreRbaNames = dict
End Function

Public Function AddRbaNames(dict As Object, wb As Workbook, tag As String, r_start As Long, r_end As Long, c_start As Long, c_end As Long) As Object
    Dim sht As Worksheet
    Dim nr As String
    Dim ret_val As Variant
    Set sht = wb.Sheets("ENG")
    Dim r, C, rw, cl As Long
    For r = r_start To r_end
        cl = 0
        For C = c_start To c_end
            rw = Abs(r_end - r)
            nr = tag & "_" & cl & rw
            ret_val = CreateNamedRange(wb, nr, sht, CLng(r), CLng(C))
            dict.Add nr, IIf(ret_val = nullstr, nullstr, ret_val)
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
    Set ws = shtRbaParser
    
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
        dict.Add Key:=Name, item:=IIf(Range(wb.Names(Name)).value = nullstr, _
                                      nullstr, Range(wb.Names(Name)).value)
    Next i
    Set ApplyNames = dict
End Function
