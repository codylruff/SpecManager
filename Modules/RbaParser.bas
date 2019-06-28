Attribute VB_Name = "RbaParser"
Option Explicit
' ==============================================
' RBA PARSER
' ==============================================
Public Sub ParseAll()
    Dim material_number As String
    Dim json_string As String
    Dim json_object As Object
    Dim json_file_path As String
    Dim r As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Set ws = shtRbaParser
    
    For r = 1 To 16
        material_number = Cells(r, 1)
        Set json_object = ParseRBA(ThisWorkbook.path & "\RBAs\" & material_number & ".xlsx")
        json_object("article_code") = material_number
        json_string = JsonVBA.ConvertToJson(json_object)
        ws.Cells(r, 2).Value = json_string
    Next r
End Sub

Public Function ParseRBA(path As String) As Object
    Dim wb As Workbook
    Dim strFile As String
    Dim rba_dict As Object
    Dim prop As Variant
    Dim nr As Name
    Dim rng As Object
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
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
            rba_dict.Item(prop) = Utils.CleanString(rba_dict(prop), _
                    Array("mm", "cm", "CM", "IN", "inches", "in", "inch", "ppi", "cN/filo", "RPM", "yards", "yds", "YARDS", "rpm", "cn", "perdent"), _
                    True)
        End If
    Next prop
    ret_val = JsonVBA.WriteJsonObject(path & ".json", rba_dict)
    Set ParseRBA = rba_dict
    Set rba_dict = Nothing
    wb.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
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
        .Add Key:="actual_weft_count", Item:=IIf(Range(wb.Names("actual_weft_count")).Value = vbNullString, vbNullString, Range(wb.Names("actual_weft_count")).Value)
        .Add Key:="article_code", Item:=IIf(Range(wb.Names("article_code")).Value = vbNullString, vbNullString, Range(wb.Names("article_code")).Value)
        .Add Key:="aux_selvedges_closing_degrees", Item:=IIf(Range(wb.Names("aux_selvedges_closing_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("aux_selvedges_closing_degrees")).Value)
        .Add Key:="bottom_rapier_clamps", Item:=IIf(Range(wb.Names("bottom_rapier_clamps")).Value = vbNullString, vbNullString, Range(wb.Names("bottom_rapier_clamps")).Value)
        .Add Key:="bottom_spreader_bars", Item:=IIf(Range(wb.Names("bottom_spreader_bars")).Value = vbNullString, vbNullString, Range(wb.Names("bottom_spreader_bars")).Value)
        .Add Key:="central_selvedges_drawing_in", Item:=IIf(Range(wb.Names("central_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_drawing_in")).Value)
        .Add Key:="central_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("central_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_ends_per_dent")).Value)
        .Add Key:="central_selvedges_number_ends", Item:=IIf(Range(wb.Names("central_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_number_ends")).Value)
        .Add Key:="central_selvedges_weave", Item:=IIf(Range(wb.Names("central_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_weave")).Value)
        .Add Key:="central_selvedges_yarn_count", Item:=IIf(Range(wb.Names("central_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_yarn_count")).Value)
        .Add Key:="cutting_degrees", Item:=IIf(Range(wb.Names("cutting_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("cutting_degrees")).Value)
        .Add Key:="date", Item:=IIf(Range(wb.Names("date")).Value = vbNullString, vbNullString, Range(wb.Names("date")).Value)
        .Add Key:="dorn_left_selvedges_drawing_in", Item:=IIf(Range(wb.Names("dorn_left_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_drawing_in")).Value)
        .Add Key:="dorn_left_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("dorn_left_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_ends_per_dent")).Value)
        .Add Key:="dorn_left_selvedges_number_ends", Item:=IIf(Range(wb.Names("dorn_left_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_number_ends")).Value)
        .Add Key:="dorn_left_selvedges_weave", Item:=IIf(Range(wb.Names("dorn_left_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_weave")).Value)
        .Add Key:="dorn_left_selvedges_yarn_count", Item:=IIf(Range(wb.Names("dorn_left_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_yarn_count")).Value)
        .Add Key:="draw_in_harness", Item:=IIf(Range(wb.Names("draw_in_harness")).Value = vbNullString, vbNullString, Range(wb.Names("draw_in_harness")).Value)
        .Add Key:="draw_in_reed", Item:=IIf(Range(wb.Names("draw_in_reed")).Value = vbNullString, vbNullString, Range(wb.Names("draw_in_reed")).Value)
        .Add Key:="fabric_width", Item:=IIf(Range(wb.Names("fabric_width")).Value = vbNullString, vbNullString, Range(wb.Names("fabric_width")).Value)
        .Add Key:="first_heddle", Item:=IIf(Range(wb.Names("first_heddle")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle")).Value)
        .Add Key:="first_heddle_1", Item:=IIf(Range(wb.Names("first_heddle_1")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle_1")).Value)
        .Add Key:="first_heddle_guide", Item:=IIf(Range(wb.Names("first_heddle_guide")).Value = vbNullString, vbNullString, Range(wb.Names("first_heddle_guide")).Value)
        .Add Key:="harness_configuration", Item:=IIf(Range(wb.Names("harness_configuration")).Value = vbNullString, vbNullString, Range(wb.Names("harness_configuration")).Value)
        .Add Key:="horizontal_back_rest_roller", Item:=IIf(Range(wb.Names("horizontal_back_rest_roller")).Value = vbNullString, vbNullString, Range(wb.Names("horizontal_back_rest_roller")).Value)
        .Add Key:="last_heddle", Item:=IIf(Range(wb.Names("last_heddle")).Value = vbNullString, vbNullString, Range(wb.Names("last_heddle")).Value)
        .Add Key:="last_heddle_guide", Item:=IIf(Range(wb.Names("last_heddle_guide")).Value = vbNullString, vbNullString, Range(wb.Names("last_heddle_guide")).Value)
        .Add Key:="left_main_selvedges_drawing_in", Item:=IIf(Range(wb.Names("left_main_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_drawing_in")).Value)
        .Add Key:="left_main_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("left_main_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_ends_per_dent")).Value)
        .Add Key:="left_main_selvedges_number_ends", Item:=IIf(Range(wb.Names("left_main_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_number_ends")).Value)
        .Add Key:="left_main_selvedges_weave", Item:=IIf(Range(wb.Names("left_main_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_weave")).Value)
        .Add Key:="left_main_selvedges_yarn_count", Item:=IIf(Range(wb.Names("left_main_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_yarn_count")).Value)
        .Add Key:="left_selvedges_drawing_in", Item:=IIf(Range(wb.Names("left_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_drawing_in")).Value)
        .Add Key:="left_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("left_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_ends_per_dent")).Value)
        .Add Key:="left_selvedges_number_ends", Item:=IIf(Range(wb.Names("left_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_number_ends")).Value)
        .Add Key:="left_selvedges_weave", Item:=IIf(Range(wb.Names("left_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_weave")).Value)
        .Add Key:="left_selvedges_yarn_count", Item:=IIf(Range(wb.Names("left_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_yarn_count")).Value)
        .Add Key:="loom_number", Item:=IIf(Range(wb.Names("loom_number")).Value = vbNullString, vbNullString, Range(wb.Names("loom_number")).Value)
        .Add Key:="loom_type", Item:=IIf(Range(wb.Names("loom_type")).Value = vbNullString, vbNullString, Range(wb.Names("loom_type")).Value)
        .Add Key:="number_ends_wo_selvedges", Item:=IIf(Range(wb.Names("number_ends_wo_selvedges")).Value = vbNullString, vbNullString, Range(wb.Names("number_ends_wo_selvedges")).Value)
        .Add Key:="number_harnesses", Item:=IIf(Range(wb.Names("number_harnesses")).Value = vbNullString, vbNullString, Range(wb.Names("number_harnesses")).Value)
        .Add Key:="pinch_roller_felt_type", Item:=IIf(Range(wb.Names("pinch_roller_felt_type")).Value = vbNullString, vbNullString, Range(wb.Names("pinch_roller_felt_type")).Value)
        .Add Key:="press_roller_type", Item:=IIf(Range(wb.Names("press_roller_type")).Value = vbNullString, vbNullString, Range(wb.Names("press_roller_type")).Value)
        .Add Key:="rba_number", Item:=IIf(Range(wb.Names("rba_number")).Value = vbNullString, vbNullString, Range(wb.Names("rba_number")).Value)
        .Add Key:="reed", Item:=IIf(Range(wb.Names("reed")).Value = vbNullString, vbNullString, Range(wb.Names("reed")).Value)
        .Add Key:="reed_width", Item:=IIf(Range(wb.Names("reed_width")).Value = vbNullString, vbNullString, Range(wb.Names("reed_width")).Value)
        .Add Key:="right_main_selvedges_drawing_in", Item:=IIf(Range(wb.Names("right_main_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_drawing_in")).Value)
        .Add Key:="right_main_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("right_main_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_ends_per_dent")).Value)
        .Add Key:="right_main_selvedges_number_ends", Item:=IIf(Range(wb.Names("right_main_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_number_ends")).Value)
        .Add Key:="right_main_selvedges_weave", Item:=IIf(Range(wb.Names("right_main_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_weave")).Value)
        .Add Key:="right_main_selvedges_yarn_count", Item:=IIf(Range(wb.Names("right_main_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_yarn_count")).Value)
        .Add Key:="right_selvedges_drawing_in", Item:=IIf(Range(wb.Names("right_selvedges_drawing_in")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_drawing_in")).Value)
        .Add Key:="right_selvedges_ends_per_dent", Item:=IIf(Range(wb.Names("right_selvedges_ends_per_dent")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_ends_per_dent")).Value)
        .Add Key:="right_selvedges_number_ends", Item:=IIf(Range(wb.Names("right_selvedges_number_ends")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_number_ends")).Value)
        .Add Key:="right_selvedges_weave", Item:=IIf(Range(wb.Names("right_selvedges_weave")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_weave")).Value)
        .Add Key:="right_selvedges_yarn_count", Item:=IIf(Range(wb.Names("right_selvedges_yarn_count")).Value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_yarn_count")).Value)
        .Add Key:="sand_roller_type", Item:=IIf(Range(wb.Names("sand_roller_type")).Value = vbNullString, vbNullString, Range(wb.Names("sand_roller_type")).Value)
        .Add Key:="selvedges_type", Item:=IIf(Range(wb.Names("selvedges_type")).Value = vbNullString, vbNullString, Range(wb.Names("selvedges_type")).Value)
        .Add Key:="shed_closing_degrees", Item:=IIf(Range(wb.Names("shed_closing_degrees")).Value = vbNullString, vbNullString, Range(wb.Names("shed_closing_degrees")).Value)
        .Add Key:="speed", Item:=IIf(Range(wb.Names("speed")).Value = vbNullString, vbNullString, Range(wb.Names("speed")).Value)
        .Add Key:="springs_type", Item:=IIf(Range(wb.Names("springs_type")).Value = vbNullString, vbNullString, Range(wb.Names("springs_type")).Value)
        .Add Key:="style_number", Item:=IIf(Range(wb.Names("style_number")).Value = vbNullString, vbNullString, Range(wb.Names("style_number")).Value)
        .Add Key:="temples_composition", Item:=IIf(Range(wb.Names("temples_composition")).Value = vbNullString, vbNullString, Range(wb.Names("temples_composition")).Value)
        .Add Key:="upper_rapier_clamps", Item:=IIf(Range(wb.Names("upper_rapier_clamps")).Value = vbNullString, vbNullString, Range(wb.Names("upper_rapier_clamps")).Value)
        .Add Key:="upper_spreader_bars", Item:=IIf(Range(wb.Names("upper_spreader_bars")).Value = vbNullString, vbNullString, Range(wb.Names("upper_spreader_bars")).Value)
        .Add Key:="vertical_back_rest_roller", Item:=IIf(Range(wb.Names("vertical_back_rest_roller")).Value = vbNullString, vbNullString, Range(wb.Names("vertical_back_rest_roller")).Value)
        .Add Key:="warp_tension", Item:=IIf(Range(wb.Names("warp_tension")).Value = vbNullString, vbNullString, Range(wb.Names("warp_tension")).Value)
        .Add Key:="weave_pattern", Item:=IIf(Range(wb.Names("weave_pattern")).Value = vbNullString, vbNullString, Range(wb.Names("weave_pattern")).Value)
        .Add Key:="weft_count_set_point", Item:=IIf(Range(wb.Names("weft_count_set_point")).Value = vbNullString, vbNullString, Range(wb.Names("weft_count_set_point")).Value)
        .Add Key:="notes1", Item:=IIf(Range(wb.Names("notes1")).Value = vbNullString, vbNullString, Range(wb.Names("notes1")).Value)
        .Add Key:="notes2", Item:=IIf(Range(wb.Names("notes2")).Value = vbNullString, vbNullString, Range(wb.Names("notes2")).Value)
        .Add Key:="notes3", Item:=IIf(Range(wb.Names("notes3")).Value = vbNullString, vbNullString, Range(wb.Names("notes3")).Value)
        .Add Key:="notes4", Item:=IIf(Range(wb.Names("notes4")).Value = vbNullString, vbNullString, Range(wb.Names("notes4")).Value)
        .Add Key:="notes5", Item:=IIf(Range(wb.Names("notes5")).Value = vbNullString, vbNullString, Range(wb.Names("notes5")).Value)
        .Add Key:="notes6", Item:=IIf(Range(wb.Names("notes6")).Value = vbNullString, vbNullString, Range(wb.Names("notes6")).Value)
        .Add Key:="notes7", Item:=IIf(Range(wb.Names("notes7")).Value = vbNullString, vbNullString, Range(wb.Names("notes7")).Value)
        .Add Key:="notes8", Item:=IIf(Range(wb.Names("notes8")).Value = vbNullString, vbNullString, Range(wb.Names("notes8")).Value)
        .Add Key:="roll_length", Item:=IIf(Range(wb.Names("roll_length")).Value = vbNullString, vbNullString, Range(wb.Names("roll_length")).Value)
    End With
    Set AddMoreRbaNames = dict
End Function

Public Function AddRbaNames(dict As Object, wb As Workbook, tag As String, r_start As Long, r_end As Long, c_start As Long, c_end As Long) As Object
    Dim sht As Worksheet
    Dim nr As String
    Dim ret_val As Variant
    Set sht = wb.Sheets("ENG")
    Dim r, c, rw, cl As Long
    For r = r_start To r_end
        cl = 0
        For c = c_start To c_end
            rw = Abs(r_end - r)
            nr = tag & "_" & cl & rw
            ret_val = CreateNamedRange(wb, nr, sht, CLng(r), CLng(c))
            dict.Add nr, IIf(ret_val = vbNullString, vbNullString, ret_val)
            cl = cl + 1
        Next c
    Next r
    Set AddRbaNames = dict
End Function
