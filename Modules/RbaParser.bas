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
    
    For r = 19 To 20
        material_number = Cells(r, 1)
        Set json_object = ParseRBA(ThisWorkbook.path & "\RBAs\RBA_" & material_number & ".xlsx")
        json_object("article_code") = material_number
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
    App.Start
    file_path = SelectRBAFile
    'progress_bar = App.gDll.ShowProgressBar(4)
    ' Task 1
    'progress_bar = App.gDll.SetProgressBar(progress_bar, 1, "Task 1/4")
    path_no_ext = Replace(file_path, ".xlsx", vbNullString)
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
        .Add Key:="actual_weft_count", item:=IIf(Range(wb.Names("actual_weft_count")).value = vbNullString, vbNullString, Range(wb.Names("actual_weft_count")).value)
        .Add Key:="article_code", item:=IIf(Range(wb.Names("article_code")).value = vbNullString, vbNullString, Range(wb.Names("article_code")).value)
        .Add Key:="aux_selvedges_closing_degrees", item:=IIf(Range(wb.Names("aux_selvedges_closing_degrees")).value = vbNullString, vbNullString, Range(wb.Names("aux_selvedges_closing_degrees")).value)
        .Add Key:="bottom_rapier_clamps", item:=IIf(Range(wb.Names("bottom_rapier_clamps")).value = vbNullString, vbNullString, Range(wb.Names("bottom_rapier_clamps")).value)
        .Add Key:="bottom_spreader_bars", item:=IIf(Range(wb.Names("bottom_spreader_bars")).value = vbNullString, vbNullString, Range(wb.Names("bottom_spreader_bars")).value)
        .Add Key:="central_selvedges_drawing_in", item:=IIf(Range(wb.Names("central_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_drawing_in")).value)
        .Add Key:="central_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("central_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_ends_per_dent")).value)
        .Add Key:="central_selvedges_number_ends", item:=IIf(Range(wb.Names("central_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_number_ends")).value)
        .Add Key:="central_selvedges_weave", item:=IIf(Range(wb.Names("central_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_weave")).value)
        .Add Key:="central_selvedges_yarn_count", item:=IIf(Range(wb.Names("central_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("central_selvedges_yarn_count")).value)
        .Add Key:="cutting_degrees", item:=IIf(Range(wb.Names("cutting_degrees")).value = vbNullString, vbNullString, Range(wb.Names("cutting_degrees")).value)
        .Add Key:="date", item:=IIf(Range(wb.Names("date")).value = vbNullString, vbNullString, Range(wb.Names("date")).value)
        .Add Key:="dorn_left_selvedges_drawing_in", item:=IIf(Range(wb.Names("dorn_left_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_drawing_in")).value)
        .Add Key:="dorn_left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_ends_per_dent")).value)
        .Add Key:="dorn_left_selvedges_number_ends", item:=IIf(Range(wb.Names("dorn_left_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_number_ends")).value)
        .Add Key:="dorn_left_selvedges_weave", item:=IIf(Range(wb.Names("dorn_left_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_weave")).value)
        .Add Key:="dorn_left_selvedges_yarn_count", item:=IIf(Range(wb.Names("dorn_left_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("dorn_left_selvedges_yarn_count")).value)
        .Add Key:="draw_in_harness", item:=IIf(Range(wb.Names("draw_in_harness")).value = vbNullString, vbNullString, Range(wb.Names("draw_in_harness")).value)
        .Add Key:="draw_in_reed", item:=IIf(Range(wb.Names("draw_in_reed")).value = vbNullString, vbNullString, Range(wb.Names("draw_in_reed")).value)
        .Add Key:="fabric_width", item:=IIf(Range(wb.Names("fabric_width")).value = vbNullString, vbNullString, Range(wb.Names("fabric_width")).value)
        .Add Key:="first_heddle", item:=IIf(Range(wb.Names("first_heddle")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle")).value)
        .Add Key:="first_heddle_1", item:=IIf(Range(wb.Names("first_heddle_1")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle_1")).value)
        .Add Key:="first_heddle_guide", item:=IIf(Range(wb.Names("first_heddle_guide")).value = vbNullString, vbNullString, Range(wb.Names("first_heddle_guide")).value)
        .Add Key:="harness_configuration", item:=IIf(Range(wb.Names("harness_configuration")).value = vbNullString, vbNullString, Range(wb.Names("harness_configuration")).value)
        .Add Key:="horizontal_back_rest_roller", item:=IIf(Range(wb.Names("horizontal_back_rest_roller")).value = vbNullString, vbNullString, Range(wb.Names("horizontal_back_rest_roller")).value)
        .Add Key:="last_heddle", item:=IIf(Range(wb.Names("last_heddle")).value = vbNullString, vbNullString, Range(wb.Names("last_heddle")).value)
        .Add Key:="last_heddle_guide", item:=IIf(Range(wb.Names("last_heddle_guide")).value = vbNullString, vbNullString, Range(wb.Names("last_heddle_guide")).value)
        .Add Key:="left_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_main_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_drawing_in")).value)
        .Add Key:="left_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_main_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_ends_per_dent")).value)
        .Add Key:="left_main_selvedges_number_ends", item:=IIf(Range(wb.Names("left_main_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_number_ends")).value)
        .Add Key:="left_main_selvedges_weave", item:=IIf(Range(wb.Names("left_main_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_weave")).value)
        .Add Key:="left_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_main_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("left_main_selvedges_yarn_count")).value)
        .Add Key:="left_selvedges_drawing_in", item:=IIf(Range(wb.Names("left_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_drawing_in")).value)
        .Add Key:="left_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("left_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_ends_per_dent")).value)
        .Add Key:="left_selvedges_number_ends", item:=IIf(Range(wb.Names("left_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_number_ends")).value)
        .Add Key:="left_selvedges_weave", item:=IIf(Range(wb.Names("left_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_weave")).value)
        .Add Key:="left_selvedges_yarn_count", item:=IIf(Range(wb.Names("left_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("left_selvedges_yarn_count")).value)
        .Add Key:="loom_number", item:=IIf(Range(wb.Names("loom_number")).value = vbNullString, vbNullString, Range(wb.Names("loom_number")).value)
        .Add Key:="loom_type", item:=IIf(Range(wb.Names("loom_type")).value = vbNullString, vbNullString, Range(wb.Names("loom_type")).value)
        .Add Key:="number_ends_wo_selvedges", item:=IIf(Range(wb.Names("number_ends_wo_selvedges")).value = vbNullString, vbNullString, Range(wb.Names("number_ends_wo_selvedges")).value)
        .Add Key:="number_harnesses", item:=IIf(Range(wb.Names("number_harnesses")).value = vbNullString, vbNullString, Range(wb.Names("number_harnesses")).value)
        .Add Key:="pinch_roller_felt_type", item:=IIf(Range(wb.Names("pinch_roller_felt_type")).value = vbNullString, vbNullString, Range(wb.Names("pinch_roller_felt_type")).value)
        .Add Key:="press_roller_type", item:=IIf(Range(wb.Names("press_roller_type")).value = vbNullString, vbNullString, Range(wb.Names("press_roller_type")).value)
        .Add Key:="rba_number", item:=IIf(Range(wb.Names("rba_number")).value = vbNullString, vbNullString, Range(wb.Names("rba_number")).value)
        .Add Key:="reed", item:=IIf(Range(wb.Names("reed")).value = vbNullString, vbNullString, Range(wb.Names("reed")).value)
        .Add Key:="reed_width", item:=IIf(Range(wb.Names("reed_width")).value = vbNullString, vbNullString, Range(wb.Names("reed_width")).value)
        .Add Key:="right_main_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_main_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_drawing_in")).value)
        .Add Key:="right_main_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_main_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_ends_per_dent")).value)
        .Add Key:="right_main_selvedges_number_ends", item:=IIf(Range(wb.Names("right_main_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_number_ends")).value)
        .Add Key:="right_main_selvedges_weave", item:=IIf(Range(wb.Names("right_main_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_weave")).value)
        .Add Key:="right_main_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_main_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("right_main_selvedges_yarn_count")).value)
        .Add Key:="right_selvedges_drawing_in", item:=IIf(Range(wb.Names("right_selvedges_drawing_in")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_drawing_in")).value)
        .Add Key:="right_selvedges_ends_per_dent", item:=IIf(Range(wb.Names("right_selvedges_ends_per_dent")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_ends_per_dent")).value)
        .Add Key:="right_selvedges_number_ends", item:=IIf(Range(wb.Names("right_selvedges_number_ends")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_number_ends")).value)
        .Add Key:="right_selvedges_weave", item:=IIf(Range(wb.Names("right_selvedges_weave")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_weave")).value)
        .Add Key:="right_selvedges_yarn_count", item:=IIf(Range(wb.Names("right_selvedges_yarn_count")).value = vbNullString, vbNullString, Range(wb.Names("right_selvedges_yarn_count")).value)
        .Add Key:="sand_roller_type", item:=IIf(Range(wb.Names("sand_roller_type")).value = vbNullString, vbNullString, Range(wb.Names("sand_roller_type")).value)
        .Add Key:="selvedges_type", item:=IIf(Range(wb.Names("selvedges_type")).value = vbNullString, vbNullString, Range(wb.Names("selvedges_type")).value)
        .Add Key:="shed_closing_degrees", item:=IIf(Range(wb.Names("shed_closing_degrees")).value = vbNullString, vbNullString, Range(wb.Names("shed_closing_degrees")).value)
        .Add Key:="speed", item:=IIf(Range(wb.Names("speed")).value = vbNullString, vbNullString, Range(wb.Names("speed")).value)
        .Add Key:="springs_type", item:=IIf(Range(wb.Names("springs_type")).value = vbNullString, vbNullString, Range(wb.Names("springs_type")).value)
        .Add Key:="style_number", item:=IIf(Range(wb.Names("style_number")).value = vbNullString, vbNullString, Range(wb.Names("style_number")).value)
        .Add Key:="temples_composition", item:=IIf(Range(wb.Names("temples_composition")).value = vbNullString, vbNullString, Range(wb.Names("temples_composition")).value)
        .Add Key:="upper_rapier_clamps", item:=IIf(Range(wb.Names("upper_rapier_clamps")).value = vbNullString, vbNullString, Range(wb.Names("upper_rapier_clamps")).value)
        .Add Key:="upper_spreader_bars", item:=IIf(Range(wb.Names("upper_spreader_bars")).value = vbNullString, vbNullString, Range(wb.Names("upper_spreader_bars")).value)
        .Add Key:="vertical_back_rest_roller", item:=IIf(Range(wb.Names("vertical_back_rest_roller")).value = vbNullString, vbNullString, Range(wb.Names("vertical_back_rest_roller")).value)
        .Add Key:="warp_tension", item:=IIf(Range(wb.Names("warp_tension")).value = vbNullString, vbNullString, Range(wb.Names("warp_tension")).value)
        .Add Key:="weave_pattern", item:=IIf(Range(wb.Names("weave_pattern")).value = vbNullString, vbNullString, Range(wb.Names("weave_pattern")).value)
        .Add Key:="weft_count_set_point", item:=IIf(Range(wb.Names("weft_count_set_point")).value = vbNullString, vbNullString, Range(wb.Names("weft_count_set_point")).value)
        .Add Key:="notes1", item:=IIf(Range(wb.Names("notes1")).value = vbNullString, vbNullString, Range(wb.Names("notes1")).value)
        .Add Key:="notes2", item:=IIf(Range(wb.Names("notes2")).value = vbNullString, vbNullString, Range(wb.Names("notes2")).value)
        .Add Key:="notes3", item:=IIf(Range(wb.Names("notes3")).value = vbNullString, vbNullString, Range(wb.Names("notes3")).value)
        .Add Key:="notes4", item:=IIf(Range(wb.Names("notes4")).value = vbNullString, vbNullString, Range(wb.Names("notes4")).value)
        .Add Key:="notes5", item:=IIf(Range(wb.Names("notes5")).value = vbNullString, vbNullString, Range(wb.Names("notes5")).value)
        .Add Key:="notes6", item:=IIf(Range(wb.Names("notes6")).value = vbNullString, vbNullString, Range(wb.Names("notes6")).value)
        .Add Key:="notes7", item:=IIf(Range(wb.Names("notes7")).value = vbNullString, vbNullString, Range(wb.Names("notes7")).value)
        .Add Key:="notes8", item:=IIf(Range(wb.Names("notes8")).value = vbNullString, vbNullString, Range(wb.Names("notes8")).value)
        .Add Key:="roll_length", item:=IIf(Range(wb.Names("roll_length")).value = vbNullString, vbNullString, Range(wb.Names("roll_length")).value)
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
    Set ws = shtRbaParser
    
    For r = 1 To 19
        material_number = ws.Cells(r, 1)
        Set wb = OpenWorkbook(ThisWorkbook.path & "\RBAs\" & ws.Cells(r, 3) & ".xlsx")
        wb.SaveAs ThisWorkbook.path & "\RBAs\" & material_number & ".xlsx"
        wb.Close
    Next r
End Sub
