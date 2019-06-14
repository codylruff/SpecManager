import xlrd, os

def parse(rba_file):

    workbook = xlrd.open_workbook(rba_file)
    ws = workbook.sheet_by_name('ENG')
    dict = {}
    # value = worksheet.cell(row, column)
    dict["actual_weft_count"] = ws.cell(26, 29)
    dict["article_code"] = ws.cell(14, 10)
    dict["aux_selvedges_closing_degrees"] = ws.cell(36, 29)
    dict["bottom_rapier_clamps"] = ws.cell(49, 29)
    dict["bottom_spreader_bars"] = ws.cell(59, 29)
    dict["central_selvedges_drawing_in"] = ws.cell(69, 20)
    dict["central_selvedges_ends_per_dent"] = ws.cell(69, 26)
    dict["central_selvedges_number_ends"] = ws.cell(69, 10)
    dict["central_selvedges_weave"] = ws.cell(69, 32)
    dict["central_selvedges_yarn_count"] = ws.cell(69, 14)
    dict["cutting_degrees"] = ws.cell(38, 10)
    dict["date"] = ws.cell(8, 29)
    dict["dorn_left_selvedges_drawing_in"] = ws.cell(66, 20)
    dict["dorn_left_selvedges_ends_per_dent"] = ws.cell(66, 26)
    dict["dorn_left_selvedges_number_ends"] = ws.cell(66, 10)
    dict["dorn_left_selvedges_weave"] = ws.cell(66, 32)
    dict["dorn_left_selvedges_yarn_count"] = ws.cell(66, 14)
    dict["draw_in_harness"] = ws.cell(18, 29)
    dict["draw_in_reed"] = ws.cell(20, 29)
    dict["fabric_width"] = ws.cell(12, 10)
    dict["first_heddle"] = ws.cell(30, 10)
    dict["first_heddle_1"] = ws.cell(30, 10)
    dict["first_heddle_guide"] = ws.cell(34, 10)
    dict["harness_configuration"] = ws.cell(22, 10)
    dict["horizontal_back_rest_roller"] = ws.cell(42, 10)
    dict["last_heddle"] = ws.cell(30, 29)
    dict["last_heddle_guide"] = ws.cell(34, 29)
    dict["left_main_selvedges_drawing_in"] = ws.cell(67, 20)
    dict["left_main_selvedges_ends_per_dent"] = ws.cell(67, 26)
    dict["left_main_selvedges_number_ends"] = ws.cell(67, 10)
    dict["left_main_selvedges_weave"] = ws.cell(67, 32)
    dict["left_main_selvedges_yarn_count"] = ws.cell(67, 14)
    dict["left_selvedges_drawing_in"] = ws.cell(64, 20)
    dict["left_selvedges_ends_per_dent"] = ws.cell(64, 26)
    dict["left_selvedges_number_ends"] = ws.cell(64, 10)
    dict["left_selvedges_weave"] = ws.cell(64, 32)
    dict["left_selvedges_yarn_count"] = ws.cell(64, 14)
    dict["loom_number"] = ws.cell(10, 29)
    dict["loom_type"] = ws.cell(12, 29)
    dict["number_ends_wo_selvedges"] = ws.cell(24, 10)
    dict["number_harnesses"] = ws.cell(20, 10)
    dict["pinch_roller_felt_type"] = ws.cell(55, 10)
    dict["press_roller_type"] = ws.cell(53, 10)
    dict["rba_number"] = ws.cell(8, 10)
    dict["reed"] = ws.cell(16, 10)
    dict["reed_width"] = ws.cell(16, 29)
    dict["right_main_selvedges_drawing_in"] = ws.cell(68, 20)
    dict["right_main_selvedges_ends_per_dent"] = ws.cell(68, 26)
    dict["right_main_selvedges_number_ends"] = ws.cell(68, 10)
    dict["right_main_selvedges_weave"] = ws.cell(68, 32)
    dict["right_main_selvedges_yarn_count"] = ws.cell(68, 14)
    dict["right_selvedges_drawing_in"] = ws.cell(65, 20)
    dict["right_selvedges_ends_per_dent"] = ws.cell(65, 26)
    dict["right_selvedges_number_ends"] = ws.cell(65, 10)
    dict["right_selvedges_weave"] = ws.cell(65, 32)
    dict["right_selvedges_yarn_count"] = ws.cell(65, 14)
    dict["sand_roller_type"] = ws.cell(53, 29)
    dict["selvedges_type"] = ws.cell(22, 29)
    dict["shed_closing_degrees"] = ws.cell(36, 10)
    dict["speed"] = ws.cell(14, 29)
    dict["springs_type"] = ws.cell(44, 10)
    dict["style_number"] = ws.cell(10, 10)
    dict["temples_composition"] = ws.cell(44, 29)
    dict["upper_rapier_clamps"] = ws.cell(49, 10)
    dict["upper_spreader_bars"] = ws.cell(59, 10)
    dict["vertical_back_rest_roller"] = ws.cell(42, 29)
    dict["warp_tension"] = ws.cell(26, 10)
    dict["weave_pattern"] = ws.cell(18, 10)
    dict["weft_count_set_point"] = ws.cell(24, 29)

    # print values
    for x, y in dict.items():
        print(x, y)

def main():

    dir_path = os.path.dirname(os.path.realpath(__file__))
    for ws_name in os.listdir(dir_path):
        if ws_name.endswith(".xlsx"):
            parse(ws_name)

if __name__ == "__main__":
    main()