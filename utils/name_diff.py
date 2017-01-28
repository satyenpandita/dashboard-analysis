def get_name_list(workbook, col):
    worksheet = workbook.sheet_by_index(0)
    rows = worksheet.nrows
    name_list = []
    for row in range(rows):
        cell = worksheet.cell(row, col)
        formatting_info = workbook.xf_list[cell.xf_index]
        background_color_index = formatting_info.background.background_colour_index
        if background_color_index == 65:
            name_list.append(cell.value)
    return name_list

