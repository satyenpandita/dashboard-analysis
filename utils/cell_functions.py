import xlrd


def find_cell(worksheet, val):
    rows, cols = worksheet.nrows, worksheet.ncols
    for row in range(rows):
        for col in range(cols):
            cell = worksheet.cell(row, col)
            if cell.ctype == xlrd.XL_CELL_TEXT:
                cell_val = cell_value(worksheet, row, col)
                if cell_val and cell_val.strip().lower() == val.strip().lower():
                    return row, col


def cell_value(worksheet, rowx, colx):
    cell = worksheet.cell(rowx, colx)
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    else:
        return worksheet.cell(rowx, colx).value


def get_next_column(col_str):
    col = __get_next_column_by_offset(col_str, 1)
    return col


# implement the function for general offset right now just implemented for 1
def __get_next_column_by_offset(col_str, offset):
    if len(col_str) == 1:
        if col_str == 'Z':
            return 'AA'
        else:
            return chr(ord(col_str) + offset)
    elif len(col_str) == 2:
        if __not_last(col_str[-1]):
            return col_str[0] + chr(ord(col_str[-1]) + offset)
        else:
            return chr(ord(col_str[0]) + offset) + 'A'


def __not_last(col_str):
    return col_str != 'Z'