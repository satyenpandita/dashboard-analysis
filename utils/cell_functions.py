import xlrd


def find_cell(worksheet, val, row_fixed=None):
    if row_fixed:
        return __cell_by_row(worksheet, row_fixed, val)
    else:
        for row in range(worksheet.nrows):
            return __cell_by_row(worksheet, row, val)


def __cell_by_row(worksheet, row, val):
    cols = worksheet.ncols
    for col in range(cols):
        cell = worksheet.cell(row, col)
        if cell.ctype == xlrd.XL_CELL_TEXT:
            cell_val = cell_value(worksheet, row, col)
            if isinstance(val, list):
                if cell_val in val:
                    return row, col, cell_val
            else:
                if cell_val and cell_val.strip().lower() == val.strip().lower():
                    return row, col

def cell_value(worksheet, rowx, colx):
    cell = worksheet.cell(rowx, colx)
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    else:
        return worksheet.cell(rowx, colx).value


def cell_value_by_key(worksheet, key_str, row_offset=0, col_offset=1):
    cell_address = find_cell(worksheet, key_str)
    if cell_address:
        row, col = cell_address
        return cell_value(row + row_offset, col+col_offset)

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