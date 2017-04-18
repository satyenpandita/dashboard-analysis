import xlrd
import functools


def find_cell(worksheet, val, row_fixed=None, row_offset=None, col_fixed=None):
    limit_row, limit_col = None, None
    cell_address = find_cell_wrapped(worksheet, "Total Assets", row_offset=100, col_fixed=1)
    if cell_address:
        limit_row, limit_col = cell_address
    return find_cell_wrapped(worksheet, val, row_fixed, row_offset, col_fixed, limit_row)


def find_cell_wrapped(worksheet, val, row_fixed=None, row_offset=0, col_fixed=0, limit=None):
    if row_fixed:
        return __cell_by_row(worksheet, row_fixed, val)
    elif col_fixed:
        return __cell_by_col(worksheet, col_fixed, val)
    else:
        max_limit = worksheet.nrows if limit is None else limit
        min_limit = row_offset if row_offset is not None else 0
        for row in range(min_limit, max_limit):
            cell = __cell_by_row(worksheet, row, val)
            if cell:
                return cell


def __cell_by_col(worksheet, col, val):
    rows = worksheet.nrows
    for row in range(rows):
        cell = worksheet.cell(row, col)
        if cell.ctype == xlrd.XL_CELL_TEXT:
            cell_val = cell_value(worksheet, row, col)
            if isinstance(val, list):
                if cell_val in val:
                    return row, col, cell_val
            else:
                if cell_val and cell_val.strip().lower() == val.strip().lower():
                    return row, col


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
        val = worksheet.cell(rowx, colx).value
        if val == 'NM':
            return 0
        else:
            if cell.ctype == xlrd.XL_CELL_TEXT:
                if "#N/A" in val:
                    return None
                else:
                    return val.strip()
            else:
                return val


def cell_value_by_key(worksheet, key_str, row_offset=0, col_offset=1):
    cell_address = find_cell(worksheet, key_str)
    if cell_address:
        row, col = cell_address
        return cell_value(worksheet, row + row_offset, col + col_offset)


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
