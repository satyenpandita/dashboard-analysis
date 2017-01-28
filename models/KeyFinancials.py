from utils.cell_functions import *


def populate_data(worksheet, row, col):
    kf = dict()
    keys = ['cq_minus_4a', 'cq_minus_1a', 'cq',
            'current_year_minus_four', 'current_year_minus_four', 'current_year_minus_four', 'current_year_minus_four',
            'current_year'
            'current_year_minus_four', 'current_year_minus_four', 'current_year_minus_four', 'current_year_minus_four']
    for key in keys:
        kf[key] = cell_value(worksheet, row, col+1)
    return kf


class KeyFinancials:

    def __init__(self, worksheet):
        rows, col = worksheet.nrows, 2
        for row in range(rows):
            cell_val = cell_value(worksheet, row, col)
            clean_value = cell_val.strip().lower()
            if clean_value == 'gross rev':
                self.gross_rev = populate_data(worksheet, row, col)
            elif 'net rev' in clean_value or 'gross profit' in clean_value:
                self.net_revenue = populate_data(worksheet, row, col)
            elif 'opex' in clean_value:
                self.opex = populate_data(worksheet, row, col)
            elif 'adj' in clean_value and 'ebitda' in clean_value:
                self.adj_ebitda = populate_data(worksheet, row, col)