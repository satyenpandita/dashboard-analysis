from utils.cell_functions import *


def populate_data(worksheet, row, col):
    kf = dict()
    subkeys = ['net_debt', 'capital_employed', 'leverage', 'net_debt_per_adj_ebidta', 'ebidta_by_capex',
               'ebidta_by_interest', 'roe', 'roce_ebidta_post_tax', 'roce_wacc_country', 'roce_wacc_company',
               'incremental_ebitda_per_capex', 'incremental_ebitda_margin', 'incremental_roce']
    for key in subkeys:
        row += 1
        if key == 'roce_ebidta_post_tax' or key == 'incremental_ebitda_per_capex':
            row += 1
        kf[key] = cell_value(worksheet, row, col)

    return kf


class LeverageAndReturns:

    def __init__(self, worksheet):
        cell_address = find_cell(worksheet, 'Leverage and Returns')
        if cell_address:
            row,col = cell_address
            self.current_year = populate_data(worksheet, row, col+3)
            self.current_year_plus_one = populate_data(worksheet, row, col+4)
            self.current_year_plus_two = populate_data(worksheet, row, col+5)
            self.current_year_plus_three = populate_data(worksheet, row, col+6)
            self.current_year_plus_four = populate_data(worksheet, row, col+7)



