from utils.cell_functions import *


def populate_data(worksheet, row, col):
    kf = dict()
    keys = ['current_year', 'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three',
            'current_year_plus_four']
    for idx, key in enumerate(keys):
        kf[key] = cell_value(worksheet, row, col+idx+3)
    return kf


class LeverageAndReturns:

    def __init__(self, worksheet):
        cell_address = find_cell(worksheet, 'Leverage and Returns')
        if cell_address:
            row,col = cell_address
            self.net_debt = populate_data(worksheet, row + 1, col)
            self.capital_employed = populate_data(worksheet, row+2, col)
            self.leverage = populate_data(worksheet, row+3, col)
            self.net_debt_per_adj_ebidta = populate_data(worksheet, row+4, col)
            self.ebidta_by_capex = populate_data(worksheet, row+5, col)
            self.ebidta_by_interest = populate_data(worksheet, row+6, col)
            self.roe = populate_data(worksheet, row+7, col)
            self.roce_ebidta_post_tax = populate_data(worksheet, row+9, col)
            self.roce_wacc_country = populate_data(worksheet, row+10, col)
            self.roce_wacc_company = populate_data(worksheet, row+11, col)
            self.incremental_ebitda_per_capex = populate_data(worksheet, row+13, col)
            self.incremental_ebitda_margin = populate_data(worksheet, row+14, col)
            self.incremental_roce = populate_data(worksheet, row+15, col)



