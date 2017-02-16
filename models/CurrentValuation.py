from utils.cell_functions import *


def get_object(worksheet, rowx, colx):
    dvc = dict()
    keys = ['current_year', 'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three']
    for idx, key in enumerate(keys):
        year_data = dict()
        year_data['aim'] = cell_value(worksheet, rowx+1, colx+1+idx)
        year_data['consensus'] = cell_value(worksheet,rowx + 2, colx+1+idx)
        dvc[key] = year_data
    return dvc


class CurrentValuation:

    def __init__(self, worksheet):
        super(CurrentValuation, self).__init__()
        self.ev_per_gross_revenue = get_object(worksheet, 3, 11)
        self.ev_per_adj_ebidta = get_object(worksheet, 8, 11)
        self.cap_per_adj_eps = get_object(worksheet, 11, 11)
        self.cap_per_others = get_object(worksheet, 14, 11)