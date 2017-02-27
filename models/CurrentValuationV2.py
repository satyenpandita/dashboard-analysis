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


class CurrentValuationV2(object):

    REV_CHOICES = ['EV/Gross Rev  - AIM', 'EV/Net Rev  - AIM','EV/Net Interest Income - AIM', 'EV/GMV - AIM']
    EB_CHOICES = ['EV/Adj. EBITDA - AIM', 'EV/EBITDAR - AIM', 'EV/EBITA - AIM', 'EV/EBIT - AIM', 'EV/PPOP - AIM']
    FCF_CHOICES = ['Free Cash Flow / P - Consensus', 'Core Free Cash Flow / P - Consensus']

    def __init__(self, worksheet):
        super(CurrentValuationV2, self).__init__()

        cell_address = find_cell(worksheet, 'Current Valuation')

        if cell_address:
            row, col = cell_address
            self.ev_per_gross_revenue = None
            self.ev_per_net_revenue = None
            self.ev_per_nii = None
            self.ev_per_gmv = None
            metric = cell_value(worksheet, row + 1, col)
            if metric == CurrentValuationV2.REV_CHOICES[0]:
                self.ev_per_gross_revenue = get_object(worksheet, row + 1, col)
            elif metric == CurrentValuationV2.REV_CHOICES[1]:
                self.ev_per_net_revenue = get_object(worksheet, row + 1, col)
            elif metric == CurrentValuationV2.REV_CHOICES[2]:
                self.ev_per_nii = get_object(worksheet, row + 1, col)
            elif metric == CurrentValuationV2.REV_CHOICES[3]:
                self.ev_per_gmv = get_object(worksheet, row + 1, col)

            self.ev_per_adj_ebidta = None
            self.ev_per_ebidtar =None
            self.ev_per_ebita =None
            self.ev_per_ebit = None
            self.ev_per_ppop = None

            metric = cell_value(worksheet, row + 3, col)
            if metric == CurrentValuationV2.EB_CHOICES[0]:
                self.ev_per_adj_ebidta = get_object(worksheet, row + 3, col)
            elif metric == CurrentValuationV2.EB_CHOICES[1]:
                self.ev_per_ebidtar = get_object(worksheet, row + 3, col)
            elif metric == CurrentValuationV2.EB_CHOICES[2]:
                self.ev_per_ebita = get_object(worksheet, row + 3, col)
            elif metric == CurrentValuationV2.EB_CHOICES[3]:
                self.ev_per_ebit = get_object(worksheet, row + 3, col)
            elif metric == CurrentValuationV2.EB_CHOICES[3]:
                self.ev_per_ppop = get_object(worksheet, row + 3, col)

            self.cap_per_adj_eps = get_object(worksheet, row + 7, col)
            self.cap_per_others = get_object(worksheet, row + 9, col)

            self.fcf_per_p = None
            self.cfcf_per_p = None
            metric = cell_value(worksheet, row + 11, col)
            if metric == CurrentValuationV2.FCF_CHOICES[0]:
                self.fcf_per_p = get_object(worksheet, row + 11, col)
            elif metric == CurrentValuationV2.FCF_CHOICES[1]:
                self.cfcf_per_p = get_object(worksheet, row + 11, col)

