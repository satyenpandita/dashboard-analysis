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


class CurrentValuation(object):

    REV_CHOICES = ['EV/Gross Rev  - AIM', 'EV/Net Rev  - AIM','EV/Net Interest Income - AIM', 'EV/GMV - AIM']
    EB_CHOICES = ['EV/Adj. EBITDA - AIM', 'EV/EBITDAR - AIM', 'EV/EBITA - AIM', 'EV/EBIT - AIM', 'EV/PPOP - AIM']
    FCF_CHOICES = ['Free Cash Flow / P - Consensus', 'Core Free Cash Flow / P - Consensus']

    def __init__(self, worksheet):
        super(CurrentValuation, self).__init__()
        self.ev_per_gross_revenue = None
        self.ev_per_net_revenue = None
        self.ev_per_nii = None
        self.ev_per_gmv = None
        metric = cell_value(worksheet, 3, 11)
        if metric == CurrentValuation.REV_CHOICES[0]:
            self.ev_per_gross_revenue = get_object(worksheet, 3, 11)
        elif metric == CurrentValuation.REV_CHOICES[1]:
            self.ev_per_net_revenue = get_object(worksheet, 3, 11)
        elif metric == CurrentValuation.REV_CHOICES[2]:
            self.ev_per_nii = get_object(worksheet, 3, 11)
        elif metric == CurrentValuation.REV_CHOICES[3]:
            self.ev_per_gmv = get_object(worksheet, 3, 11)

        self.ev_per_adj_ebidta = None
        self.ev_per_ebidtar =None
        self.ev_per_ebita =None
        self.ev_per_ebit = None
        self.ev_per_ppop = None

        metric = cell_value(worksheet, 5, 11)
        if metric == CurrentValuation.EB_CHOICES[0]:
            self.ev_per_adj_ebidta = get_object(worksheet, 5, 11)
        elif metric == CurrentValuation.EB_CHOICES[1]:
            self.ev_per_ebidtar = get_object(worksheet, 5, 11)
        elif metric == CurrentValuation.EB_CHOICES[2]:
            self.ev_per_ebita = get_object(worksheet, 5, 11)
        elif metric == CurrentValuation.EB_CHOICES[3]:
            self.ev_per_ebit = get_object(worksheet, 5, 11)
        elif metric == CurrentValuation.EB_CHOICES[3]:
            self.ev_per_ppop = get_object(worksheet, 5, 11)

        self.cap_per_adj_eps = get_object(worksheet, 9, 11)
        self.cap_per_others = get_object(worksheet, 11, 11)

        self.fcf_per_p = None
        self.cfcf_per_p = None
        metric = cell_value(worksheet, 13, 11)
        if metric == CurrentValuation.FCF_CHOICES[0]:
            self.fcf_per_p = get_object(worksheet, 13, 11)
        elif metric == CurrentValuation.FCF_CHOICES[1]:
            self.cfcf_per_p = get_object(worksheet, 13, 11)

