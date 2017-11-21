from utils.cell_functions import find_cell, cell_value


def get_object(worksheet, row, col):
    dvc = dict()
    keys = ['pt_bull', 'pt_base', 'pt_bear']
    for idx, key in enumerate(['current_year', 'current_year_plus_three']):
        data = dict()
        for idx1, key1 in enumerate(keys):
            data[key1] = cell_value(worksheet, row + idx1, col + idx)
        dvc[key] = data
    return dvc


class ImpliedMultiple(object):

    REV_CHOICES = ['EV/Gross Rev  - AIM', 'EV/Net Rev  - AIM','EV/Net Interest Income - AIM', 'EV/GMV - AIM']
    EB_CHOICES = ['EV/Adj. EBITDA - AIM', 'EV/EBITDAR - AIM', 'EV/EBITA - AIM', 'EV/EBIT - AIM',
                  'EV/PPOP - AIM', 'P/EV - AIM']
    FCF_CHOICES = ['Free Cash Flow / P - AIM', 'Core Free Cash Flow / P - AIM',
                   'Free Surpus (Insurance)/P - AIM']

    def __init__(self, worksheet):
        cell_address = find_cell(worksheet, 'Implied multiple')
        if cell_address:
            row, col = cell_address
            metric = cell_value(worksheet, 4, 22)
            self.ev_per_gross_revenue = None
            self.ev_per_net_revenue = None
            self.ev_per_nii = None
            self.ev_per_gmv = None
            if metric == ImpliedMultiple.REV_CHOICES[0]:
                self.ev_per_gross_revenue = get_object(worksheet, row + 2, col)
            elif metric == ImpliedMultiple.REV_CHOICES[1]:
                self.ev_per_net_revenue = get_object(worksheet, row + 2, col)
            elif metric == ImpliedMultiple.REV_CHOICES[2]:
                self.ev_per_nii = get_object(worksheet, row + 2, col)
            elif metric == ImpliedMultiple.REV_CHOICES[3]:
                self.ev_per_gmv = get_object(worksheet, row + 2, col)

            self.ev_per_adj_ebidta = None
            self.ev_per_ebidtar =None
            self.ev_per_ebita =None
            self.ev_per_ebit = None
            self.ev_per_ppop = None
            self.p_ev = None

            metric = cell_value(worksheet, 7, 22)
            if metric == ImpliedMultiple.EB_CHOICES[0]:
                self.ev_per_adj_ebidta = get_object(worksheet, row + 5, col)
            elif metric == ImpliedMultiple.EB_CHOICES[1]:
                self.ev_per_ebidtar = get_object(worksheet, row + 5, col)
            elif metric == ImpliedMultiple.EB_CHOICES[2]:
                self.ev_per_ebita = get_object(worksheet, row + 5, col)
            elif metric == ImpliedMultiple.EB_CHOICES[3]:
                self.ev_per_ebit = get_object(worksheet, row + 5, col)
            elif metric == ImpliedMultiple.EB_CHOICES[3]:
                self.ev_per_ppop = get_object(worksheet, row + 5, col)
            elif metric == ImpliedMultiple.EB_CHOICES[4]:
                self.p_ev = get_object(worksheet, row + 5, col)

            self.adj_eps = get_object(worksheet, row+8, col)

            self.fcf_per_p = None
            self.cfcf_per_p = None
            self.free_surpus_price = None
            metric = cell_value(worksheet, 13, 22)
            if metric == ImpliedMultiple.FCF_CHOICES[0]:
                self.fcf_per_p = get_object(worksheet, row + 11, col)
            elif metric == ImpliedMultiple.FCF_CHOICES[1]:
                self.cfcf_per_p = get_object(worksheet, row + 11, col)
            elif metric == ImpliedMultiple.FCF_CHOICES[2]:
                self.cfcf_per_p = get_object(worksheet, row + 11, col)