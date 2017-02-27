from utils.cell_functions import *


def get_object(worksheet, row, col):
    kf = dict()
    keys = ['cq_minus_4a', 'cq_minus_1a', 'cq',
            'current_year_minus_four', 'current_year_minus_three', 'current_year_minus_two', 'current_year_minus_one',
            'current_year',
            'current_year_plus_one', 'current_year_plus_two', 'current_year_plus_three', 'current_year_plus_four']
    for idx, key in enumerate(keys):
        kf[key] = cell_value(worksheet, row, col+idx+1)
    return kf


class KeyFinancials(object):

    REV_CHOICES = ['Gross Rev', 'Net Rev', 'Net Interest Income', 'GMV']
    EB_CHOICES = ['Adj. EBITDA', 'EBITDAR', 'EBITA', 'EBIT', 'PPOP']
    FCF_CHOICES = ['Free Cash Flow', 'Core Free Cash Flow']
    INCOME_CHOICES = ['Net Income (GAAP)', 'Adj Net Income']
    CAPEX_CHOICES = ['Total CAPEX', 'Maintenance Capex']

    def __init__(self, worksheet):
        self.gross_rev = None
        self.net_rev = None
        self.net_nii = None
        self.gmv = None
        cell_address = find_cell(worksheet, KeyFinancials.REV_CHOICES, row_offset= 74)
        if cell_address:
            row_rev, col_rev, metric = cell_address
            if metric == KeyFinancials.REV_CHOICES[0]:
                self.gross_rev = get_object(worksheet, row_rev, col_rev)
            elif metric == KeyFinancials.REV_CHOICES[1]:
                self.net_rev = get_object(worksheet, row_rev, col_rev)
            elif metric == KeyFinancials.REV_CHOICES[2]:
                self.net_nii = get_object(worksheet, row_rev, col_rev)
            elif metric == KeyFinancials.REV_CHOICES[3]:
                self.gmv = get_object(worksheet, row_rev, col_rev)

        # EBIT Values
        self.adj_ebitda = None
        self.ebitdar = None
        self.ebita = None
        self.ebit = None
        self.ppop = None
        cell_address = find_cell(worksheet, KeyFinancials.EB_CHOICES, row_offset= 74)
        if cell_address:
            row_ebit, col_ebit, metric = cell_address
            if metric == KeyFinancials.EB_CHOICES[0]:
                self.adj_ebitda = get_object(worksheet, row_ebit, col_ebit)
            elif metric == KeyFinancials.EB_CHOICES[1]:
                self.ebitdar = get_object(worksheet, row_ebit, col_ebit)
            elif metric == KeyFinancials.EB_CHOICES[2]:
                self.ebita = get_object(worksheet, row_ebit, col_ebit)
            elif metric == KeyFinancials.EB_CHOICES[2]:
                self.ebit = get_object(worksheet, row_ebit, col_ebit)
            elif metric == KeyFinancials.EB_CHOICES[2]:
                self.ppop = get_object(worksheet, row_ebit, col_ebit)

        self.opex = None
        cell_address = find_cell(worksheet, 'OPEX', row_offset=74)
        if cell_address:
            row_opex, col_opex = cell_address
            self.opex = get_object(worksheet, row_opex, col_opex)

        self.adj_net_income = None
        self.net_income_gaap = None
        cell_address = find_cell(worksheet, KeyFinancials.INCOME_CHOICES, row_offset=74)
        if cell_address:
            row_income, col_income, metric = cell_address
            if metric == KeyFinancials.INCOME_CHOICES[0]:
                self.net_income_gaap = get_object(worksheet, row_income, col_income)
            elif metric == KeyFinancials.INCOME_CHOICES[1]:
                self.adj_net_income = get_object(worksheet, row_income, col_income)

        self.eps_fully_diluted = None
        cell_address = find_cell(worksheet, 'EPS (fully-diluted)', row_offset=74)
        if cell_address:
            row_eps, col_eps = cell_address
            self.eps_fully_diluted = get_object(worksheet, row_eps, col_eps)

        self.ocf = None
        cell_address = find_cell(worksheet, 'OCF', row_offset=74)
        if cell_address:
            row_ocf, col_ocf = cell_address
            self.ocf = get_object(worksheet, row_ocf, col_ocf)

        self.total_capex = None
        self.maintenance_capex = None
        cell_address = find_cell(worksheet, KeyFinancials.CAPEX_CHOICES, row_offset=74)
        if cell_address:
            row_capex, col_capex, metric = cell_address
            if metric == KeyFinancials.CAPEX_CHOICES[0]:
                self.total_capex = get_object(worksheet, row_capex, col_capex)
            elif metric == KeyFinancials.CAPEX_CHOICES[1]:
                self.maintenance_capex = get_object(worksheet, row_capex, col_capex)

        self.pre_financing_fcf = None
        cell_address = find_cell(worksheet, 'Pre-financing FCF', row_offset=74)
        if cell_address:
            row_ocf, col_ocf = cell_address
            self.pre_financing_fcf = get_object(worksheet, row_ocf, col_ocf)

        self.free_cash_flow = None
        self.core_free_cash_flow = None
        cell_address = find_cell(worksheet, KeyFinancials.FCF_CHOICES, row_offset=74)
        if cell_address:
            row_pcfcf, col_pcfcf, metric = cell_address
            if metric == KeyFinancials.FCF_CHOICES[0]:
                self.free_cash_flow = get_object(worksheet, row_pcfcf, col_pcfcf)
            elif metric == KeyFinancials.FCF_CHOICES[1]:
                self.core_free_cash_flow = get_object(worksheet, row_pcfcf, col_pcfcf)

        self.net_cash = None
        cell_address = find_cell(worksheet, 'Net Cash', row_offset=74)
        if cell_address:
            row_cash, col_cash = cell_address
            self.net_cash = get_object(worksheet, row_cash, col_cash)

        self.total_se_liabilities = None
        cell_address = find_cell(worksheet, 'Total SE and liabilities', row_offset=74)
        if cell_address:
            row_total_se, col_total_se = cell_address
            self.total_se_liabilities = get_object(worksheet, row_total_se, col_total_se)

        self.total_assets = None
        cell_address = find_cell(worksheet, 'Total Assets', row_offset=74)
        if cell_address:
            row_total_assets, col_total_assets = cell_address
            self.total_assets = get_object(worksheet, row_total_assets, col_total_assets)