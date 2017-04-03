from models.DeltaVsConsensusV2 import DeltaVsConsensusV2
from models.TargetPrice import TargetPrice
from models.DataTracking import DataTracking
from models.FinancialInfo import FinancialInfo
from models.CurrentValuationV2 import CurrentValuationV2
from models.LeverageAndReturns import LeverageAndReturns
from models.ShortMetrics import ShortMetrics
from models.AnalystFillCells import AnalystFillCells
from models.IRRDecomp import IRRDecomp
from models.KeyFinancials import KeyFinancials
from models.Tam import Tam
from utils.cell_functions import cell_value, cell_value_by_key, find_cell


def get_likely_outcome(worksheet):
    likely_outcome = dict()
    cell_address = find_cell(worksheet, 'Most likely outcome:')
    if cell_address:
        row, col = cell_address
        likely_outcome['next_1quarter'] = cell_value(worksheet, row + 1, col)
        likely_outcome['next_1year'] = cell_value(worksheet, row + 4, col)
        likely_outcome['next_3year'] = cell_value(worksheet, row + 8, col)
    return likely_outcome


def get_opp_thesis(worksheet):
    opp_thesis = dict()
    cell_address = find_cell(worksheet, 'Opposite thesis/ Inv Risks (Bear for Long, Bull for Short):')
    if cell_address:
        row, col = cell_address
        opp_thesis['inv_risks'] = cell_value(worksheet, row + 1, col)
        opp_thesis['next_opposite_thesis'] = cell_value(worksheet, row + 4, col)
        opp_thesis['next_living_will'] = cell_value(worksheet, row + 7, col)
    return opp_thesis


class DashboardV2(object):
    """docstring for Dashboard"""

    def __init__(self, data):
        if isinstance(data, dict):
            for k, v in data.items():
                setattr(self, k, v)
        else:
            self.company = cell_value_by_key(data, 'Company:')
            self.stock_code = cell_value_by_key(data, 'Stock Code:')
            self.fiscal_year_end = cell_value_by_key(data, 'Fiscal Yea end:')
            self.adto_20days = cell_value_by_key(data, 'ADTO (20 day, US$mn):')
            self.free_float_mshs = cell_value_by_key(data, 'Free Float (m shs/% of FDO):')
            self.free_float_pfdo = cell_value_by_key(data, 'Free Float (m shs/% of FDO):', col_offset=2)
            self.wacc_country = cell_value_by_key(data, 'WACC (Country/ Company):')
            self.wacc_company = cell_value_by_key(data, 'WACC (Country/ Company):', col_offset=2)
            self.direction = cell_value_by_key(data, 'Direction (L/S) & Current Size:')
            self.current_size = cell_value_by_key(data, 'Direction (L/S) & Current Size:', col_offset=2)
            self.scenario = cell_value_by_key(data, 'Scenario & (Base+Bear):')
            self.analyst_primary = cell_value_by_key(data, 'Analyst (Primary/ Secondary):')
            self.analyst_secondary = cell_value_by_key(data, 'Analyst (Primary/ Secondary):', col_offset=2)
            self.size_reco_primary = cell_value_by_key(data, 'Size Reco (Primary/ Secondary):')
            self.size_reco_secondary = cell_value_by_key(data, 'Size Reco (Primary/ Secondary):', col_offset=2)
            self.last_updated = cell_value_by_key(data, 'Last Updated:')
            self.next_earnings = cell_value_by_key(data, 'Next earnings: ')
            self.forecast_period = cell_value_by_key(data, 'Forecast Period:')
            self.likely_outcome = get_likely_outcome(data)
            self.opp_thesis = get_opp_thesis(data)
            self.delta_consensus = DeltaVsConsensusV2(data).__dict__
            self.short_metrics = ShortMetrics(data).__dict__
            self.irr_decomp = IRRDecomp(data).__dict__
            self.tam = Tam(data).__dict__
            self.analyst_fill_cells = AnalystFillCells(data).__dict__
            self.target_price = TargetPrice(data).__dict__
            self.data_tracking = DataTracking(data).__dict__
            self.current_valuation = CurrentValuationV2(data).__dict__
            self.financial_info = FinancialInfo(data).__dict__
            self.leverage_and_returns = LeverageAndReturns(data).__dict__
            self.key_financials = KeyFinancials(data).__dict__
            # self.yoy_growth_revenue = self.calculate_growth('gross_revenue')
            # self.cagr_4years_revenue = self.calculate_cagr_4yrs('gross_revenue')
            # self.yoy_growth_eps = self.calculate_growth('adj_eps')
            # self.cagr_4years_eps = self.calculate_cagr_4yrs('adj_eps')
            # self.base_plus_bear = self.calculate_base_plus_bear()

    def calculate_base_plus_bear(self):
        base = self.target_price.get('base').get('return_1year')
        bear = self.target_price.get('bear').get('return_1year')
        if base and bear:
            return base + bear

    # def calculate_fcf(self):
    #     target_year = self.delta_consensus.get('current_year_plus_two')
    #     cash = self.financial_info.get('cash')
    #     if target_year and cash:
    #         other = target_year.get('others')
    #         if other.get('aim'):
    #             return cash / other.get('aim')
    #     return None
    #
    # def calculate_growth(self, metric):
    #     curr_year = self.delta_consensus.get('current_year')
    #     next_year = self.delta_consensus.get('current_year_plus_one')
    #     if curr_year and next_year:
    #         curr_metric = curr_year.get(metric)
    #         next_metric = next_year.get(metric)
    #         if curr_metric and next_metric:
    #             try:
    #                 change_percent = (next_metric.get('aim') / curr_metric.get('aim')) - 1
    #                 return change_percent
    #             except Exception:
    #                 return None
    #     return None
    #
    # def calculate_cagr_4yrs(self, metric):
    #     curr_year = self.delta_consensus_list.get('current_year')
    #     target_year = self.delta_consensus_list.get('current_year_plus_three')
    #     if curr_year and target_year:
    #         curr_metric = curr_year.get(metric)
    #         target_year = target_year.get(metric)
    #         if curr_metric and target_year:
    #             try:
    #                 ratio = target_year.get('aim') / curr_metric.get('aim')
    #                 if ratio > 0:
    #                     cagr = pow(ratio, .25) - 1
    #                     return cagr
    #                 else:
    #                     return None
    #             except Exception:
    #                 return None
    #     return None

    def direction_char(self):
        if self.direction.lower() == 'short':
            return 'Short'
        elif self.direction.lower() == 'long':
            return 'Long'
        else:
            return ''

    def model_present(self):
        return u'\u2713'

    def dashboard_present(self):
        return u'\u2713'

    def base_1yr_present(self):
        try:
            if self.target_price and self.target_price['base'] and self.target_price['base']['pt_1year']:
                return True
            else:
                return False
        except KeyError:
            return False

    def base_3yr_present(self):
        try:
            if self.target_price and self.target_price['base'] and self.target_price['base']['pt_3year']:
                return True
            else:
                return False
        except KeyError:
            return False

    def bear_1yr_present(self):
        try:
            if self.target_price and self.target_price['bear'] and self.target_price['bear']['pt_1year']:
                return True
            else:
                return False
        except KeyError:
            return False

    def kpi1_present(self):
        print(self.company)
        try:
            if self.data_tracking and self.data_tracking['kpi1'] and self.data_tracking['kpi1']['tracking_metric']:
                return True
            else:
                return False
        except KeyError:
            return False

    def kpi2_present(self):
        try:
            if self.data_tracking and self.data_tracking['kpi2'] and self.data_tracking['kpi2']['tracking_metric']:
                return True
            else:
                return False
        except KeyError:
            return False

    def kpi3_present(self):
        try:
            if self.data_tracking and self.data_tracking['kpi3'] and self.data_tracking['kpi3']['tracking_metric']:
                return True
            else:
                return False
        except KeyError:
            return False

    def get_kpi(self, key):
        return self.data_tracking[key]['tracking_metric']

    # def get_delta_consensus(self, year, metric):
    #     target_year = self.delta_consensus_list.get(year)
    #     if target_year:
    #         target_metric = target_year.get(metric)
    #         if target_metric:
    #             try:
    #                 change_percent = target_metric.get('aim') / target_metric.get('consensus') - 1
    #                 return change_percent
    #             except Exception:
    #                 return None


def get_delta_consensus(worksheet):
    dvc_dict = dict()
    keys = ['cq', 'current_year', 'current_year_plus_one', 'current_year_plus_two',
            'current_year_plus_three']
    cell_address = find_cell(worksheet, 'Delta v/s Consensus')
    if cell_address:
        row, col = cell_address
        colx = col + 2
        for idx, key in enumerate(keys):
            dvc_dict[key] = DeltaVsConsensusV2(worksheet, row, colx + idx).__dict__
            colx += 1
    return dvc_dict
