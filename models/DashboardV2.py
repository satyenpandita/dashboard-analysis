from models.DeltaVsConsensusV2 import DeltaVsConsensusV2
from models.TargetPrice import TargetPrice
from models.DataTracking import DataTracking
from models.FinancialInfo import FinancialInfo
from models.CurrentValuationV2 import CurrentValuationV2
from models.LeverageAndReturns import LeverageAndReturns
from models.ShortMetrics import ShortMetrics
from models.AnalystFillCells import AnalystFillCells
from models.IRRDecomp import IRRDecomp
from models.Tam import Tam
from utils.cell_functions import cell_value


class DashboardV2(object):
    """docstring for Dashboard"""

    def __init__(self, data):
        if isinstance(data, dict):
            for k, v in data.items():
                setattr(self, k, v)
        else:
            self.company = cell_value(data, 2, 2)
            self.stock_code = cell_value(data, 3, 2)
            self.fiscal_year_end = cell_value(data, 4, 2)
            self.adto_20days = cell_value(data, 7, 2)
            self.free_float_mshs = cell_value(data, 8, 2)
            self.free_float_pfdo = cell_value(data, 8, 3)
            self.wacc_country = cell_value(data, 9, 2)
            self.wacc_company = cell_value(data, 9, 3)
            self.direction = cell_value(data, 10, 2)
            self.current_size = cell_value(data, 10, 3)
            self.scenario = cell_value(data, 11, 2)
            self.analyst_primary = cell_value(data, 12, 2)
            self.analyst_secondary = cell_value(data, 12, 3)
            self.size_reco_primary = cell_value(data, 13, 2)
            self.size_reco_secondary = cell_value(data, 13, 3)
            self.last_updated = cell_value(data, 14, 2)
            self.next_earnings = cell_value(data, 15, 2)
            self.forecast_period = cell_value(data, 16, 2)
            self.likely_outcome = cell_value(data, 21, 11)
            self.delta_consensus_list = get_consensus_list(data)
            self.short_metrics = ShortMetrics(data).__dict__
            self.irr_decomp = IRRDecomp(data).__dict__
            self.tam = Tam(data).__dict__
            self.analyst_fill_cells = AnalystFillCells(data).__dict__
            self.target_price = TargetPrice(data).__dict__
            self.data_tracking = DataTracking(data).__dict__
            self.current_valuation = CurrentValuationV2(data).__dict__
            self.financial_info = FinancialInfo(data).__dict__
            self.leverage_and_returns = LeverageAndReturns(data).__dict__
            self.yoy_growth_revenue = self.calculate_growth('gross_revenue')
            self.cagr_4years_revenue = self.calculate_cagr_4yrs('gross_revenue')
            self.yoy_growth_eps = self.calculate_growth('adj_eps')
            self.cagr_4years_eps = self.calculate_cagr_4yrs('adj_eps')
            self.base_plus_bear = self.calculate_base_plus_bear()

    def calculate_base_plus_bear(self):
        base = self.target_price.get('base').get('return_1year')
        bear = self.target_price.get('bear').get('return_1year')
        if base and bear:
            return base + bear

    def calculate_fcf(self):
        target_year = self.delta_consensus_list.get('current_year_plus_two')
        cash = self.financial_info.get('cash')
        if target_year and cash:
            other = target_year.get('others')
            if other.get('aim'):
                return cash/other.get('aim')
        return None

    def calculate_growth(self, metric):
        curr_year = self.delta_consensus_list.get('current_year')
        next_year = self.delta_consensus_list.get('current_year_plus_one')
        if curr_year and next_year:
            curr_metric = curr_year.get(metric)
            next_metric = next_year.get(metric)
            if curr_metric and next_metric:
                try:
                    change_percent = (next_metric.get('aim')/curr_metric.get('aim')) - 1
                    return change_percent
                except Exception:
                    return None
        return None

    def calculate_cagr_4yrs(self, metric):
        curr_year = self.delta_consensus_list.get('current_year')
        target_year = self.delta_consensus_list.get('current_year_plus_three')
        if curr_year and target_year:
            curr_metric = curr_year.get(metric)
            target_year = target_year.get(metric)
            if curr_metric and target_year:
                try:
                    ratio = target_year.get('aim')/curr_metric.get('aim')
                    if ratio > 0:
                        cagr = pow(ratio, .25) - 1
                        return cagr
                    else:
                        return None
                except Exception:
                    return None
        return None

    def direction_char(self):
        if self.direction.lower() == 'short':
            return 'S'
        elif self.direction.lower() == 'long':
            return 'L'
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

    def get_delta_consensus(self, year, metric):
        target_year = self.delta_consensus_list.get(year)
        if target_year:
            target_metric = target_year.get(metric)
            if target_metric:
                try:
                    change_percent = target_metric.get('aim')/target_metric.get('consensus') - 1
                    return change_percent
                except Exception:
                    return None


def get_consensus_list(worksheet):
    dvc_dict = dict()
    keys = ['current_quarter', 'current_year', 'current_year_plus_one', 'current_year_plus_two',
            'current_year_plus_three']
    colx = 3
    for key in keys:
        dvc_dict[key] = DeltaVsConsensusV2(worksheet, colx).__dict__
        colx += 1
    return dvc_dict
