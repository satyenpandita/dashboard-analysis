from models.DeltaVsConsensus import DeltaVsConsensus


class Dashboard(object):
    """docstring for Dashboard"""

    def __init__(self, worksheet):
        super(Dashboard, self).__init__()
        self.worksheet = worksheet
        self.company = cell_value(worksheet, 2, 2)
        self.stock_code = cell_value(worksheet, 3, 2)
        self.fiscal_year_end = cell_value(worksheet, 4, 2)
        self.adto_20days = cell_value(worksheet, 7, 2)
        self.free_float_mshs = cell_value(worksheet, 8, 2)
        self.free_float_pfdo = cell_value(worksheet, 8, 3)
        self.wacc_country = cell_value(worksheet, 9, 2)
        self.wacc_company = cell_value(worksheet, 9, 3)
        self.direction = cell_value(worksheet, 10, 2)
        self.current_size = cell_value(worksheet, 10, 3)
        self.scenario = cell_value(worksheet, 11, 2)
        self.base_plus_bear = cell_value(worksheet, 11, 3)
        self.analyst_primary = cell_value(worksheet, 12, 2)
        self.analyst_secondary = cell_value(worksheet, 12, 3)
        self.size_reco_primary = cell_value(worksheet, 13, 2)
        self.size_reco_secondary = cell_value(worksheet, 13, 3)
        self.last_updated = cell_value(worksheet, 14, 2)
        self.next_earnings = cell_value(worksheet, 15, 2)
        self.forecast_period = cell_value(worksheet, 16, 2)
        self.delta_consensus_list = get_consensus_list(worksheet)

        
def cell_value(worksheet, rowx, colx):
    return worksheet.cell(rowx, colx).value


def get_consensus_list(worksheet):
    dvc_list = list()
    keys = ['current_quarter', 'current_year','current_year_plus_one', 'current_year_plus_two',
            'current_year_plus_three']
    for key in keys:
        dvc_dict = dict()
        colx = 3
        dvc_dict[key] = DeltaVsConsensus(worksheet, colx)
        dvc_list.append(dvc_dict)
    return dvc_list
