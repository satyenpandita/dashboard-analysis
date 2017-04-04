from models.DashboardV2 import DashboardV2
from models.CumulativeDashBoard import CumulativeDashBoard
from config.mongo_config import db


class DashboardParserV3(object):

    def __init__(self, workbook):
        self.bull = DashboardV2(workbook.sheet_by_index(0))
        self.base = DashboardV2(workbook.sheet_by_index(1))
        self.bear = DashboardV2(workbook.sheet_by_index(2))
        self.stock_code = self.base.stock_code

    def save_dashboard(self):
        cum_dash = CumulativeDashBoard(self.stock_code, self.base, self.bear, self.bull)
        print(cum_dash.to_json())
        db.cumulative_dashboards.insert_one(cum_dash.to_json())