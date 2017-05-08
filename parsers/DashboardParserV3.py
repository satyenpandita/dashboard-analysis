from models.DashboardV2 import DashboardV2
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardArchive import DashboardArchive
from config.mongo_config import db


class DashboardParserV3(object):

    def __init__(self, workbook):
        self.base = None
        self.bear = None
        self.bull = None
        for sheet in workbook.sheets():
            if "bull" in sheet.name.lower():
                self.bull = DashboardV2(sheet)
            elif "base" in sheet.name.lower():
                self.base = DashboardV2(sheet)
            elif "bear" in sheet.name.lower():
                self.bear = DashboardV2(sheet)
        self.stock_code = self.base.stock_code

    def save_dashboard(self):
        cum_dash = CumulativeDashBoard(self.stock_code, self.base, self.bull, self.bear)
        document = db.cumulative_dashboards.find_one({'stock_code': self.stock_code})
        if document is not None:
            archive = DashboardArchive(CumulativeDashBoard.from_dict(document))
            db.dashboard_archives.insert_one(archive.to_json())
            db.cumulative_dashboards.delete_one({'stock_code': self.stock_code})
        db.cumulative_dashboards.insert_one(cum_dash.to_json())