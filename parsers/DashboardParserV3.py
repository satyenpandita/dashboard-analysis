from models.DashboardV2 import DashboardV2
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardArchive import DashboardArchive
from utils.diff_email import target_price_diff_ids
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

    def save_dashboard(self, file=None):
        doc_exists = False
        archive_id = None
        cum_dash = CumulativeDashBoard(self.stock_code, self.base, self.bull, self.bear)
        document = db.cumulative_dashboards.find_one({'stock_code': self.stock_code})
        if document is not None:
            doc_exists = True
            archive = DashboardArchive(CumulativeDashBoard.from_dict(document))
            res = db.dashboard_archives.insert_one(archive.to_json())
            archive_id = res.inserted_id
            db.cumulative_dashboards.delete_one({'stock_code': self.stock_code})
        res = db.cumulative_dashboards.insert_one(cum_dash.to_json())
        cum_dash_id = res.inserted_id
        # if doc_exists:
        #     target_price_diff_ids.delay(str(cum_dash_id), str(archive_id), file=file)
