import time
from models.BaseModel import BaseModel
from models.DashboardArchive import DashboardArchive
from config.mongo_config import db


class CumulativeDashBoard(BaseModel):
    def __init__(self, stock_code, base, bull, bear):
        self.stock_code = stock_code
        self.base = base
        self.bull = bull
        self.bear = bear
        self.created_at = time.time()

    @classmethod
    def from_dict(cls, obj):
        base = obj['base']
        bull = obj['bull']
        bear = obj['bear']
        stock_code = obj['stock_code']
        return CumulativeDashBoard(stock_code, base, bull, bear)

    def save(self):
        document = db.cumulative_dashboards.find_one({'stock_code': self.stock_code})
        if document is not None:
            archive = DashboardArchive(CumulativeDashBoard.from_dict(document))
            db.dashboard_archives.insert_one(archive.to_json())
            db.cumulative_dashboards.delete_one({'stock_code': self.stock_code})
        db.cumulative_dashboards.insert_one(self.to_json())
