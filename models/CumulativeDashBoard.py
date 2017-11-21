import time
import date_converter
from models.BaseModel import BaseModel
from models.DashboardArchive import DashboardArchive
from config.mongo_config import db


class CumulativeDashBoard(BaseModel):
    def __init__(self, stock_code, base, bull, bear, created_at=time.time()):
        self.stock_code = stock_code
        self.base = base
        self.bull = bull
        self.bear = bear
        self.created_at = created_at

    @classmethod
    def from_dict(cls, obj):
        base = obj['base']
        bull = obj['bull']
        bear = obj['bear']
        stock_code = obj['stock_code']
        created_at = obj['created_at']
        return CumulativeDashBoard(stock_code, base, bull, bear, created_at)

    @classmethod
    def dashboard_publish_report(cls):
        data = []
        docs = db.cumulative_dashboards.find({})
        for doc in docs:
            dash = CumulativeDashBoard.from_dict(doc)
            date_updated = date_converter.timestamp_to_string(dash.created_at, "%d-%m-%Y")
            data.append((dash.stock_code, date_updated))
        return data

    def save(self):
        document = db.cumulative_dashboards.find_one({'stock_code': self.stock_code})
        if document is not None:
            archive = DashboardArchive(CumulativeDashBoard.from_dict(document))
            db.dashboard_archives.insert_one(archive.to_json())
            db.cumulative_dashboards.delete_one({'stock_code': self.stock_code})
        db.cumulative_dashboards.insert_one(self.to_json())
