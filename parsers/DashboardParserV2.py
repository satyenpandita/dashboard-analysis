from models.DashboardV2 import DashboardV2
from config.mongo_config import db


class DashboardParserV2:

    def __init__(self, worksheet):
        self.dashboard = DashboardV2(worksheet)

    def save_dashboard(self):
        print(self.dashboard.__dict__)
        db.dashboards_v2.insert_one(self.dashboard.__dict__)