from config.mongo_config import db
from models.obselete.Dashboard import Dashboard


class DashboardParser:

    def __init__(self, worksheet):
        self.dashboard = Dashboard(worksheet)

    def save_dashboard(self):
        print(self.dashboard.__dict__)
        db.dashboards.insert_one(self.dashboard.__dict__)