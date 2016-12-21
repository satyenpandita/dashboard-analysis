from models.Dashboard import Dashboard
import jsonpickle


class DashboardParser:

    def __init__(self, worksheet):
        self.dashboard = Dashboard(worksheet)

    def save_dashboard(self):
        print(jsonpickle.encode(self.dashboard))