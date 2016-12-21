from models.Dashboard import Dashboard


class DashboardParser:

    def __init__(self, worksheet):
        self.dashboard = Dashboard(worksheet)

    def save_dashboard(self):
        print(self.dashboard.__dict__)