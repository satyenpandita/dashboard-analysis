import time
from models.BaseModel import BaseModel
from models.CumulativeDashBoard import CumulativeDashBoard


class DashboardArchive(BaseModel):
    def __init__(self, cdsh: CumulativeDashBoard):
        self.stock_code = cdsh.stock_code
        self.base = cdsh.base
        self.bull = cdsh.bull
        self.bear = cdsh.bear
        self.created_at = cdsh.created_at
        self.deleted_at = time.time()