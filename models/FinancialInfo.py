from utils.cell_functions import *


class FinancialInfo:

    def __init__(self, worksheet):
        super(FinancialInfo, self).__init__()
        self.fd_shares = cell_value(worksheet, 2, 6)
        self.price_per_share = cell_value(worksheet, 3, 6)
        self.market_cap = self.fd_shares*self.price_per_share
        self.cash = cell_value(worksheet, 5, 6)
        self.debt = cell_value(worksheet, 6, 6)
        self.others_adj = cell_value(worksheet, 7, 6)
        self.enterprise_value = cell_value(worksheet, 8, 6)
        self.book_value_per_share = cell_value(worksheet, 9, 6)