from utils.cell_functions import cell_value_by_key


class FinancialInfo:

    def __init__(self, worksheet):
        super(FinancialInfo, self).__init__()
        self.fd_shares = cell_value_by_key(worksheet, 'FD Shares (m):')
        self.price_per_share = cell_value_by_key(worksheet, 'Price (USD/sh):')
        self.market_cap = None
        if self.fd_shares and self.price_per_share:
            self.market_cap = self.fd_shares*self.price_per_share
        self.cash = cell_value_by_key(worksheet, '  - Cash')
        self.debt = cell_value_by_key(worksheet, '  + Debt')
        self.others_adj = cell_value_by_key(worksheet, '  + Other Adj')
        self.enterprise_value = cell_value_by_key(worksheet, 'EV')
        self.book_value_per_share = cell_value_by_key(worksheet, 'Book Value/shr:')