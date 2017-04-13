from utils.cell_functions import cell_value_by_key, find_cell, cell_value


class FinancialInfo:

    def __init__(self, worksheet):
        super(FinancialInfo, self).__init__()
        self.fd_shares = cell_value_by_key(worksheet, 'FD Shares (m):')
        cell_address = find_cell(worksheet, "Mkt Cap (M):")
        self.market_cap = None
        self.price_per_share = None
        if cell_address:
            row, col = cell_address
            self.price_per_share = cell_value(worksheet, row - 1, col + 1)
            self.market_cap = cell_value(worksheet, row, col + 1)

        self.cash = cell_value_by_key(worksheet, '  - Cash')
        self.debt = cell_value_by_key(worksheet, '  + Debt')
        self.others_adj = cell_value_by_key(worksheet, '  + Other Adj')
        self.enterprise_value = cell_value_by_key(worksheet, 'EV')
        self.book_value_per_share = cell_value_by_key(worksheet, 'Book Value/shr:')