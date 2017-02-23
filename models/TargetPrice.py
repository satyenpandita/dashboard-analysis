from utils.cell_functions import cell_value, find_cell


class BullScenario(object):
    def __init__(self, worksheet, ref_row, ref_col):
        super(BullScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, ref_row + 1, ref_col + 1)
        self.pt_3year = cell_value(worksheet, ref_row + 1, ref_col + 2)
        self.return_1year = cell_value(worksheet, ref_row + 1, ref_col + 3)
        self.return_3year = cell_value(worksheet, ref_row + 1, ref_col + 4)
        self.prob_1year = cell_value(worksheet, ref_row + 4, ref_col + 1)
        self.prob_3year = cell_value(worksheet, ref_row + 4, ref_col + 2)


class BaseScenario(object):
    def __init__(self, worksheet, ref_row, ref_col):
        super(BaseScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, ref_row + 2, ref_col + 1)
        self.pt_3year = cell_value(worksheet, ref_row + 2, ref_col + 2)
        self.return_1year = abs(cell_value(worksheet, ref_row + 2, ref_col + 3)) \
            if cell_value(worksheet, ref_row + 2, ref_col + 3) else None
        self.return_3year = abs(cell_value(worksheet, ref_row + 2, ref_col + 4)) \
            if cell_value(worksheet, ref_row + 2, ref_col + 4) else None
        self.prob_1year = cell_value(worksheet, ref_row + 5, ref_col + 1)
        self.prob_3year = cell_value(worksheet, ref_row + 5, ref_col + 2)
        self.base_case = {
            "1year": {
                "market_cap" : cell_value(worksheet, ref_row + 1, ref_col + 6),
                "net_cash_per_mcap": cell_value(worksheet, ref_row + 2, ref_col + 6),
                "enterprise_value": cell_value(worksheet, ref_row + 3, ref_col + 6),
                "price_earnings_growth": cell_value(worksheet, ref_row + 4, ref_col + 6),
                "price_to_fcf": cell_value(worksheet, ref_row + 5, ref_col + 6),
                "adj_fcf_mult": cell_value(worksheet, ref_row + 6, ref_col + 6)
            },
            "3year": {
                "market_cap" : cell_value(worksheet, ref_row + 1, ref_col + 7),
                "net_cash_per_mcap": cell_value(worksheet, ref_row + 2, ref_col + 7),
                "enterprise_value": cell_value(worksheet, ref_row + 3, ref_col + 7),
                "price_earnings_growth": cell_value(worksheet, ref_row + 4, ref_col + 7),
                "price_to_fcf": cell_value(worksheet, ref_row + 5, ref_col + 7),
                "adj_fcf_mult": cell_value(worksheet, ref_row + 6, ref_col + 7)
            }
        }


class BearScenario(object):
    def __init__(self, worksheet,  ref_row, ref_col):
        super(BearScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, ref_row + 3, ref_col + 1)
        self.pt_3year = cell_value(worksheet, ref_row + 3, ref_col + 2)
        self.return_1year = cell_value(worksheet, ref_row + 3, ref_col + 3)
        self.return_3year = cell_value(worksheet, ref_row + 3, ref_col + 4)
        self.prob_1year = cell_value(worksheet, ref_row + 6, ref_col + 1)
        self.prob_3year = cell_value(worksheet, ref_row + 6, ref_col + 2)


class TargetPrice(object):
    def __init__(self, worksheet):
        super(TargetPrice, self).__init__()
        row, col = find_cell(worksheet, 'Target Price (TP):')
        self.base = BaseScenario(worksheet, row, col).__dict__
        self.bear = BearScenario(worksheet, row, col).__dict__
        self.bull = BullScenario(worksheet, row, col).__dict__
        self.__validate_and_shuffle()
        self.expected_value_1year = cell_value(worksheet, row + 7, col + 1)
        self.expected_value_3year = cell_value(worksheet, row + 7, col + 2)
        self.borrow_cost_1year = cell_value(worksheet, row + 8, col + 3)
        self.borrow_cost_3year = cell_value(worksheet, row + 8, col + 4)
        self.net_ret_1year = cell_value(worksheet, row + 9, col + 3)
        self.net_ret_3year = cell_value(worksheet, row + 9, col + 4)

    def __validate_and_shuffle(self):
        bear_ret1 = self.bear.get('return_1year')
        bull_ret1 = self.bull.get('return_1year')
        bear_ret3 = self.bear.get('return_3year')
        bull_ret3 = self.bull.get('return_3year')

        if bear_ret1 and bear_ret1 > 0:
            self.bull['return_1year'] = bear_ret1
            self.bull['return_3year'] = bear_ret3
            self.bear['return_1year'] = bull_ret1
            self.bear['return_3year'] = bull_ret3
