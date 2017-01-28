from utils.cell_functions import cell_value


class BullScenario(object):
    def __init__(self, worksheet):
        super(BullScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, 40, 2)
        self.pt_3year = cell_value(worksheet, 40, 3)
        self.return_1year = cell_value(worksheet, 40, 4)
        self.return_3year = cell_value(worksheet, 40, 5)
        self.prob_1year = cell_value(worksheet, 43, 2)
        self.prob_3year = cell_value(worksheet, 43, 3)


class BaseScenario(object):
    def __init__(self, worksheet):
        super(BaseScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, 41, 2)
        self.pt_3year = cell_value(worksheet, 41, 3)
        self.return_1year = abs(cell_value(worksheet, 41, 4)) if cell_value(worksheet, 41, 4) else None
        self.return_3year = abs(cell_value(worksheet, 41, 5)) if cell_value(worksheet, 41, 5) else None
        self.prob_1year = cell_value(worksheet, 44, 2)
        self.prob_3year = cell_value(worksheet, 44, 3)
        self.base_case = {
            "1year": {
                "market_cap" : cell_value(worksheet, 40, 7),
                "net_cash_per_mcap": cell_value(worksheet, 41, 7),
                "enterprise_value": cell_value(worksheet, 42, 7),
                "price_earnings_growth": cell_value(worksheet, 43, 7),
                "price_to_fcf": cell_value(worksheet, 44, 7),
                "adj_fcf_mult": cell_value(worksheet, 45, 7)
            },
            "3year": {
                "market_cap" : cell_value(worksheet, 40, 8),
                "net_cash_per_mcap" : cell_value(worksheet, 41, 8),
                "enterprise_value" : cell_value(worksheet, 42, 8),
                "price_earnings_growth" : cell_value(worksheet, 43, 8),
                "price_to_fcf" : cell_value(worksheet, 44, 8),
                "adj_fcf_mult" : cell_value(worksheet, 45, 8)
            }
        }


class BearScenario(object):
    def __init__(self, worksheet):
        super(BearScenario, self).__init__()
        self.pt_1year = cell_value(worksheet, 42, 2)
        self.pt_3year = cell_value(worksheet, 42, 3)
        self.return_1year = cell_value(worksheet, 42, 4)
        self.return_3year = cell_value(worksheet, 42, 5)
        self.prob_1year = cell_value(worksheet, 45, 2)
        self.prob_3year = cell_value(worksheet, 45, 3)


class TargetPrice(object):
    def __init__(self, worksheet):
        super(TargetPrice, self).__init__()
        self.base = BaseScenario(worksheet).__dict__
        self.bear = BearScenario(worksheet).__dict__
        self.bull = BullScenario(worksheet).__dict__
        self.__validate_and_shuffle()
        self.expected_value_1year = cell_value(worksheet, 46, 4)
        self.expected_value_3year = cell_value(worksheet, 46, 5)
        self.borrow_cost_1year = cell_value(worksheet, 47, 4)
        self.borrow_cost_3year = cell_value(worksheet, 47, 5)
        self.net_ret_1year = cell_value(worksheet, 48, 4)
        self.net_ret_3year = cell_value(worksheet, 48, 5)

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
