class BullScenario(object):
    def __init__(self, worksheet):
        super(BullScenario, self).__init__()
        self.pt_1year = worksheet.cell(40, 2).value
        self.pt_3year = worksheet.cell(40, 3).value
        self.return_1year = worksheet.cell(40, 4).value
        self.return_3year = worksheet.cell(40, 5).value
        self.prob_1year = worksheet.cell(43, 2).value
        self.prob_3year = worksheet.cell(43, 3).value


class BaseScenario(object):
    def __init__(self, worksheet):
        super(BaseScenario, self).__init__()
        self.pt_1year = worksheet.cell(41, 2).value
        self.pt_3year = worksheet.cell(41, 3).value
        self.return_1year = worksheet.cell(41, 4).value
        self.return_3year = worksheet.cell(41, 5).value
        self.prob_1year = worksheet.cell(44, 2).value
        self.prob_3year = worksheet.cell(44, 3).value
        self.base_case = {
            "1year": {
                "market_cap" : worksheet.cell(40, 7).value,
                "net_cash_per_mcap": worksheet.cell(41, 7).value,
                "enterprise_value": worksheet.cell(42, 7).value,
                "price_earnings_growth": worksheet.cell(43, 7).value,
                "price_to_fcf": worksheet.cell(44, 7).value,
                "adj_fcf_mult": worksheet.cell(45, 7).value
            },
            "3year": {
                "market_cap" : worksheet.cell(40, 8).value,
                "net_cash_per_mcap" : worksheet.cell(41, 8).value,
                "enterprise_value" : worksheet.cell(42, 8).value,
                "price_earnings_growth" : worksheet.cell(43, 8).value,
                "price_to_fcf" : worksheet.cell(44, 8).value,
                "adj_fcf_mult" : worksheet.cell(45, 8).value
            }
        }


class BearScenario(object):
    def __init__(self, worksheet):
        super(BearScenario, self).__init__()
        self.pt_1year = worksheet.cell(42, 2).value
        self.pt_3year = worksheet.cell(42, 3).value
        self.return_1year = worksheet.cell(42, 4).value
        self.return_3year = worksheet.cell(42, 5).value
        self.prob_1year = worksheet.cell(45, 2).value
        self.prob_3year = worksheet.cell(45, 3).value


class TargetPrice(object):
    def __init__(self, worksheet):
        super(TargetPrice, self).__init__()
        self.base = BaseScenario(worksheet).__dict__
        self.bear = BearScenario(worksheet).__dict__
        self.bull = BullScenario(worksheet).__dict__
        self.expected_value_1year = worksheet.cell(46, 4).value
        self.expected_value_3year = worksheet.cell(46, 5).value
        self.borrow_cost_1year = worksheet.cell(47, 4).value
        self.borrow_cost_3year = worksheet.cell(47, 5).value
        self.net_ret_1year = worksheet.cell(48, 4).value
        self.net_ret_3year = worksheet.cell(48, 5).value
