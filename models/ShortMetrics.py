from utils.cell_functions import cell_value

class ShortMetrics:

    def __init__(self, worksheet):
        self.borrow_cost = cell_value(worksheet, 11, 6)
        self.si_mshares = cell_value(worksheet, 11, 6)
        self.sir_bberg = cell_value(worksheet, 11, 6)
        self.sir_calc = cell_value(worksheet, 11, 6)
        self.si_as_of_ff = cell_value(worksheet, 11, 6)
