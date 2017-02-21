from utils.cell_functions import cell_value, cell_value_by_key


class ShortMetrics:

    def __init__(self, worksheet):
        self.borrow_cost = cell_value_by_key(worksheet, 'Borrow Cost')
        self.si_mshares = cell_value_by_key(worksheet, 'SI (m shares)')
        self.sir_bberg = cell_value_by_key(worksheet, 'SIR (Bberg)')
        self.sir_calc = cell_value_by_key(worksheet, 'SIR (Calc)')
        self.si_as_of_ff = cell_value_by_key(worksheet, 'SI as of FF')
