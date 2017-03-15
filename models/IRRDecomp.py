from utils.cell_functions import cell_value, cell_value_by_key


class IRRDecomp(object):
    def __init__(self, workbook):
        self.irr_target = cell_value_by_key(workbook, 'IRR Target %')
        self.eps_growth = cell_value_by_key(workbook,'EPS Growth')
        self.irr_yield = cell_value_by_key(workbook, 'Yield')
        self.multiple_expansion = cell_value_by_key(workbook, 'Multiple Expansion')
