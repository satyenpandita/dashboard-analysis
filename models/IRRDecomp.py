from utils.cell_functions import cell_value


class IRRDecomp(object):
    def __init__(self, workbook):
        self.irr_target = cell_value(workbook, 13, 9)
        self.eps_growth = cell_value(workbook, 14, 9)
        self.irr_yield = cell_value(workbook, 15, 9)
        self.multiple_expansion = cell_value(workbook, 16, 9)
