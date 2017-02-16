from utils.cell_functions import cell_value


class Tam(object):
    def __init__(self, workbook):
        self.tam_t = cell_value(workbook, 19, 9)
        self.tam_t3 = cell_value(workbook, 20, 9)
        self.cagr = cell_value(workbook, 21, 9)
        self.mkt_share_t = cell_value(workbook, 22, 9)
        self.mkt_share_t3 = cell_value(workbook, 23, 9)
        self.key_comps = [cell_value(workbook, 24, 9), cell_value(workbook, 25, 9), cell_value(workbook, 26, 9)]
