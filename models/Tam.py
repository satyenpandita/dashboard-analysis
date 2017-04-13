from utils.cell_functions import cell_value, cell_value_by_key, find_cell


class Tam(object):
    def __init__(self, workbook):
        self.tam_t = cell_value_by_key(workbook, 'TAM(t)')
        self.tam_t3 = cell_value_by_key(workbook, 'TAM(t+3)')
        self.cagr = cell_value_by_key(workbook, 'Cagr')
        self.mkt_share_t = cell_value_by_key(workbook,'Mkt Share (t)')
        self.mkt_share_t3 = cell_value_by_key(workbook, 'Mkt Share (t+3)')
        self.key_comps = ["", "", ""]
        cell_address = find_cell(workbook, 'Key Comps')
        if cell_address:
            row, col = cell_address
            self.key_comps = [cell_value(workbook, row, col+1), cell_value(workbook, row+1, col+1), cell_value(workbook, row+2, col+1)]
