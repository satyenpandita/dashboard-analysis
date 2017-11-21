from utils.cell_functions import find_cell, cell_value_by_key, cell_value


class QualitativeAssessment(object):

    def __init__(self, worksheet):
        cell_address = find_cell(worksheet, 'Qualitative Assessment', like=True)
        if cell_address:
            row,col = cell_address
            self.macro = cell_value(worksheet, row+2, col)
            self.moat = cell_value(worksheet, row+2, col+1)
            self.mgmt = cell_value(worksheet, row+2, col+2)
            self.corp_gov = cell_value(worksheet, row+2, col+3)
            self.other = cell_value(worksheet, row+2, col+4)
            self.mgmt_dialogue = cell_value(worksheet, row+2, col+6)