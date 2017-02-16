from utils.cell_functions import cell_value


class AnalystFillCells(object):
    def __init__(self, workbook):
        self.base_model = cell_value(workbook, 3, 9)
        self.best_research = cell_value(workbook, 4, 9)
        self.bull_analyst = cell_value(workbook, 5, 9)
        self.bear_analyst = cell_value(workbook, 6, 9)
        self.best_expert_call = cell_value(workbook, 7, 9)
        self.av_theme = cell_value(workbook, 8, 9)
        self.sub_theme = cell_value(workbook, 9, 9)
        self.aisa_angle = cell_value(workbook, 10, 9)
