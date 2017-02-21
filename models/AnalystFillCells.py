from utils.cell_functions import cell_value, cell_value_by_key


class AnalystFillCells(object):
    def __init__(self, workbook):
        self.base_model = cell_value_by_key(workbook, 'Base Model')
        self.best_research = cell_value_by_key(workbook, 'Best research')
        self.bull_analyst = cell_value_by_key(workbook, 'Bull Analyst')
        self.bear_analyst = cell_value_by_key(workbook, 'Bear Analyst')
        self.best_expert_call = cell_value_by_key(workbook, 'Best expert call')
        self.av_theme = cell_value_by_key(workbook,'AV Theme')
        self.sub_theme = cell_value_by_key(workbook, 'Sub-Theme')
        self.aisa_angle = cell_value_by_key(workbook, 'Asia angle')
