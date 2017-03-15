from utils.cell_functions import cell_value, cell_value_by_key


class AnalystFillCells(object):
    def __init__(self, data):
        if isinstance(data, dict):
            for k, v in data.items():
                setattr(self, k, v)
        else:
            self.base_model = cell_value_by_key(data, 'Base Model')
            self.best_research = cell_value_by_key(data, 'Best research')
            self.bull_analyst = cell_value_by_key(data, 'Bull Analyst')
            self.bear_analyst = cell_value_by_key(data, 'Bear Analyst')
            self.best_expert_call = cell_value_by_key(data, 'Best expert call')
            self.av_theme = cell_value_by_key(data, 'AV Theme')
            self.sub_theme = cell_value_by_key(data, 'Sub-Theme')
            self.aisa_angle = cell_value_by_key(data, 'Asia angle')
